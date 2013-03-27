<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim rs, sql
If Action = "Check" Then
    Call CheckUser
Else
    Call main
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Call CloseConn

Sub main()
%>
<html>
<head>
<title>注册用户登录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Skin/DefaultSkin.css" rel="stylesheet" type="text/css">
<script language=javascript>
Function SetFocus()
{
if (document.Login.UserName.value=="")
    document.Login.UserName.focus();
Else
    document.Login.UserName.select();
}
Function CheckForm()
{
    if(document.Login.UserName.value=="")
    {
        alert("请输入用户名！");
        document.Login.UserName.focus();
        return false;
    }
    if(document.Login.Password.value == "")
    {
        alert("请输入密码！");
        document.Login.Password.focus();
        return false;
    }
    if(document.Login.CheckNum.value == "")
    {
        alert("请输入验证码！");
        document.Login.CheckNum.focus();
        return false;
    }
}
</script>
</head>
<body onLoad="SetFocus();" leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<form name="Login" action="User_RegCheck.asp" method="post" onSubmit="return CheckForm();">
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="center_tdbgall">
    <tr>
      <td>
  <br>
    <table width="400" border="0" align="center" cellpadding="5" cellspacing="0" class="border" >
      <tr class="title">
        
      <td colspan="2" align="center"> <strong>注册用户认证</strong></td>
      </tr>
      
    <tr>
      <td height="120" colspan="2" class="tdbg">请输入您注册时填写的用户名和密码，以及本站发给你的确认信中的随机验证码。必须完全正确后，你的帐户才会激活。
        <table width="250" border="0" cellspacing="8" cellpadding="0" align="center">
          <tr>
            <td align="right">用户名称：</td>
            <td><input name="UserName"  type="text"  id="UserName" size="23" maxlength="20"></td>
          </tr>
          <tr>
            <td align="right">用户密码：</td>
            <td><input name="Password"  type="password" id="Password" size="23" maxlength="20"></td>
          </tr>
          <tr>
            <td height='25' align='right'>随机验证码：</td>
            <td height='25'><input name="CheckNum" type="text" id="CheckNum" size="23" maxlength="6"></td>
          </tr>
          <tr align="center">
            <td colspan="2"> <input name="Action" type="hidden" id="Action" value="Check">
              <input   type="submit" name="Submit" value=" 确认 "> &nbsp; <input name="reset" type="reset"  id="reset" value=" 清除 ">
            </td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
      <br></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
End Sub

Sub CheckUser()
    Dim password, CheckNum, trs
    UserName = UserNamefilter(Trim(Request("username")))
    password = ReplaceBadChar(Trim(Request("password")))
    CheckNum = ReplaceBadChar(Trim(Request("CheckNum")))

    If UserName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>用户名不能为空！</li>"
    End If
    If password = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>密码不能为空！</li>"
    End If
    If CheckNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>验证码不能为空！</li>"
    End If

    If FoundErr = True Then
        Exit Sub
    End If
    
    password = MD5(password, 16)
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from PE_User where IsLocked=" & PE_False & " and UserName='" & UserName & "' and UserPassword='" & password & "'"
    rs.open sql, Conn, 1, 3
    If rs.bof And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>用户名或密码错误！！！</li>"
    Else
        If password <> rs("UserPassword") Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>用户名或密码错误！！！</li>"
        Else
            If Trim(rs("CheckNum")) <> Trim(CheckNum) Then
                FoundErr = True
                ErrMsg = ErrMsg & "<br><li>验证码不对!</li>"
            Else
                If AdminCheckReg = True Then
                    Set trs = Conn.Execute("select GroupID,GroupSetting from PE_UserGroup where GroupType=1")
                    Call WriteSuccessMsg("恭喜你通过了Email验证。请等待管理开通你的帐号。开通后，你就正式正为本站的一员了。", "../User/")
                Else
                    Set trs = Conn.Execute("select GroupID,GroupSetting from PE_UserGroup where GroupType=2")
                    Call WriteSuccessMsg("恭喜你正式成为本站的一员，请返回首页登录。", "../User/")
                    'Response.Cookies(Site_Sn)("UserName") = rs("UserName")
                    'Response.Cookies(Site_Sn)("Password") = rs("UserPassword")
                    'Response.Cookies(Site_Sn)("LastPassword") = rs("LastPassword")
                    'Response.Cookies(Site_Sn)("CookieDate") = 0
                End If
                GroupID = trs(0)
                Dim GroupSetting
                GroupSetting = Split(trs(1), ",")
                Set trs = Nothing
                rs("GroupID") = GroupID
                rs.Update
            End If
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub

'**************************************************
'函数名：UserNamefilter(
'作  用：过滤用户名(增强过滤,用户名现用于建立个人文集目录)
'**************************************************
Function UserNamefilter(strChar)
    If strChar = "" Or IsNull(strChar) Then
        UserNamefilter = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ",*,|,"""
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    UserNamefilter = Replace(Replace(Replace(Replace(LCase(tempChar), "cdx", ""), "cer", ""), "asp", ""), "asa", "")
End Function
%>
