<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="../API/API_Config.asp"-->
<!--#include file="../API/API_Function.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim sUserName, rsGetPassword
Dim Answer, Password, PwdConfirm, CheckCode

Sub Execute()
    sUserName = Trim(Request("UserName"))
    Answer = Trim(Request("Answer"))
    Password = Trim(Request("Password"))
    PwdConfirm = Trim(Request("PwdConfirm"))
    Select Case Action
    Case "step2"
        Call Step2
    Case "step3"
        Call Step3
    Case "step4"
        Call Step4
    Case Else
        Call Step1
    End Select
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
End Sub

Sub Step1()
    Response.Write "<br>" & vbCrLf
    Response.Write "<table align='center' width='300' border='0' cellpadding='4' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height='15' colspan='2' class='title'>找回密码 &gt;&gt; 第一步：输入用户名</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height='100' colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <form name='form1' method='post' action=''>" & vbCrLf
    Response.Write "        <strong> 请输入你的用户名：</strong>" & vbCrLf
    Response.Write "        <input name='UserName' type='text' id='UserName' size='20' maxlength='20'><br><br>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='step2'>" & vbCrLf
    Response.Write "        <input name='Next' type='submit' id='Next' style='cursor:hand;' value='下一步'>" & vbCrLf
    Response.Write "        <input name='Cancel' type='button' id='Cancel' style='cursor:hand;' onclick='window.location.href=""../index.asp""' value=' 取消 '>" & vbCrLf
    Response.Write "      </form>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub

Sub Step2()
    sUserName = ReplaceBadChar(sUserName)
    Set rsGetPassword = server.CreateObject("adodb.recordset")
    rsGetPassword.open "select UserName,Question,Answer,UserPassword from PE_User where UserName='" & sUserName & "'", Conn, 1, 1
    If rsGetPassword.bof And rsGetPassword.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>对不起，你输入的用户名不存在！</li>"
        rsGetPassword.Close
        Set rsGetPassword = Nothing
        Exit Sub
    End If

    Response.Write "<script language=javascript>" & vbCrLf
    Response.Write "<!--" & vbCrLf
    Response.Write "function refreshimg(){document.all.checkcode.src='../Inc/CheckCode.asp?'+Math.random();}" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<table align='center' width='300' border='0' cellpadding='4' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height='15' colspan='2' class='title'>找回密码 &gt;&gt; 第二步：回答问题</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height='100' colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <form name='form1' method='post' action=''>" & vbCrLf
    Response.Write "        <table width='100%' border='0' cellspacing='5' cellpadding='0'>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td width='44%' align='right'><strong>密码提示问题：</strong></td>" & vbCrLf
    Response.Write "            <td width='56%'>" & rsGetPassword("Question") & "</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td align='right'><strong>你的答案：</strong></td>" & vbCrLf
    Response.Write "            <td><input name='Answer' type='text' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td align='right'><strong>验证码：</strong></td>" & vbCrLf
    Response.Write "            <td><input name='CheckCode' type='text' size='6' maxlength='6'> <a href='javascript:refreshimg()' title='看不清楚，换个图片'><img id='checkcode' src='../inc/checkcode.asp' style='border: 1px solid #ffffff'></a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
    Response.Write "        <br>" & vbCrLf
    Response.Write "        <input name='UserName' type='hidden' id='UserName' value='" & rsGetPassword("UserName") & "'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='step3'>" & vbCrLf
    Response.Write "        <input name='PrevStep' type='button' id='PrevStep' value='上一步' style='cursor:hand;' onclick='history.go(-1)'>&nbsp;" & vbCrLf
    Response.Write "        <input name='NextStep' type='submit' id='NextStep' style='cursor:hand;' value='下一步'>&nbsp;" & vbCrLf
    Response.Write "        <input name='Cancel' type='button' id='Cancel' style='cursor:hand;' onclick='window.location.href=""../index.asp""' value=' 取消 '>" & vbCrLf
    Response.Write "      </form>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    rsGetPassword.Close
    Set rsGetPassword = Nothing
End Sub

Sub Step3()
    If Answer = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请输入提示问题的答案！</li>"
        Exit Sub
    End If
    CheckCode = LCase(ReplaceBadChar(Trim(Request("CheckCode"))))
    If CheckCode = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>验证码不能为空！</li>"
    End If
    If Trim(Session("CheckCode")) = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>你登录时间过长，请重新返回登录页面进行登录。</li>"
    End If
    If CheckCode <> Session("CheckCode") Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>您输入的确认码和系统产生的不一致，请重新输入。</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    sUserName = ReplaceBadChar(sUserName)
    Set rsGetPassword = server.CreateObject("adodb.recordset")
    rsGetPassword.open "select UserName,Question,Answer,UserPassword from PE_User where UserName='" & sUserName & "'", Conn, 1, 1
    If rsGetPassword.bof And rsGetPassword.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>对不起，用户名不存在！可能已经被管理员删除了。</li>"
    Else
        If rsGetPassword("Answer") <> MD5(Answer, 16) Then
            '对动网加密结果的兼容处理
            If rsGetPassword("Answer") <> MD5(Answer, 16) Then
                FoundErr = True
                ErrMsg = ErrMsg & "<br><li>对不起，您的答案不对！</li>"
            End If
        End If
    End If
    If FoundErr = True Then
        rsGetPassword.Close
        Set rsGetPassword = Nothing
        Exit Sub
    End If

    Response.Write "<br>" & vbCrLf
    Response.Write "<table align='center' width='300' border='0' cellpadding='4' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height='15' colspan='2' class='title'>找回密码 &gt;&gt; 第三步：设置新密码</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height='100' colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <form name='form1' method='post' action=''>" & vbCrLf
    Response.Write "        <table width='100%' border='0' cellspacing='5' cellpadding='0'>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td width='44%' align='right'><strong>密码提示问题：</strong></td>" & vbCrLf
    Response.Write "            <td width='56%'>" & rsGetPassword("Question") & "</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td align='right'><strong>你的答案：</strong></td>" & vbCrLf
    Response.Write "            <td>" & Answer & " <input name='Answer' type='hidden' id='Answer' value='" & Answer & "'></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td align='right'><strong>新密码：</strong></td>" & vbCrLf
    Response.Write "            <td><input name='Password' type='password' id='Password' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td align='right'><strong>确认新密码：</strong></td>" & vbCrLf
    Response.Write "            <td><input name='PwdConfirm' type='password' id='PwdConfirm' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
    Response.Write "        <br>" & vbCrLf
    Response.Write "        <input name='UserName' type='hidden' id='UserName' value='" & rsGetPassword("Username") & "'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='step4'>" & vbCrLf
    Response.Write "        <input name='PrevStep' type='button' id='PrevStep' value='上一步' style='cursor:hand;' onclick='history.go(-1)'>&nbsp;" & vbCrLf
    Response.Write "        <input name='Next' type='submit' id='Next' style='cursor:hand;' value='下一步'>&nbsp;" & vbCrLf
    Response.Write "        <input name='Cancel' type='button' id='Cancel' style='cursor:hand;' onclick='window.location.href=""../index.asp""' value=' 取消 '>" & vbCrLf
    Response.Write "      </form>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    rsGetPassword.Close
    Set rsGetPassword = Nothing
End Sub

Sub Step4()
    If Password = "" Or GetStrLen(Password) > 12 Or GetStrLen(Password) < 6 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请输入密码(不能大于12小于6)</li>"
    Else
        If InStr(Password, "=") > 0 Or InStr(Password, "%") > 0 Or InStr(Password, Chr(32)) > 0 Or InStr(Password, "?") > 0 Or InStr(Password, "&") > 0 Or InStr(Password, ";") > 0 Or InStr(Password, ",") > 0 Or InStr(Password, "'") > 0 Or InStr(Password, ",") > 0 Or InStr(Password, Chr(34)) > 0 Or InStr(Password, Chr(9)) > 0 Or InStr(Password, "") > 0 Or InStr(Password, "$") > 0 Then
            ErrMsg = ErrMsg + "<br><li>密码中含有非法字符</li>"
            FoundErr = True
        End If
    End If
    If PwdConfirm = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请输入确认密码(不能大于12小于6)</li>"
    Else
        If Password <> PwdConfirm Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>密码和确认密码不一致</li>"
        End If
    End If
    If FoundErr = True Then Exit Sub


    sUserName = ReplaceBadChar(sUserName)
    Set rsGetPassword = server.CreateObject("adodb.recordset")
    rsGetPassword.open "select UserName,Question,Answer,UserPassword from PE_User where UserName='" & sUserName & "'", Conn, 1, 3
    If rsGetPassword.bof And rsGetPassword.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>对不起，用户名不存在！可能已经被管理员删除了。</li>"
    Else
        If rsGetPassword("Answer") <> MD5(Answer, 16) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>对不起，你的答案不对！</li>"
        End If
    End If
    If FoundErr = True Then
        rsGetPassword.Close
        Set rsGetPassword = Nothing
        Exit Sub
    End If

    rsGetPassword("UserPassword") = MD5(Password, 16)
    rsGetPassword.Update

    '加入对整合接口的支持
    If API_Enable Then
        FoundErr = False
        ErrMsg = ""
        If createXmlDom Then
            sPE_Items(conAction, 1) = "update"
            sPE_Items(conUsername, 1) = sUserName
            sPE_Items(conPassword, 1) = Password
            prepareXml True
            SendPost
            If FoundErr Then
                ErrMsg = "li>" & ErrMsg & "</li>"
            End If
        Else
            FoundErr = True
            ErrMsg = "<br><li>用户服务目前不可用。[APIError-XmlDom-Runtime]</li>"
        End If
    End If
    If FoundErr = True Then Exit Sub
    '完毕

    Response.Write "<br>" & vbCrLf
    Response.Write "<table align='center' width='300' border='0' cellpadding='4' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height='15' colspan='2' class='title'>找回密码 &gt;&gt; 第四步：成功设置新密码</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height='100' colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <table width='90%' border='0' cellspacing='5' cellpadding='0'>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td width='80' align='right'><strong>用户名：</strong></td>" & vbCrLf
    Response.Write "          <td>" & sUserName & "</td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td width='80' align='right'><strong>新密码：</strong></td>" & vbCrLf
    Response.Write "          <td><strong>" & Password & "</strong></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "      <br>" & vbCrLf
    Response.Write "      <font color='#FF0000'>请记住您的新密码并使用新密码<a href='../index.asp'>登录</a>！</font><br><br>" & vbCrLf
    Response.Write "      <a href='../index.asp'>【返回首页】</a>&nbsp;&nbsp;" & vbCrLf
    Response.Write "      <a href='javascript:window.close();'>【关闭窗口】</a>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    rsGetPassword.Close
    Set rsGetPassword = Nothing
End Sub
%>
