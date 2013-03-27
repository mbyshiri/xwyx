<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="../API/API_Config.asp"-->
<!--#include file="../API/API_Function.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
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
    Response.Write "    <td height='15' colspan='2' class='title'>�һ����� &gt;&gt; ��һ���������û���</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height='100' colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <form name='form1' method='post' action=''>" & vbCrLf
    Response.Write "        <strong> ����������û�����</strong>" & vbCrLf
    Response.Write "        <input name='UserName' type='text' id='UserName' size='20' maxlength='20'><br><br>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='step2'>" & vbCrLf
    Response.Write "        <input name='Next' type='submit' id='Next' style='cursor:hand;' value='��һ��'>" & vbCrLf
    Response.Write "        <input name='Cancel' type='button' id='Cancel' style='cursor:hand;' onclick='window.location.href=""../index.asp""' value=' ȡ�� '>" & vbCrLf
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
        ErrMsg = ErrMsg & "<br><li>�Բ�����������û��������ڣ�</li>"
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
    Response.Write "    <td height='15' colspan='2' class='title'>�һ����� &gt;&gt; �ڶ������ش�����</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height='100' colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <form name='form1' method='post' action=''>" & vbCrLf
    Response.Write "        <table width='100%' border='0' cellspacing='5' cellpadding='0'>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td width='44%' align='right'><strong>������ʾ���⣺</strong></td>" & vbCrLf
    Response.Write "            <td width='56%'>" & rsGetPassword("Question") & "</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td align='right'><strong>��Ĵ𰸣�</strong></td>" & vbCrLf
    Response.Write "            <td><input name='Answer' type='text' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td align='right'><strong>��֤�룺</strong></td>" & vbCrLf
    Response.Write "            <td><input name='CheckCode' type='text' size='6' maxlength='6'> <a href='javascript:refreshimg()' title='�������������ͼƬ'><img id='checkcode' src='../inc/checkcode.asp' style='border: 1px solid #ffffff'></a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
    Response.Write "        <br>" & vbCrLf
    Response.Write "        <input name='UserName' type='hidden' id='UserName' value='" & rsGetPassword("UserName") & "'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='step3'>" & vbCrLf
    Response.Write "        <input name='PrevStep' type='button' id='PrevStep' value='��һ��' style='cursor:hand;' onclick='history.go(-1)'>&nbsp;" & vbCrLf
    Response.Write "        <input name='NextStep' type='submit' id='NextStep' style='cursor:hand;' value='��һ��'>&nbsp;" & vbCrLf
    Response.Write "        <input name='Cancel' type='button' id='Cancel' style='cursor:hand;' onclick='window.location.href=""../index.asp""' value=' ȡ�� '>" & vbCrLf
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
        ErrMsg = ErrMsg & "<br><li>��������ʾ����Ĵ𰸣�</li>"
        Exit Sub
    End If
    CheckCode = LCase(ReplaceBadChar(Trim(Request("CheckCode"))))
    If CheckCode = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��֤�벻��Ϊ�գ�</li>"
    End If
    If Trim(Session("CheckCode")) = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>���¼ʱ������������·��ص�¼ҳ����е�¼��</li>"
    End If
    If CheckCode <> Session("CheckCode") Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�������ȷ�����ϵͳ�����Ĳ�һ�£����������롣</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    sUserName = ReplaceBadChar(sUserName)
    Set rsGetPassword = server.CreateObject("adodb.recordset")
    rsGetPassword.open "select UserName,Question,Answer,UserPassword from PE_User where UserName='" & sUserName & "'", Conn, 1, 1
    If rsGetPassword.bof And rsGetPassword.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�Բ����û��������ڣ������Ѿ�������Աɾ���ˡ�</li>"
    Else
        If rsGetPassword("Answer") <> MD5(Answer, 16) Then
            '�Զ������ܽ���ļ��ݴ���
            If rsGetPassword("Answer") <> MD5(Answer, 16) Then
                FoundErr = True
                ErrMsg = ErrMsg & "<br><li>�Բ������Ĵ𰸲��ԣ�</li>"
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
    Response.Write "    <td height='15' colspan='2' class='title'>�һ����� &gt;&gt; ������������������</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height='100' colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <form name='form1' method='post' action=''>" & vbCrLf
    Response.Write "        <table width='100%' border='0' cellspacing='5' cellpadding='0'>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td width='44%' align='right'><strong>������ʾ���⣺</strong></td>" & vbCrLf
    Response.Write "            <td width='56%'>" & rsGetPassword("Question") & "</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td align='right'><strong>��Ĵ𰸣�</strong></td>" & vbCrLf
    Response.Write "            <td>" & Answer & " <input name='Answer' type='hidden' id='Answer' value='" & Answer & "'></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td align='right'><strong>�����룺</strong></td>" & vbCrLf
    Response.Write "            <td><input name='Password' type='password' id='Password' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td align='right'><strong>ȷ�������룺</strong></td>" & vbCrLf
    Response.Write "            <td><input name='PwdConfirm' type='password' id='PwdConfirm' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
    Response.Write "        <br>" & vbCrLf
    Response.Write "        <input name='UserName' type='hidden' id='UserName' value='" & rsGetPassword("Username") & "'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='step4'>" & vbCrLf
    Response.Write "        <input name='PrevStep' type='button' id='PrevStep' value='��һ��' style='cursor:hand;' onclick='history.go(-1)'>&nbsp;" & vbCrLf
    Response.Write "        <input name='Next' type='submit' id='Next' style='cursor:hand;' value='��һ��'>&nbsp;" & vbCrLf
    Response.Write "        <input name='Cancel' type='button' id='Cancel' style='cursor:hand;' onclick='window.location.href=""../index.asp""' value=' ȡ�� '>" & vbCrLf
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
        ErrMsg = ErrMsg & "<br><li>����������(���ܴ���12С��6)</li>"
    Else
        If InStr(Password, "=") > 0 Or InStr(Password, "%") > 0 Or InStr(Password, Chr(32)) > 0 Or InStr(Password, "?") > 0 Or InStr(Password, "&") > 0 Or InStr(Password, ";") > 0 Or InStr(Password, ",") > 0 Or InStr(Password, "'") > 0 Or InStr(Password, ",") > 0 Or InStr(Password, Chr(34)) > 0 Or InStr(Password, Chr(9)) > 0 Or InStr(Password, "��") > 0 Or InStr(Password, "$") > 0 Then
            ErrMsg = ErrMsg + "<br><li>�����к��зǷ��ַ�</li>"
            FoundErr = True
        End If
    End If
    If PwdConfirm = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>������ȷ������(���ܴ���12С��6)</li>"
    Else
        If Password <> PwdConfirm Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>�����ȷ�����벻һ��</li>"
        End If
    End If
    If FoundErr = True Then Exit Sub


    sUserName = ReplaceBadChar(sUserName)
    Set rsGetPassword = server.CreateObject("adodb.recordset")
    rsGetPassword.open "select UserName,Question,Answer,UserPassword from PE_User where UserName='" & sUserName & "'", Conn, 1, 3
    If rsGetPassword.bof And rsGetPassword.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�Բ����û��������ڣ������Ѿ�������Աɾ���ˡ�</li>"
    Else
        If rsGetPassword("Answer") <> MD5(Answer, 16) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br><li>�Բ�����Ĵ𰸲��ԣ�</li>"
        End If
    End If
    If FoundErr = True Then
        rsGetPassword.Close
        Set rsGetPassword = Nothing
        Exit Sub
    End If

    rsGetPassword("UserPassword") = MD5(Password, 16)
    rsGetPassword.Update

    '��������Ͻӿڵ�֧��
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
            ErrMsg = "<br><li>�û�����Ŀǰ�����á�[APIError-XmlDom-Runtime]</li>"
        End If
    End If
    If FoundErr = True Then Exit Sub
    '���

    Response.Write "<br>" & vbCrLf
    Response.Write "<table align='center' width='300' border='0' cellpadding='4' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height='15' colspan='2' class='title'>�һ����� &gt;&gt; ���Ĳ����ɹ�����������</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height='100' colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <table width='90%' border='0' cellspacing='5' cellpadding='0'>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td width='80' align='right'><strong>�û�����</strong></td>" & vbCrLf
    Response.Write "          <td>" & sUserName & "</td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td width='80' align='right'><strong>�����룺</strong></td>" & vbCrLf
    Response.Write "          <td><strong>" & Password & "</strong></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "      <br>" & vbCrLf
    Response.Write "      <font color='#FF0000'>���ס���������벢ʹ��������<a href='../index.asp'>��¼</a>��</font><br><br>" & vbCrLf
    Response.Write "      <a href='../index.asp'>��������ҳ��</a>&nbsp;&nbsp;" & vbCrLf
    Response.Write "      <a href='javascript:window.close();'>���رմ��ڡ�</a>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    rsGetPassword.Close
    Set rsGetPassword = Nothing
End Sub
%>
