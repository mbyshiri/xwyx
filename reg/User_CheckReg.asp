<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="../API/API_Config.asp"-->
<!--#include file="../API/API_Function.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2008 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************


Call User_CheckReg
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Call CloseConn

Sub User_CheckReg()
    Dim RegUserName, rsCheckReg
    RegUserName = Trim(request("UserName"))
    If CheckUserBadChar(RegUserName) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�û����к��зǷ��ַ�</li>"
    End If

    If FoundErr = True Then Exit Sub

    If RegUserName = "" Or GetStrLen(RegUserName) > UserNameMax Or GetStrLen(RegUserName) < UserNameLimit Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�������û���(���ܴ���" & UserNameMax & "С��" & UserNameLimit & ")</li>"
    End If

    If FoundInArr(UserName_RegDisabled, RegUserName, "|") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��������û���Ϊϵͳ��ֹע����û�����</li>"
    End If

    Set rsCheckReg = Conn.Execute("select UserName from PE_User where UserName='" & RegUserName & "'")
    If Not (rsCheckReg.bof And rsCheckReg.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��" & RegUserName & "���Ѿ����ڣ��뻻һ���û��������ԣ�</li>"
    End If

    rsCheckReg.Close
    Set rsCheckReg = Nothing
    If FoundErr = True Then Exit Sub
    
    '��Ӷ����Ͻӿڵ�֧��
    If API_Enable Then
        sPE_Items(conAction, 1) = "checkname"
        sPE_Items(conUsername, 1) = RegUserName
        If createXmlDom Then
            prepareXml True
            SendPost
            If FoundErr Then
                ErrMsg = "<li>" & ErrMsg & "</li>" & vbNewLine
            End If
        Else
            FoundErr = True
            ErrMsg = "<li>��������֧��MSXML����ע����񲻿���! [APIError-XmlDom-Runtime]</li>" & vbNewLine
        End If
    End If
    '���
    If FoundErr = True Then Exit Sub

    Call WriteSuccessMsg("��" & RegUserName & "�� ��δ����ʹ�ã��Ͻ�ע��ɣ�", ComeUrl)
End Sub

%>