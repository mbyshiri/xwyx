<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="../API/API_Config.asp"-->
<!--#include file="../API/API_Function.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2008 佛山市动易网络科技有限公司 版权所有
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
        ErrMsg = ErrMsg & "<br><li>用户名中含有非法字符</li>"
    End If

    If FoundErr = True Then Exit Sub

    If RegUserName = "" Or GetStrLen(RegUserName) > UserNameMax Or GetStrLen(RegUserName) < UserNameLimit Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>请输入用户名(不能大于" & UserNameMax & "小于" & UserNameLimit & ")</li>"
    End If

    If FoundInArr(UserName_RegDisabled, RegUserName, "|") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>您输入的用户名为系统禁止注册的用户名！</li>"
    End If

    Set rsCheckReg = Conn.Execute("select UserName from PE_User where UserName='" & RegUserName & "'")
    If Not (rsCheckReg.bof And rsCheckReg.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>“" & RegUserName & "”已经存在！请换一个用户名再试试！</li>"
    End If

    rsCheckReg.Close
    Set rsCheckReg = Nothing
    If FoundErr = True Then Exit Sub
    
    '添加对整合接口的支持
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
            ErrMsg = "<li>服务器不支持MSXML对象，注册服务不可用! [APIError-XmlDom-Runtime]</li>" & vbNewLine
        End If
    End If
    '完毕
    If FoundErr = True Then Exit Sub

    Call WriteSuccessMsg("“" & RegUserName & "” 尚未被人使用，赶紧注册吧！", ComeUrl)
End Sub

%>