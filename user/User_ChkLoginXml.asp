<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="../API/API_Config.asp"-->
<!--#include file="../API/API_Function.asp"-->
<!--#include file="../Include/PowerEasy.UserXml.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim sql, rs
Dim CookieDate
Dim UserPassword, RndPassword, CheckCode
Dim UserInfReceived, rootNode


Set UserInfReceived = CreateObject("Microsoft.XMLDOM")
UserInfReceived.async = False
UserInfReceived.Load Request
Set rootNode = UserInfReceived.getElementsByTagName("root")
If rootNode.Length < 1 Then
    FoundErr = True
    ErrMsg = ErrMsg & "��������Ϊ��"
Else
    UserName = Replace(ReplaceBadChar(rootNode(0).selectSingleNode("username").text), ".", "")
    UserPassword = ReplaceBadChar(rootNode(0).selectSingleNode("password").text)
    CheckCode = LCase(ReplaceBadChar(rootNode(0).selectSingleNode("checkcode").text))
    CookieDate = PE_CLng(rootNode(0).selectSingleNode("cookiesdate").text)
    If EnableCheckCodeOfLogin = True Then
       If Trim(Session("CheckCode")) = "" Then
           FoundErr = True
           ErrMsg = ErrMsg & "��֤�볬ʱʧЧ��"
       End If
       If CheckCode <> Session("CheckCode") Then
           FoundErr = True
           ErrMsg = ErrMsg & "��֤��������������롣"
       End If
    End If
    If UserName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "�û�������Ϊ�գ�"
    End If
                
    If UserPassword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "���벻��Ϊ�գ�"
    End If
End If
Set UserInfReceived = Nothing
Response.ContentType = "text/xml; charset=gb2312"

If FoundErr <> True Then
    Set rs = Conn.Execute("select UserID,UserName,UserPassword,LoginTimes,IsLocked from PE_User where UserName='" & UserName & "'")
    If rs.bof And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "�û������ڣ�����"
    Else
        UserPassword = MD5(UserPassword, 16)
        LoginTimes = rs("LoginTimes") + 1
        If UserPassword <> rs("UserPassword") Then
            FoundErr = True
            ErrMsg = ErrMsg & "������󣡣���"
        End If
        If rs("IsLocked") = True Then
            FoundErr = True
            ErrMsg = ErrMsg & "�û��Ѿ�����ס���޷���¼���������Ҫ�������������Ա��ϵ��"
        End If
        If FoundErr = False Then
            RndPassword = GetRndPassword(16)
            '���µ�¼״̬
            Conn.Execute ("Update PE_User Set LastPassword='" & RndPassword & "',LastLoginIP='" & UserTrueIP & "',LastLoginTime=" & PE_Now & ",LoginTimes=LoginTimes+1 where UserID=" & rs("UserID") & "")
            Session("UserID") = rs("UserID")
            
            '���¹��ﳵ���û���
            Dim CartID
            CartID = ReplaceBadChar(Trim(Request.Cookies("Cart" & Site_Sn)("CartID")))
            Conn.Execute ("update PE_ShoppingCarts set UserName='" & UserName & "' where CartID='" & CartID & "'")
            
            If Enable_SubDomain Then
                Dim iSubDomainIndex, strSite_Sn
                For iSubDomainIndex = 0 To UBound(arrSubDomains)
                    strSite_Sn = LCase(arrSubDomains(iSubDomainIndex) & Replace(Replace(DomainRoot & InstallDir, "/", ""), ".", ""))
                    Response.Cookies(strSite_Sn).Domain = DomainRoot
                    Select Case CookieDate
                        Case 0
                            'not save
                        Case 1
                            Response.Cookies(strSite_Sn).Expires = Date + 1
                        Case 2
                            Response.Cookies(strSite_Sn).Expires = Date + 31
                        Case 3
                            Response.Cookies(strSite_Sn).Expires = Date + 365
                    End Select
                    Response.Cookies(strSite_Sn)("UserName") = UserName
                    Response.Cookies(strSite_Sn)("UserPassword") = UserPassword
                    Response.Cookies(strSite_Sn)("LastPassword") = RndPassword
                    Response.Cookies(strSite_Sn)("CookieDate") = CookieDate
                Next
            Else
                Select Case CookieDate
                    Case 0
                        'not save
                    Case 1
                        Response.Cookies(Site_Sn).Expires = Date + 1
                    Case 2
                        Response.Cookies(Site_Sn).Expires = Date + 31
                    Case 3
                        Response.Cookies(Site_Sn).Expires = Date + 365
                End Select
                Response.Cookies(Site_Sn)("UserName") = UserName
                Response.Cookies(Site_Sn)("UserPassword") = UserPassword
                Response.Cookies(Site_Sn)("LastPassword") = RndPassword
                Response.Cookies(Site_Sn)("CookieDate") = CookieDate
            End If

            Call GetUser(UserName)
            Call ShowUserXml(True)

        End If
    End If
    rs.Close
    Set rs = Nothing
End If
If FoundErr = True Then
    Call ShowUserErr
End If
Call CloseConn
%>
