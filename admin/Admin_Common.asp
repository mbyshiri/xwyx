<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.Cache.asp"-->
<!--#include file="../Include/PowerEasy.Channel.asp"-->
<!--#include file="../Include/PowerEasy.Common.Manage.asp"-->
<!--#include file="../Include/PowerEasy.Common.Purview.asp"-->
<!--#include file="Admin_CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Server.ScriptTimeOut = 9999999

ChannelID = PE_CLng(Trim(Request("ChannelID")))
If ChannelID > 0 Then
    Call GetChannel(ChannelID)
End If

If NeedCheckComeUrl = True Then
    Call CheckComeUrl
End If

'������Ա�Ƿ��¼
Dim AdminID, AdminName, AdminPassword, RndPassword, AdminLoginCode, AdminPurview, PurviewPassed
Dim AdminPurview_Channel, AdminPurview_Others, AdminPurview_GuestBook
Dim arrClass_GuestBook, arrKind_House
Dim rsGetAdmin, sqlGetAdmin
Dim arrPurview(30), PurviewIndex, strThisFile
AdminName = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminName")))
AdminPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminPassword")))
RndPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("RndPassword")))
AdminLoginCode = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminLoginCode")))
If AdminName = "" Or AdminPassword = "" Or RndPassword = "" Or (EnableSiteManageCode = True And AdminLoginCode <> SiteManageCode) Then
    Call WriteEntry(1, "", "����Աδ��¼")
    Call CloseConn
    Response.redirect "Admin_login.asp"
End If
sqlGetAdmin = "select * from PE_Admin where AdminName='" & AdminName & "' and Password='" & AdminPassword & "'"
Set rsGetAdmin = Server.CreateObject("adodb.recordset")
rsGetAdmin.Open sqlGetAdmin, Conn, 1, 1
If rsGetAdmin.BOF And rsGetAdmin.EOF Then
    Call WriteEntry(4, "", "�û������������")
    rsGetAdmin.Close
    Set rsGetAdmin = Nothing
    Call CloseConn
    Response.redirect "Admin_login.asp"
Else
    If rsGetAdmin("EnableMultiLogin") <> True And Trim(rsGetAdmin("RndPassword")) <> RndPassword Then
        Response.write "<br><p align=center><font color='red'>�Բ���Ϊ��ϵͳ��ȫ����ϵͳ������������ʹ��ͬһ������Ա�ʺŽ��е�¼��</font></p><p>��Ϊ���������Ѿ��������ط�ʹ�ô˹���Ա�ʺŽ��е�¼�ˣ������㽫���ܼ������к�̨���������</p><p>�����<a href='Admin_Login.asp' target='_top'>������µ�¼</a>��</p>"
        Call WriteEntry(1, AdminName, "����ʹ��ͬһ����Ա�ʺ�")
        rsGetAdmin.Close
        Set rsGetAdmin = Nothing
        Call CloseConn
        Response.End
    End If
End If
AdminID = rsGetAdmin("ID")
UserName = rsGetAdmin("UserName")
AdminPurview = rsGetAdmin("Purview")
AdminPurview_Others = rsGetAdmin("AdminPurview_Others")
AdminPurview_GuestBook = rsGetAdmin("AdminPurview_GuestBook")
arrClass_View = rsGetAdmin("arrClass_View") & ""
arrClass_Input = rsGetAdmin("arrClass_Input") & ""
arrClass_Check = rsGetAdmin("arrClass_Check") & ""
arrClass_Manage = rsGetAdmin("arrClass_Manage") & ""
arrClass_GuestBook = rsGetAdmin("arrClass_GuestBook") & ""
arrKind_House = Split(rsGetAdmin("arrClass_House") & "", "|||")


PurviewPassed = False   'Ĭ������Ϊû��Ȩ��
If AdminPurview = 1 Then   '����ǳ�������Ա��ֱ��������Ȩ��
    PurviewPassed = True
Else
    '�������ͨ����Ա�������ļ��������ж��Ƿ�����Ӧ��Ȩ��
    Select Case PurviewLevel
    Case 0       '���������Ȩ�޼�飬ֱ����Ȩ��
        PurviewPassed = True
    Case 1       '���Ҫ�󳬼�����Ա����ֱ���ж�û��Ȩ��
        PurviewPassed = False
    Case 2    '���Ҫ����ͨ����Ա
        If PurviewLevel_Channel <= 0 Then  '�����Ҫ���Ƶ��Ȩ������
            PurviewPassed = True
        Else
            AdminPurview_Channel = PE_CLng(rsGetAdmin("AdminPurview_" & ChannelDir))
            If AdminPurview_Channel = 0 Then AdminPurview_Channel = 5
            
            If AdminPurview_Channel <= PurviewLevel_Channel Then
                PurviewPassed = True
            End If
        End If
        If PurviewLevel_Others <> "" Then  '�����Ҫ�������Ȩ�ޣ���Ȩ����������Ϊ׼
            PurviewPassed = CheckPurview_Other(AdminPurview_Others, PurviewLevel_Others)
        End If
    End Select
End If

If PurviewPassed = False Then
    Response.write "<br><p align=center><font color='red'>�Բ�����û�д��������Ȩ�ޡ�</font></p>"
    Call WriteEntry(1, AdminName, "ԽȨʹ��")
    Call CloseConn
    Response.End
End If

%>
