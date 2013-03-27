<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="UpdateOrder.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

%>
<HTML>
<HEAD>
<TITLE>����֧�����</TITLE>
</HEAD>
<BODY style="font-size:9pt;">
<%
Const IsMessageShow = True
Const PlatformID = 4  '�й�����
Call CheckPlatformID(PlatformID)
Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
v_mid = AccountsID

Dim EncodeMsg, SignMsg
EncodeMsg = Trim(Request("EncodeMsg"))                       '֧�����������Ϣ
SignMsg = Trim(Request("SignMsg"))                           'ʱ���ǩ��

'���������Ƿ���ȷ
If Len(EncodeMsg) = 0 Or Len(SignMsg) = 0 Then
    Response.Write "The Payment Result Parameters Is Empty!"
    Response.End
End If

'����Ϣ���н��ܲ�У��ʱ���ǩ��
Dim obj, bolRet, DecryptedMsg, ErrMsg, SignerCert, SignedTime

Set obj = Server.CreateObject("OpenVendorV34.NetTran")

Dim SendCertPath, RcvCertPath, RcvCertPWD
SendCertPath = "c:\certs\GNETEWEB-TEST.cer"         '���ͷ�֤��·��(����֤��)
RcvCertPath = "c:\certs\MERCHANT.pfx"               '���շ�֤��·��(�̻�֤��)
RcvCertPWD = "12345678"                                     '�����շ�֤������(�̻�֤��)

'���н���
If obj.DecryptMsg(EncodeMsg, RcvCertPath, RcvCertPWD) = 0 Then
    DecryptedMsg = obj.LastResult
Else
    Response.Write "<font color=red>Err No.: 103<br>Err Description: The PayGate's Encrypt Information Is Incorrect!</font>"
    Response.End
End If

'У��ǩ���Ƿ�һ��
If obj.VerifyMsg(SignMsg, DecryptedMsg, SendCertPath) <> 0 Then
    Response.Write "<font color=red>Err No.: 104<br>Err Description: The PayGate's Sign Information Is Incorrect!</font>" & ErrMsg
    Response.End
End If
Set obj = Nothing

'���ݽ��ܺ�����ݷֽ��������Ϣ
Dim OrderNo, PayNo, PayAmount, CurrCode, SystemSSN, RespCode, SettDate, Reserved01, Reserved02
OrderNo = GetValue(DecryptedMsg, "OrderNo")         '�̻�������
PayNo = GetValue(DecryptedMsg, "PayNo")             '֧������
PayAmount = GetValue(DecryptedMsg, "PayAmount")         '֧������ʽ��Ԫ.�Ƿ�
CurrCode = GetValue(DecryptedMsg, "CurrCode")           '���Ҵ���
SystemSSN = GetValue(DecryptedMsg, "SystemSSN")         'ϵͳ�ο���
RespCode = GetValue(DecryptedMsg, "RespCode")           '��Ӧ��
SettDate = GetValue(DecryptedMsg, "SettDate")           '�������ڣ���ʽ����������
Reserved01 = GetValue(DecryptedMsg, "Reserved01")       '������1
Reserved02 = GetValue(DecryptedMsg, "Reserved02")       '������2

'���֧��������˿�
'----------------------------------------------------------------------------------------
If RespCode = "00" Then
    v_oid = OrderNo
    v_amount = PayAmount
    v_pstring = SystemSSN
    v_pmode = ""
    Response.Write "<br>��ϲ�㣡����֧���ɹ���<br><br>"
    Call UpdateOrder(v_oid, v_amount, v_pstring, v_pmode, 3, True, True)
Else
    Response.Write "<font color=red>֧��ʧ��!��Ӧ��Ϊ��" & RespCode & "</font>"
End If

Call CloseConn
%>
</BODY>
</HTML>