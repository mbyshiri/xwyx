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
Const PlatformID = 5  '����֧��
Call CheckPlatformID(PlatformID)
Dim PaySuccess
PaySuccess = False

Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
Dim md5string

v_mid = Request("MerchantID")
'ע���̻������жϴ��̻�ID�ǲ��������̻�ID
v_oid = Request("MerchantOrderNumber") '���̻�֧�������еĶ�������ͬ
'WestPayOrderNumber = Request("WestPayOrderNumber")
v_amount = Request("PaidAmount") 'WestPay���ص�ʵ��֧������CCURתΪ�����͡�
'ע���̻�����������Ǵ����̻�ԭʼ�������ҵ�ԭʼ�������Ƚ�ʵ������ԭʼ��������ͬ����֧���ɹ���

Dim objHttp, str

' ׼���ش�֧��֪ͨ��
str = Request.Form & "&cmd=validate"
Set objHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
 
'��WestPay������֪ͨ�����ٴ��ص�WestPay����֤��ȷ��֪ͨ��Ϣ����ʵ��
objHttp.Open "POST", "http://www.yeepay.com/pay/ISPN.asp", False    'ISPN: Instant Secure Payment Notification
objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.Send str
If (objHttp.Status <> 200) Then
    'HTTP ������
    Response.Write ("Status=" & objHttp.Status)
ElseIf (objHttp.ResponseText = "VERIFIED") Then
    '֧��֪ͨ��֤�ɹ�
    If Trim(v_mid) = Trim(AccountsID) Then '�жϴ˶����ǲ��Ǹ��̻��Ķ�����
        PaySuccess = True
    End If
ElseIf (objHttp.ResponseText = "INVALID") Then
    '֧��֪ͨ��֤ʧ��
    Response.Write ("Invalid")
Else
    '֧��֪ͨ��֤�����г��ִ���
    Response.Write ("Error")
End If
Set objHttp = Nothing

If PaySuccess = True Then
    Response.Write "<br>��ϲ�㣡����֧���ɹ���<br><br>"
    Call UpdateOrder(v_oid, v_amount, v_pstring, v_pmode, 3, True, True)
Else
    Response.Write "����֧��ʧ�ܣ�"
End If
Call CloseConn
%>
</BODY>
</HTML>