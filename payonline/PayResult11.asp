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
Const PlatformID = 11  '��Ǯ������
Call CheckPlatformID(PlatformID)
Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
Dim md5string

Dim merchantAcctId, key, version, language, signType, payType, orderId, orderTime, orderAmount, dealId, dealTime, payAmount, cardNumber, cardPwd, billOrderTime
Dim ext1, ext2, payResult, signMsg, merchantSignMsgVal

merchantAcctId = Trim(Request("merchantAcctId")) '��ȡ�����������˻���
key = MD5Key '����������������Կ
version = Trim(Request("version")) '������汾�Ź̶�Ϊv2.0
language = Trim(Request("language")) '1�������ģ�2����Ӣ��
payType = Trim(Request("payType")) '20���������п���ֱ��֧����22�����Ǯ�˻����������֧��
cardNumber = Trim(Request("cardNumber")) '�����п����,���ͨ�������п�ֱ��֧��ʱ����
cardPwd = Trim(Request("cardPwd")) '��ȡ�����п�����,���ͨ�������п�ֱ��֧��ʱ����
orderId = Trim(Request("orderId")) '��ȡ�̻�������
orderAmount = Trim(Request("orderAmount")) '��ȡԭʼ�������
dealId = Trim(Request("dealId")) '��ȡ��Ǯ���׺�
orderTime = Trim(Request("orderTime")) '��ȡ�̻��ύ����ʱ��ʱ��
ext1 = Trim(Request("ext1")) '��ȡ��չ�ֶ�1
ext2 = Trim(Request("ext2")) '��ȡ��չ�ֶ�2
payAmount = Trim(Request("payAmount")) '��ȡʵ��֧�����,��λΪ��
billOrderTime = Trim(Request("billOrderTime")) '��ȡ��Ǯ����ʱ��

payResult = Trim(Request("payResult")) '��ȡ������,10����֧���ɹ��� 11����֧��ʧ��
signType = Trim(Request("signType")) '��ȡǩ������,1����MD5ǩ��
signMsg = Trim(Request("signMsg")) '��ȡ����ǩ����

'���ɼ��ܴ������뱣������˳��
merchantSignMsgVal = appendParam(merchantSignMsgVal, "merchantAcctId", merchantAcctId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "version", version)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "language", language)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "payType", payType)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "cardNumber", cardNumber)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "cardPwd", cardPwd)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "orderId", orderId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "orderAmount", orderAmount)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "dealId", dealId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "orderTime", orderTime)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "ext1", ext1)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "ext2", ext2)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "payAmount", payAmount)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "billOrderTime", billOrderTime)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "payResult", payResult)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "signType", signType)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "key", key)
md5string = MD5(merchantSignMsgVal, 32)

Dim rtnOk, rtnUrl
rtnOk = 0
rtnUrl = "http://" & Trim(Request.ServerVariables("HTTP_HOST")) & Trim(Request.ServerVariables("SCRIPT_NAME"))
rtnUrl = Left(rtnUrl, InStrRev(rtnUrl, "/")) & "ShowResult.asp"

''���Ƚ���ǩ���ַ�����֤
If UCase(signMsg) = UCase(md5string) And payResult = "10" Then
    ''���Ž���֧������ж�
    Response.Write "<br>��ϲ�㣡����֧���ɹ���<br><br>"
    Call UpdateOrder(orderId, orderAmount / 100, "", "", 3, True, True)
Else
    Response.Write "����֧��ʧ�ܣ�"
End If
Call CloseConn
%>
</BODY>
</HTML>
<%
'������ֵ��Ϊ�յĲ�������ַ���
Function appendParam(returnStr, paramId, paramValue)
    If returnStr <> "" Then
        If paramValue <> "" Then
            returnStr=returnStr&"&"&paramId&"="&paramValue
        End If
    Else
        If paramValue <> "" Then
            returnStr=paramId&"="&paramValue
        End If
    End If
    appendParam = returnStr
End Function
%>




