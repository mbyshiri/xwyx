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
Const PlatformID = 9  '��Ǯ֧��
Call CheckPlatformID(PlatformID)

Dim md5string

Dim merchantAcctId, key, version, language, signType, payType, bankId, orderId, orderTime, orderAmount, dealId, bankDealId, dealTime, payAmount
Dim fee, ext1, ext2, payResult, errCode, signMsg, merchantSignMsgVal

merchantAcctId = Trim(request("merchantAcctId")) '��ȡ����������˻���
key = MD5Key '���������������Կ
version = Trim(request("version")) '��ȡ���ذ汾
language = Trim(request("language")) '��ȡ��������,1�������ģ�2����Ӣ��
signType = Trim(request("signType")) 'ǩ������,1����MD5ǩ��
payType = Trim(request("payType")) '��ȡ֧����ʽ,00�����֧��,10�����п�֧��,11���绰����֧��,12����Ǯ�˻�֧��,13������֧��,14��B2B֧��
bankId = Trim(request("bankId")) '��ȡ���д���
orderId = Trim(request("orderId")) '��ȡ�̻�������
orderTime = Trim(request("orderTime")) '��ȡ�����ύʱ��
orderAmount = Trim(request("orderAmount")) '��ȡԭʼ�������
dealId = Trim(request("dealId")) '��ȡ��Ǯ���׺�
bankDealId = Trim(request("bankDealId")) '��ȡ���н��׺�
dealTime = Trim(request("dealTime")) '��ȡ�ڿ�Ǯ����ʱ��
payAmount = Trim(request("payAmount")) '��ȡʵ��֧�����,��λΪ��
fee = Trim(request("fee")) '��ȡ����������
ext1 = Trim(request("ext1")) '��ȡ��չ�ֶ�1
ext2 = Trim(request("ext2")) '��ȡ��չ�ֶ�2

'��ȡ������
''10���� �ɹ�11���� ʧ��
''00���� �¶����ɹ������Ե绰����֧���������أ�;01���� �¶���ʧ�ܣ����Ե绰����֧���������أ�
payResult = Trim(request("payResult"))
errCode = Trim(request("errCode")) '��ȡ�������,��ϸ���ĵ���������б�
signMsg = Trim(request("signMsg")) '��ȡ����ǩ����

'���ɼ��ܴ������뱣������˳��
merchantSignMsgVal = appendParam(merchantSignMsgVal, "merchantAcctId", merchantAcctId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "version", version)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "language", language)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "signType", signType)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "payType", payType)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "bankId", bankId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "orderId", orderId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "orderTime", orderTime)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "orderAmount", orderAmount)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "dealId", dealId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "bankDealId", bankDealId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "dealTime", dealTime)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "payAmount", payAmount)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "fee", fee)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "ext1", ext1)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "ext2", ext2)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "payResult", payResult)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "errCode", errCode)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "key", key)

md5string = MD5(merchantSignMsgVal, 32)

''���Ƚ���ǩ���ַ�����֤
If UCase(signMsg) = UCase(md5string) And payResult="10" Then
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


