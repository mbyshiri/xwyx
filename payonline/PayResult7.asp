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
Const PlatformID = 7  '����֧��
Call CheckPlatformID(PlatformID)
Dim PaySuccess
PaySuccess = False

Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
Dim md5string

Dim c_mid, c_order, c_orderamount, c_ymd, c_transnum, c_succmark, c_moneytype, c_cause, c_memo1, c_memo2, c_signstr

c_mid = Request("c_mid")                    '�̻���ţ��������̻��ɹ��󼴿ɻ�ã������������̻��ɹ����ʼ��л�ȡ�ñ��
c_order = Request("c_order")                '�̻��ṩ�Ķ�����
c_orderamount = Request("c_orderamount")    '�̻��ṩ�Ķ����ܽ���ԪΪ��λ��С���������λ���磺13.05
c_ymd = Request("c_ymd")                    '�̻���������Ķ����������ڣ���ʽΪ"yyyymmdd"����20050102
c_transnum = Request("c_transnum")          '����֧�������ṩ�ĸñʶ����Ľ�����ˮ�ţ����պ��ѯ���˶�ʹ�ã�
c_succmark = Request("c_succmark")          '���׳ɹ���־��Y-�ɹ� N-ʧ��
c_moneytype = Request("c_moneytype")        '֧�����֣�0Ϊ�����
c_cause = Request("c_cause")                '�������֧��ʧ�ܣ����ֵ����ʧ��ԭ��
c_memo1 = Request("c_memo1")                '�̻��ṩ����Ҫ��֧�����֪ͨ��ת�����̻�����һ
c_memo2 = Request("c_memo2")                '�̻��ṩ����Ҫ��֧�����֪ͨ��ת�����̻�������
c_signstr = Request("c_signstr")            '����֧�����ض�������Ϣ����MD5���ܺ���ַ���

md5string = MD5(c_mid & c_order & c_orderamount & c_ymd & c_transnum & c_succmark & c_moneytype & c_memo1 & c_memo2 & MD5Key, 32)

If UCase(md5string) <> UCase(c_signstr) Then
    Response.Write "ǩ����֤ʧ��"
    Response.End
End If

If Trim(AccountsID) <> c_mid Then
    Response.Write "�ύ���̻��������"
    Response.End
End If

If c_succmark <> "Y" And c_succmark <> "N" Then
    Response.Write "�����ύ����"
    Response.End
End If

PaySuccess = True
v_oid = c_order
v_amount = c_orderamount
v_pstring = ""
v_pmode = ""

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