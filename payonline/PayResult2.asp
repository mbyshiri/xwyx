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
Const PlatformID = 2  '�й�����֧����
Call CheckPlatformID(PlatformID)
Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
Dim md5string

v_mid = AccountsID
v_date = Trim(Request("v_date"))      '��������
v_oid = Trim(Request("v_oid"))       '֧��������
v_amount = Trim(Request("v_amount"))   '�������
v_pstatus = Trim(Request("v_status"))   '����״̬
v_md5 = Trim(Request("v_md5"))         'MD5ǩ��
md5string = MD5(v_date & v_mid & v_oid & v_amount & v_pstatus & MD5Key, 32)
v_pmode = ""
v_pstring = ""
If UCase(v_md5) = UCase(md5string) And v_pstatus = "00" Then
    Response.Write "<br>��ϲ�㣡����֧���ɹ���<br><br>"
    v_oid = Prefix_PaymentNum & v_oid
    Call UpdateOrder(v_oid, v_amount, v_pstring, v_pmode, 3, True, True)
Else
    Response.Write "����֧��ʧ�ܣ�"
End If
Call CloseConn
%>
</BODY>
</HTML>

