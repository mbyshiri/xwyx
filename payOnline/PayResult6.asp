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
Const PlatformID = 6  '�׸�ͨ
Call CheckPlatformID(PlatformID)
Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
Dim md5string
Dim v_sid
v_mid = AccountsID
v_oid = Trim(Request("bid"))       '֧��������
v_sid = Trim(Request("sid"))         '�׸�ͨ���׳ɹ� ��ˮ��
v_md5 = Trim(Request("md"))       '����ǩ��
v_amount = Trim(Request("prc"))       '֧�����
v_pstatus = Trim(Request("success"))       '֧��״̬
v_pmode = Trim(Request("bankcode"))       '֧������
v_pstring = Trim(Request("v_pstring"))       '֧�����˵��

md5string = MD5(MD5Key & ":" & v_oid & "," & v_sid & "," & v_amount & ",sell,," & v_mid & ",bank," & v_pstatus, 32)

If UCase(v_md5) = UCase(md5string) And LCase(v_pstatus) = "true" Then
    Response.Write "<br>��ϲ�㣡����֧���ɹ���<br><br>"
    Call UpdateOrder(v_oid, v_amount, v_pstring, v_pmode, 3, True, True)
Else
    Response.Write "MD5У��ʧ�ܣ�"
End If

Call CloseConn
%>
</BODY>
</HTML>