<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="UpdateOrder.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const IsMessageShow = False
Const PlatformID = 1  '��������
Call CheckPlatformID(PlatformID)
Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
Dim md5string

v_mid = AccountsID
v_oid = Trim(Request("v_oid"))       '֧��������
v_md5 = Trim(Request("v_md5str"))       '����ǩ��
v_amount = Trim(Request("v_amount"))       '֧�����
v_pstatus = Trim(Request("v_pstatus"))       '֧��״̬
v_moneytype = Trim(Request("v_moneytype"))   '֧������
v_pmode = Trim(Request("v_pmode"))       '֧������
v_pstring = Trim(Request("v_pstring"))       '֧�����˵��

md5string = MD5(v_oid & v_pstatus & v_amount & v_moneytype & PayOnlineKey, 32)
        
If UCase(v_md5) = UCase(md5string) And v_pstatus = "20" Then
    Response.Write "ok"
    Call UpdateOrder(v_oid, v_amount, v_pstring, v_pmode, 3, True, True)
Else
    Response.Write "error"
End If
Call CloseConn
%>
