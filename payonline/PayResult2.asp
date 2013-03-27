<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="UpdateOrder.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

%>
<HTML>
<HEAD>
<TITLE>在线支付结果</TITLE>
</HEAD>
<BODY style="font-size:9pt;">
<%
Const IsMessageShow = True
Const PlatformID = 2  '中国在线支付网
Call CheckPlatformID(PlatformID)
Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
Dim md5string

v_mid = AccountsID
v_date = Trim(Request("v_date"))      '订单日期
v_oid = Trim(Request("v_oid"))       '支付定单号
v_amount = Trim(Request("v_amount"))   '订单金额
v_pstatus = Trim(Request("v_status"))   '订单状态
v_md5 = Trim(Request("v_md5"))         'MD5签名
md5string = MD5(v_date & v_mid & v_oid & v_amount & v_pstatus & MD5Key, 32)
v_pmode = ""
v_pstring = ""
If UCase(v_md5) = UCase(md5string) And v_pstatus = "00" Then
    Response.Write "<br>恭喜你！在线支付成功！<br><br>"
    v_oid = Prefix_PaymentNum & v_oid
    Call UpdateOrder(v_oid, v_amount, v_pstring, v_pmode, 3, True, True)
Else
    Response.Write "在线支付失败！"
End If
Call CloseConn
%>
</BODY>
</HTML>

