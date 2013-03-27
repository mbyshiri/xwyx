<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<!--#include file="UpdateOrder.asp"-->
<HTML>
<HEAD>
<TITLE>在线支付结果</TITLE>
</HEAD>
<BODY style="font-size:9pt;">
<%
Const IsMessageShow = True
Const PlatformID = 13  '财付通
Call CheckPlatformID(PlatformID)
Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
Dim md5string
v_mid = AccountsID

Dim cmdno, pay_result, pay_info, bill_date, bargainor_id, transaction_id, sp_billno, total_fee, fee_type, md5_sign, attach
cmdno = Request("cmdno")
pay_result = Request("pay_result")
pay_info = Request("pay_info")
bill_date = Request("date")
bargainor_id = Request("bargainor_id")
transaction_id = Request("transaction_id")
sp_billno = Request("sp_billno")
total_fee = Request("total_fee")
fee_type = Request("fee_type")
attach = Request("attach")
md5_sign = Request("sign")

md5string = MD5("cmdno=" & cmdno & "&pay_result=" & pay_result & "&date=" & bill_date & "&transaction_id=" & transaction_id & "&sp_billno=" & sp_billno & "&total_fee=" & total_fee & "&fee_type=" & fee_type & "&attach=" & attach & "&key=" & MD5Key, 32)

If bargainor_id = v_mid And UCase(md5string) = md5_sign And pay_result = 0 Then
    Response.Write "<br>恭喜你！在线支付成功！<br><br>"
    v_oid = sp_billno
    v_amount = total_fee / 100
    v_pstring = ""
    v_pmode = ""
    Call UpdateOrder(v_oid, v_amount, v_pstring, v_pmode, 3, True, True)
Else
    Response.Write "在线支付失败！"
End If
Call CloseConn
%>
</BODY>
</HTML>

