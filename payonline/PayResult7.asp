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
Const PlatformID = 7  '云网支付
Call CheckPlatformID(PlatformID)
Dim PaySuccess
PaySuccess = False

Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
Dim md5string

Dim c_mid, c_order, c_orderamount, c_ymd, c_transnum, c_succmark, c_moneytype, c_cause, c_memo1, c_memo2, c_signstr

c_mid = Request("c_mid")                    '商户编号，在申请商户成功后即可获得，可以在申请商户成功的邮件中获取该编号
c_order = Request("c_order")                '商户提供的订单号
c_orderamount = Request("c_orderamount")    '商户提供的订单总金额，以元为单位，小数点后保留两位，如：13.05
c_ymd = Request("c_ymd")                    '商户传输过来的订单产生日期，格式为"yyyymmdd"，如20050102
c_transnum = Request("c_transnum")          '云网支付网关提供的该笔订单的交易流水号，供日后查询、核对使用；
c_succmark = Request("c_succmark")          '交易成功标志，Y-成功 N-失败
c_moneytype = Request("c_moneytype")        '支付币种，0为人民币
c_cause = Request("c_cause")                '如果订单支付失败，则该值代表失败原因
c_memo1 = Request("c_memo1")                '商户提供的需要在支付结果通知中转发的商户参数一
c_memo2 = Request("c_memo2")                '商户提供的需要在支付结果通知中转发的商户参数二
c_signstr = Request("c_signstr")            '云网支付网关对已上信息进行MD5加密后的字符串

md5string = MD5(c_mid & c_order & c_orderamount & c_ymd & c_transnum & c_succmark & c_moneytype & c_memo1 & c_memo2 & MD5Key, 32)

If UCase(md5string) <> UCase(c_signstr) Then
    Response.Write "签名验证失败"
    Response.End
End If

If Trim(AccountsID) <> c_mid Then
    Response.Write "提交的商户编号有误"
    Response.End
End If

If c_succmark <> "Y" And c_succmark <> "N" Then
    Response.Write "参数提交有误"
    Response.End
End If

PaySuccess = True
v_oid = c_order
v_amount = c_orderamount
v_pstring = ""
v_pmode = ""

If PaySuccess = True Then
    Response.Write "<br>恭喜你！在线支付成功！<br><br>"
    Call UpdateOrder(v_oid, v_amount, v_pstring, v_pmode, 3, True, True)
Else
    Response.Write "在线支付失败！"
End If
Call CloseConn
%>
</BODY>
</HTML>