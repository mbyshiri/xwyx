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
Const PlatformID = 3  '上海环迅
Call CheckPlatformID(PlatformID)
Dim billno, currency_type, amount, mydate, succ, attach, ipsbillno, retEncodeType, signature
Dim md5string, v_oid, v_amount, v_pstring, v_pmode


billno = Request.QueryString("billno")
currency_type = Request.QueryString("currency_type")
amount = Request.QueryString("amount")
mydate = Request.QueryString("date")
succ = Request.QueryString("succ")
attach = Request.QueryString("attach")
ipsbillno = Request.QueryString("ipsbillno")
retEncodeType = Request.QueryString("retencodetype")
signature = Request.QueryString("signature")



If succ = "Y" Then
    md5string = billno & amount & mydate & succ & ipsbillno & currency_type & MD5Key
    md5string = MD5(md5string, 32)
    If md5string = UCase(signature) Then
        Response.Write "<br>恭喜你！在线支付成功！<br><br>"
        v_oid = billno
        v_amount = amount
        Call UpdateOrder(v_oid, v_amount, v_pstring, v_pmode, 3, True, True)
    Else
        Response.Write "验证失败！"
    End If
Else
    Response.Write "在线支付失败！"
End If
Call CloseConn
%>
</BODY>
</HTML>

