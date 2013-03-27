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
Const PlatformID = 5  '西部支付
Call CheckPlatformID(PlatformID)
Dim PaySuccess
PaySuccess = False

Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
Dim md5string

v_mid = Request("MerchantID")
'注：商户必须判断此商户ID是不是您的商户ID
v_oid = Request("MerchantOrderNumber") '和商户支付命令中的订单号相同
'WestPayOrderNumber = Request("WestPayOrderNumber")
v_amount = Request("PaidAmount") 'WestPay传回的实际支付金额。用CCUR转为数字型。
'注：商户必须根据我们传回商户原始订单号找到原始订单，比较实付金额和原始订单金额，相同才是支付成功。

Dim objHttp, str

' 准备回传支付通知表单
str = Request.Form & "&cmd=validate"
Set objHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
 
'把WestPay传来的通知内容再传回到WestPay作验证以确保通知信息的真实性
objHttp.Open "POST", "http://www.yeepay.com/pay/ISPN.asp", False    'ISPN: Instant Secure Payment Notification
objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.Send str
If (objHttp.Status <> 200) Then
    'HTTP 错误处理
    Response.Write ("Status=" & objHttp.Status)
ElseIf (objHttp.ResponseText = "VERIFIED") Then
    '支付通知验证成功
    If Trim(v_mid) = Trim(AccountsID) Then '判断此订单是不是该商户的订单。
        PaySuccess = True
    End If
ElseIf (objHttp.ResponseText = "INVALID") Then
    '支付通知验证失败
    Response.Write ("Invalid")
Else
    '支付通知验证过程中出现错误
    Response.Write ("Error")
End If
Set objHttp = Nothing

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