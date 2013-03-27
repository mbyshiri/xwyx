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
Const PlatformID = 11  '快钱神州行
Call CheckPlatformID(PlatformID)
Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
Dim md5string

Dim merchantAcctId, key, version, language, signType, payType, orderId, orderTime, orderAmount, dealId, dealTime, payAmount, cardNumber, cardPwd, billOrderTime
Dim ext1, ext2, payResult, signMsg, merchantSignMsgVal

merchantAcctId = Trim(Request("merchantAcctId")) '获取神州行网关账户号
key = MD5Key '设置神州行网关密钥
version = Trim(Request("version")) '本代码版本号固定为v2.0
language = Trim(Request("language")) '1代表中文；2代表英文
payType = Trim(Request("payType")) '20代表神州行卡密直接支付；22代表快钱账户神州行余额支付
cardNumber = Trim(Request("cardNumber")) '神州行卡序号,如果通过神州行卡直接支付时返回
cardPwd = Trim(Request("cardPwd")) '获取神州行卡密码,如果通过神州行卡直接支付时返回
orderId = Trim(Request("orderId")) '获取商户订单号
orderAmount = Trim(Request("orderAmount")) '获取原始订单金额
dealId = Trim(Request("dealId")) '获取快钱交易号
orderTime = Trim(Request("orderTime")) '获取商户提交订单时的时间
ext1 = Trim(Request("ext1")) '获取扩展字段1
ext2 = Trim(Request("ext2")) '获取扩展字段2
payAmount = Trim(Request("payAmount")) '获取实际支付金额,单位为分
billOrderTime = Trim(Request("billOrderTime")) '获取快钱处理时间

payResult = Trim(Request("payResult")) '获取处理结果,10代表支付成功； 11代表支付失败
signType = Trim(Request("signType")) '获取签名类型,1代表MD5签名
signMsg = Trim(Request("signMsg")) '获取加密签名串

'生成加密串。必须保持如下顺序。
merchantSignMsgVal = appendParam(merchantSignMsgVal, "merchantAcctId", merchantAcctId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "version", version)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "language", language)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "payType", payType)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "cardNumber", cardNumber)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "cardPwd", cardPwd)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "orderId", orderId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "orderAmount", orderAmount)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "dealId", dealId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "orderTime", orderTime)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "ext1", ext1)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "ext2", ext2)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "payAmount", payAmount)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "billOrderTime", billOrderTime)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "payResult", payResult)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "signType", signType)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "key", key)
md5string = MD5(merchantSignMsgVal, 32)

Dim rtnOk, rtnUrl
rtnOk = 0
rtnUrl = "http://" & Trim(Request.ServerVariables("HTTP_HOST")) & Trim(Request.ServerVariables("SCRIPT_NAME"))
rtnUrl = Left(rtnUrl, InStrRev(rtnUrl, "/")) & "ShowResult.asp"

''首先进行签名字符串验证
If UCase(signMsg) = UCase(md5string) And payResult = "10" Then
    ''接着进行支付结果判断
    Response.Write "<br>恭喜你！在线支付成功！<br><br>"
    Call UpdateOrder(orderId, orderAmount / 100, "", "", 3, True, True)
Else
    Response.Write "在线支付失败！"
End If
Call CloseConn
%>
</BODY>
</HTML>
<%
'将变量值不为空的参数组成字符串
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




