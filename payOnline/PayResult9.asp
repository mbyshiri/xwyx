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
Const PlatformID = 9  '快钱支付
Call CheckPlatformID(PlatformID)

Dim md5string

Dim merchantAcctId, key, version, language, signType, payType, bankId, orderId, orderTime, orderAmount, dealId, bankDealId, dealTime, payAmount
Dim fee, ext1, ext2, payResult, errCode, signMsg, merchantSignMsgVal

merchantAcctId = Trim(request("merchantAcctId")) '获取人民币网关账户号
key = MD5Key '设置人民币网关密钥
version = Trim(request("version")) '获取网关版本
language = Trim(request("language")) '获取语言种类,1代表中文；2代表英文
signType = Trim(request("signType")) '签名类型,1代表MD5签名
payType = Trim(request("payType")) '获取支付方式,00：组合支付,10：银行卡支付,11：电话银行支付,12：快钱账户支付,13：线下支付,14：B2B支付
bankId = Trim(request("bankId")) '获取银行代码
orderId = Trim(request("orderId")) '获取商户订单号
orderTime = Trim(request("orderTime")) '获取订单提交时间
orderAmount = Trim(request("orderAmount")) '获取原始订单金额
dealId = Trim(request("dealId")) '获取快钱交易号
bankDealId = Trim(request("bankDealId")) '获取银行交易号
dealTime = Trim(request("dealTime")) '获取在快钱交易时间
payAmount = Trim(request("payAmount")) '获取实际支付金额,单位为分
fee = Trim(request("fee")) '获取交易手续费
ext1 = Trim(request("ext1")) '获取扩展字段1
ext2 = Trim(request("ext2")) '获取扩展字段2

'获取处理结果
''10代表 成功11代表 失败
''00代表 下订单成功（仅对电话银行支付订单返回）;01代表 下订单失败（仅对电话银行支付订单返回）
payResult = Trim(request("payResult"))
errCode = Trim(request("errCode")) '获取错误代码,详细见文档错误代码列表
signMsg = Trim(request("signMsg")) '获取加密签名串

'生成加密串。必须保持如下顺序。
merchantSignMsgVal = appendParam(merchantSignMsgVal, "merchantAcctId", merchantAcctId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "version", version)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "language", language)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "signType", signType)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "payType", payType)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "bankId", bankId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "orderId", orderId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "orderTime", orderTime)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "orderAmount", orderAmount)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "dealId", dealId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "bankDealId", bankDealId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "dealTime", dealTime)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "payAmount", payAmount)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "fee", fee)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "ext1", ext1)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "ext2", ext2)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "payResult", payResult)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "errCode", errCode)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "key", key)

md5string = MD5(merchantSignMsgVal, 32)

''首先进行签名字符串验证
If UCase(signMsg) = UCase(md5string) And payResult="10" Then
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


