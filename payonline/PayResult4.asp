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
Const PlatformID = 4  '中国银联
Call CheckPlatformID(PlatformID)
Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
v_mid = AccountsID

Dim EncodeMsg, SignMsg
EncodeMsg = Trim(Request("EncodeMsg"))                       '支付结果加密信息
SignMsg = Trim(Request("SignMsg"))                           '时间戳签名

'检验数据是否正确
If Len(EncodeMsg) = 0 Or Len(SignMsg) = 0 Then
    Response.Write "The Payment Result Parameters Is Empty!"
    Response.End
End If

'对信息进行解密并校验时间戳签名
Dim obj, bolRet, DecryptedMsg, ErrMsg, SignerCert, SignedTime

Set obj = Server.CreateObject("OpenVendorV34.NetTran")

Dim SendCertPath, RcvCertPath, RcvCertPWD
SendCertPath = "c:\certs\GNETEWEB-TEST.cer"         '发送方证书路径(银联证书)
RcvCertPath = "c:\certs\MERCHANT.pfx"               '接收方证书路径(商户证书)
RcvCertPWD = "12345678"                                     '发接收方证书密码(商户证书)

'进行解密
If obj.DecryptMsg(EncodeMsg, RcvCertPath, RcvCertPWD) = 0 Then
    DecryptedMsg = obj.LastResult
Else
    Response.Write "<font color=red>Err No.: 103<br>Err Description: The PayGate's Encrypt Information Is Incorrect!</font>"
    Response.End
End If

'校验签名是否一致
If obj.VerifyMsg(SignMsg, DecryptedMsg, SendCertPath) <> 0 Then
    Response.Write "<font color=red>Err No.: 104<br>Err Description: The PayGate's Sign Information Is Incorrect!</font>" & ErrMsg
    Response.End
End If
Set obj = Nothing

'根据解密后的内容分解出订单信息
Dim OrderNo, PayNo, PayAmount, CurrCode, SystemSSN, RespCode, SettDate, Reserved01, Reserved02
OrderNo = GetValue(DecryptedMsg, "OrderNo")         '商户订单号
PayNo = GetValue(DecryptedMsg, "PayNo")             '支付单号
PayAmount = GetValue(DecryptedMsg, "PayAmount")         '支付金额，格式：元.角分
CurrCode = GetValue(DecryptedMsg, "CurrCode")           '货币代码
SystemSSN = GetValue(DecryptedMsg, "SystemSSN")         '系统参考号
RespCode = GetValue(DecryptedMsg, "RespCode")           '响应码
SettDate = GetValue(DecryptedMsg, "SettDate")           '清算日期，格式：月月日日
Reserved01 = GetValue(DecryptedMsg, "Reserved01")       '保留域1
Reserved02 = GetValue(DecryptedMsg, "Reserved02")       '保留域2

'输出支付结果给顾客
'----------------------------------------------------------------------------------------
If RespCode = "00" Then
    v_oid = OrderNo
    v_amount = PayAmount
    v_pstring = SystemSSN
    v_pmode = ""
    Response.Write "<br>恭喜你！在线支付成功！<br><br>"
    Call UpdateOrder(v_oid, v_amount, v_pstring, v_pmode, 3, True, True)
Else
    Response.Write "<font color=red>支付失败!响应码为：" & RespCode & "</font>"
End If

Call CloseConn
%>
</BODY>
</HTML>