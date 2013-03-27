<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>在线支付结果</TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY>
<%
If Request("PayMessage") = "ok" Then
	Response.Write "在线支付成功！"
Else
	Response.Write "在线支付失败！"
End If
%>
</BODY>
</HTML>
