<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>����֧�����</TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY>
<%
If Request("PayMessage") = "ok" Then
	Response.Write "����֧���ɹ���"
Else
	Response.Write "����֧��ʧ�ܣ�"
End If
%>
</BODY>
</HTML>
