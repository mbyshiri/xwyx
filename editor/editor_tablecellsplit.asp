<%@language=vbscript codepage=936 %>
<%
option explicit
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
body, a, table, div, span, td, th, input, select{font-size:9pt;font-family: "����", Verdana, Arial, Helvetica, sans-serif;}
body {padding:5px}
</style>

<script language="JavaScript">

document.write("<title>��Ԫ����</title>");

// �ж�ֵ�Ƿ����0
function MoreThanOne(obj, sErr){
	var b=false;
	if (obj.value!=""){
		obj.value=parseFloat(obj.value);
		if (obj.value!="0"){
			b=true;
		}
	}
	if (b==false){
		BaseAlert(obj,sErr);
		return false;
	}
	return true;
}

// Ԥ��
function doView(opt){
	if (opt=="col"){
		d_col.checked=true;
		d_row.checked=false;
	}else{
		d_col.checked=false;
		d_row.checked=true;
	}
	if (d_col.checked){
		d_view.innerHTML = "<table border=1 cellpadding=0><tr><td width=25>&nbsp;</td><td width=25>&nbsp;</td></tr></table>";
		d_label.innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;����:";
	}
	if (d_row.checked){
		d_view.innerHTML = "<table border=1 cellpadding=0 width=50><tr><td>&nbsp;</td></tr><tr><td>&nbsp;</td></tr></table>";
		d_label.innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;����:";
	}
}
// ֻ������������
function IsDigit(){
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}


</script>

<SCRIPT event=onclick for=Ok language=JavaScript>
	// ����������Ч��
	if (!MoreThanOne(d_num,'��Ч�����������������1��')) return;

	if (d_row.checked){
		dialogArguments.TableRowSplit(parseInt(d_num.value));
	}
	if (d_col.checked){
		dialogArguments.TableColSplit(parseInt(d_num.value));
	}

	window.returnValue = null;
	window.close();
</SCRIPT>

</head>
<body bgColor=#D4D0C8>

<table border=0 cellpadding=0 cellspacing=0 align=center>
<tr>
	<td>
	<table border=0 cellpadding=0 cellspacing=0>
	<tr><td colspan=3 height=5></td></tr>
	<tr><td><input type=radio id=d_col checked onclick="doView('col')"><label for="d_col">���Ϊ��</label></td><td rowspan=3 width=30></td><td width=60 rowspan=3 id=d_view valign=middle align=center></td></tr>
	<tr><td height=5></td></tr>
	<tr><td><input type=radio id=d_row onclick="doView('row')"><label for="d_row">���Ϊ��</label></td></tr>
	<tr><td height=5 colspan=3></td></tr>
	<tr>
		<td id=d_label></td>
		<td></td>
		<td><input type=text id=d_num size=8 value="1" ONKEYPRESS="event.returnValue=IsDigit();" maxlength=3></td>
	</tr>
	</table>
</tr>
<tr><td height=5></td></tr>
<tr><td align=right><input type=submit value='  ȷ��  ' id=Ok>&nbsp;&nbsp;<input type=button value='  ȡ��  ' onclick="window.close();"></td></tr>
</table>

<Script Language=JavaScript>
doView('col');
</Script>

</body>
</html>