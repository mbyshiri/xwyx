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
body, a, table, div, span, td, th, input, select{font:9pt;font-family: "����", Verdana, Arial, Helvetica, sans-serif;}
body {padding:5px}
</style>
<script language="JavaScript">
// ȡͨ��URL�������Ĳ��� (��ʽ�� ?Param1=Value1&Param2=Value2)
var URLParams = new Object() ;
var aParams = document.location.search.substr(1).split('&') ;
for (i=0 ; i < aParams.length ; i++) {
	var aParam = aParams[i].split('=') ;
	URLParams[aParam[0]] = aParam[1] ;
}
var sAction = URLParams['action'] ;
var sTitle = "����";

var oControl;
var oSeletion;
var sRangeType;

var sRow = "2";
var sCol = "2";
var sAlign = "";
var sBorder = "2";
var sCellPadding = "3";
var sCellSpacing = "0";
var sWidth = "";
var sHeight = "";
var sBorderColor = "#CCCCCC";
var sBgColor = "#FFFFFF";

var sImage = "";
var sRepeat = "";
var sAttachment = "";
var sBorderStyle = "";

var sWidthUnit = "%";
var bWidthCheck = true;
var bWidthDisable = false;
var sWidthValue = "100";

var sHeightUnit = "%";
var bHeightCheck = false;
var bHeightDisable = true;
var sHeightValue = "";

oSelection = dialogArguments.HtmlEdit.document.selection.createRange();
sRangeType = dialogArguments.HtmlEdit.document.selection.type;

if (sAction == "modify"){
	if (sRangeType == "Control"){
		if (oSelection.item(0).tagName == "TABLE"){
			oControl = oSelection.item(0);
		}
	}else{
		oControl = getParentObject(oSelection.parentElement(), "TABLE");
	}
	if (oControl){
		sAction = "MODI";
		sTitle = "�޸�";
		sRow = oControl.rows.length;
		sCol = getColCount(oControl);
		sAlign = oControl.align;
		sBorder = oControl.border;
		sCellPadding = oControl.cellPadding;
		sCellSpacing = oControl.cellSpacing;
		sWidth = oControl.width;
		sHeight = oControl.height;
		sBorderColor = oControl.borderColor;
		sBgColor = oControl.bgColor;
		sImage = oControl.style.backgroundImage;
		sRepeat = oControl.style.backgroundRepeat;
		sAttachment = oControl.style.backgroundAttachment;
		sBorderStyle = oControl.style.borderStyle;
		sImage = sImage.substr(4, sImage.length-5);
	}
}

//=================================================
//��������SelectColor
//��  �ã���ʾ��ɫ��
//��  ����what  --- Ҫ�����ɫ�Ĳ���
//=================================================
function SelectColor(what){
	var dEL = document.all("t_"+what);
	var sEL = document.all("s_"+what);
	var url = "editor_selcolor.asp?color="+encodeURIComponent(dEL.value);
	var arr = showModalDialog(url,window,"dialogWidth:280px;dialogHeight:250px;help:no;scroll:no;status:no");
	if (arr) {
		dEL.value=arr;
		sEL.style.backgroundColor=arr;
	}
}
// �Ƿ���Ч��ɫֵ
function IsColor(color){
	var temp=color;
	if (temp=="") return true;
	if (temp.length!=7) return false;
	return (temp.search(/\#[a-fA-F0-9]{6}/) != -1);
}
// ֻ������������
function IsDigit(){
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}

// ����������ֵ��ָ��ֵƥ�䣬��ѡ��ƥ����
function SearchSelectValue(o_Select, s_Value){
	for (var i=0;i<o_Select.length;i++){
		if (o_Select.options[i].value == s_Value){
			o_Select.selectedIndex = i;
			return true;
		}
	}
	return false;
}

// ����ָ����ǩ�ĸ�����
function getParentObject(obj, tag){
	while(obj!=null && obj.tagName!=tag)
		obj=obj.parentElement;
	return obj;
}

document.write("<title>������ԣ�" + sTitle + "��</title>");

// ��ʼֵ
function InitDocument(){
	SearchSelectValue(d_align, sAlign.toLowerCase());
	SearchSelectValue(d_borderstyle, sBorderStyle.toLowerCase());

	// �޸�״̬ʱȡֵ
	if (sAction == "MODI"){
		if (sWidth == ""){
			bWidthCheck = false;
			bWidthDisable = true;
			sWidthValue = "100";
			sWidthUnit = "%";
		}else{
			bWidthCheck = true;
			bWidthDisable = false;
			if (sWidth.substr(sWidth.length-1) == "%"){
				sWidthValue = sWidth.substring(0, sWidth.length-1);
				sWidthUnit = "%";
			}else{
				sWidthUnit = "";
				sWidthValue = parseInt(sWidth);
				if (isNaN(sWidthValue)) sWidthValue = "";
			}
		}
		if (sHeight == ""){
			bHeightCheck = false;
			bHeightDisable = true;
			sHeightValue = "100";
			sHeightUnit = "%";
		}else{
			bHeightCheck = true;
			bHeightDisable = false;
			if (sHeight.substr(sHeight.length-1) == "%"){
				sHeightValue = sHeight.substring(0, sHeight.length-1);
				sHeightUnit = "%";
			}else{
				sHeightUnit = "";
				sHeightValue = parseInt(sHeight);
				if (isNaN(sHeightValue)) sHeightValue = "";
			}
		}

	}

	switch(sWidthUnit){
	case "%":
		d_widthunit.selectedIndex = 1;
		break;
	default:
		sWidthUnit = "";
		d_widthunit.selectedIndex = 0;
		break;
	}
	switch(sHeightUnit){
	case "%":
		d_heightunit.selectedIndex = 1;
		break;
	default:
		sHeightUnit = "";
		d_heightunit.selectedIndex = 0;
		break;
	}

	d_row.value = sRow;
	d_col.value = sCol;
	d_border.value = sBorder;
	d_cellspacing.value = sCellSpacing;
	d_cellpadding.value = sCellPadding;
	d_widthvalue.value = sWidthValue;
	d_widthvalue.disabled = bWidthDisable;
	d_widthunit.disabled = bWidthDisable;
	d_heightvalue.value = sHeightValue;
	d_heightvalue.disabled = bHeightDisable;
	d_heightunit.disabled = bHeightDisable;
	t_bordercolor.value = sBorderColor;
	s_bordercolor.style.backgroundColor = sBorderColor;
	t_bgcolor.value = sBgColor;
	s_bgcolor.style.backgroundColor = sBgColor;
	d_widthcheck.checked = bWidthCheck;
	d_heightcheck.checked = bHeightCheck;
	d_image.value = sImage;
	d_repeat.value = sRepeat;
	d_attachment.value = sAttachment;

}

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

// �õ��������
function getColCount(oTable) {
	var intCount = 0;
	if (oTable != null) {
		for(var i = 0; i < oTable.rows.length; i++){
			if (oTable.rows[i].cells.length > intCount) intCount = oTable.rows[i].cells.length;
		}
	}
	return intCount;
}

// ������
function InsertRows( oTable ) {
	if ( oTable ) {
		var elRow=oTable.insertRow();
		for(var i=0; i<oTable.rows[0].cells.length; i++){
			var elCell = elRow.insertCell();
			elCell.innerHTML = "&nbsp;";
		}
	}
}

// ������
function InsertCols( oTable ) {
	if ( oTable ) {
		for(var i=0; i<oTable.rows.length; i++){
			var elCell = oTable.rows[i].insertCell();
			elCell.innerHTML = "&nbsp;"
		}
	}
}

// ɾ����
function DeleteRows( oTable ) {
	if ( oTable ) {
		oTable.deleteRow();
	}
}

// ɾ����
function DeleteCols( oTable ) {
	if ( oTable ) {
		for(var i=0;i<oTable.rows.length;i++){
			oTable.rows[i].deleteCell();
		}
	}
}

</script>

<SCRIPT event=onclick for=Ok language=JavaScript>
	// �߿���ɫ����Ч��
	sBorderColor = t_bordercolor.value;
	if (!IsColor(sBorderColor)){
		BaseAlert(t_bordercolor,'��Ч�ı߿���ɫֵ��');
		return;
	}
	// ������ɫ����Ч��
	sBgColor = t_bgcolor.value;
	if (!IsColor(sBgColor)){
		BaseAlert(t_bgcolor,'��Ч�ı�����ɫֵ��');
		return;
	}
	// ��������Ч��
	if (!MoreThanOne(d_row,'��Ч������������Ҫ1�У�')) return;
	// ��������Ч��
	if (!MoreThanOne(d_col,'��Ч������������Ҫ1�У�')) return;
	// ���ߴ�ϸ����Ч��
	if (d_border.value == "") d_border.value = "0";
	if (d_cellpadding.value == "") d_cellpadding.value = "0";
	if (d_cellspacing.value == "") d_cellspacing.value = "0";
	// ȥǰ��0
	d_border.value = parseFloat(d_border.value);
	d_cellpadding.value = parseFloat(d_cellpadding.value);
	d_cellspacing.value = parseFloat(d_cellspacing.value);
	// �����Чֵ��
	var sWidth = "";
	if (d_widthcheck.checked){
		if (!MoreThanOne(d_widthvalue,'��Ч�ı���ȣ�')) return;
		sWidth = d_widthvalue.value + d_widthunit.value;
	}
	// �߶���Чֵ��
	var sHeight = "";
	if (d_heightcheck.checked){
		if (!MoreThanOne(d_heightvalue,'��Ч�ı��߶ȣ�')) return;
		sHeight = d_heightvalue.value + d_heightunit.value;
	}

	sRow = d_row.value;
	sCol = d_col.value;
	sAlign = d_align.options[d_align.selectedIndex].value;
	sBorder = d_border.value;
	sCellPadding = d_cellpadding.value;
	sCellSpacing = d_cellspacing.value;
	sImage = d_image.value;
	sRepeat = d_repeat.value;
	sAttachment = d_attachment.value;
	sBorderStyle = d_borderstyle.options[d_borderstyle.selectedIndex].value;
	if (sImage!="") {
		sImage = "url(" + sImage + ")";
	}

	if (sAction == "MODI") {
		// �޸�����
		var xCount = sRow - oControl.rows.length;
  		if (xCount > 0)
	  		for (var i = 0; i < xCount; i++) InsertRows(oControl);
  		else
	  		for (var i = 0; i > xCount; i--) DeleteRows(oControl);
		// �޸�����
  		var xCount = sCol - getColCount(oControl);
  		if (xCount > 0)
  			for (var i = 0; i < xCount; i++) InsertCols(oControl);
  		else
  			for (var i = 0; i > xCount; i--) DeleteCols(oControl);

		try {
			oControl.width = sWidth;
			oControl.style.width = sWidth;
		}
		catch(e) {
			//alert("�Բ�������������Ч�Ŀ��ֵ��\n���磺90%  200  300px  10cm��");
		}
		try {
			oControl.height = sHeight;
			oControl.style.height = sHeight;
		}
		catch(e) {
			//alert("�Բ�������������Ч�ĸ߶�ֵ��\n���磺90%  200  300px  10cm��");
		}

		oControl.align			= sAlign;
  		oControl.border			= sBorder;
  		oControl.cellSpacing	= sCellSpacing;
  		oControl.cellPadding	= sCellPadding;
  		oControl.borderColor	= sBorderColor;
  		oControl.bgColor		= sBgColor;
		oControl.style.backgroundImage = sImage;
		oControl.style.backgroundRepeat = sRepeat;
		oControl.style.backgroundAttachment = sAttachment;
		oControl.style.borderStyle = sBorderStyle;

	}else{
		var sTable = "<table align='"+sAlign+"' border='"+sBorder+"' cellpadding='"+sCellPadding+"' cellspacing='"+sCellSpacing+"' width='"+sWidth+"' height='"+sHeight+"' bordercolor='"+sBorderColor+"' bgcolor='"+sBgColor+"' style='background-image:"+sImage+";background-repeat:"+sRepeat+";background-attachment:"+sAttachment+";border-style:"+sBorderStyle+";'>";
		for (var i=1;i<=sRow;i++){
			sTable = sTable + "<tr>";
			for (var j=1;j<=sCol;j++){
				sTable = sTable + "<td>&nbsp;</td>";
			}
			sTable = sTable + "</tr>";
		}
		sTable = sTable + "</table>";
		dialogArguments.insertHTML(sTable);
	}

	window.returnValue = null;
	window.close();
</SCRIPT>

</head>
<body bgColor=#D4D0C8 onload="InitDocument()">

<table border=0 cellpadding=0 cellspacing=0 align=center>
<tr>
	<td>
	<fieldset>
	<legend>����С</legend>
	<table border=0 cellpadding=0 cellspacing=0>
	<tr><td colspan=9 height=5></td></tr>
	<tr>
		<td width=7></td>
		<td>�������:</td>
		<td width=5></td>
		<td><input type=text id=d_row size=10 value="" ONKEYPRESS="event.returnValue=IsDigit();" maxlength=3></td>
		<td width=40></td>
		<td>�������:</td>
		<td width=5></td>
		<td><input type=text id=d_col size=10 value="" ONKEYPRESS="event.returnValue=IsDigit();" maxlength=3></td>
		<td width=7></td>
	</tr>
	<tr><td colspan=9 height=5></td></tr>
	</table>
	</fieldset>
	</td>
</tr>
<tr><td height=5></td></tr>
<tr>
	<td>
	<fieldset>
	<legend>��񲼾�</legend>
	<table border=0 cellpadding=0 cellspacing=0>
	<tr><td colspan=9 height=5></td></tr>
	<tr>
		<td width=7></td>
		<td>���뷽ʽ:</td>
		<td width=5></td>
		<td>
			<select id="d_align" style="width:72px">
			<option value=''>Ĭ��</option>
			<option value='left'>�����</option>
			<option value='center'>����</option>
			<option value='right'>�Ҷ���</option>
			</select>
		</td>
		<td width=40></td>
		<td>�߿��ϸ:</td>
		<td width=5></td>
		<td><input type=text id=d_border size=10 value="" ONKEYPRESS="event.returnValue=IsDigit();"></td>
		<td width=7></td>
	</tr>
	<tr><td colspan=9 height=5></td></tr>
	<tr>
		<td width=7></td>
		<td>��Ԫ���:</td>
		<td width=5></td>
		<td><input type=text id=d_cellspacing size=10 value="" ONKEYPRESS="event.returnValue=IsDigit();" maxlength=3></td>
		<td width=40></td>
		<td>��Ԫ�߾�:</td>
		<td width=5></td>
		<td><input type=text id=d_cellpadding size=10 value="" ONKEYPRESS="event.returnValue=IsDigit();" maxlength=3></td>
		<td width=7></td>
	</tr>
	<tr><td colspan=9 height=5></td></tr>
	</table>
	</fieldset>
	</td>
</tr>
<tr><td height=5></td></tr>
<tr>
	<td>
	<fieldset>
	<legend>���ߴ�</legend>
	<table border=0 cellpadding=0 cellspacing=0 width='100%'>
	<tr><td colspan=9 height=5></td></tr>
	<tr>
		<td width=7></td>
		<td onclick="d_widthcheck.click()" noWrap valign=middle><input id="d_widthcheck" type="checkbox" onclick="d_widthvalue.disabled=(!this.checked);d_widthunit.disabled=(!this.checked);" value="1"> ָ�����Ŀ��</td>
		<td align=right width="60%">
			<input name="d_widthvalue" type="text" value="" size="5" ONKEYPRESS="event.returnValue=IsDigit();" maxlength="4">
			<select name="d_widthunit">
			<option value='px'>����</option><option value='%'>�ٷֱ�</option>
			</select>
		</td>
		<td width=7></td>
	</tr>
	<tr><td colspan=9 height=5></td></tr>
	<tr>
		<td height=7></td>
		<td onclick="d_heightcheck.click()" noWrap valign=middle><input id="d_heightcheck" type="checkbox" onclick="d_heightvalue.disabled=(!this.checked);d_heightunit.disabled=(!this.checked);" value="1"> ָ�����ĸ߶�</td>
		<td align=right height="60%">
			<input name="d_heightvalue" type="text" value="" size="5" ONKEYPRESS="event.returnValue=IsDigit();" maxlength="4">
			<select name="d_heightunit">
			<option value='px'>����</option><option value='%'>�ٷֱ�</option>
			</select>
		</td>
		<td width=7></td>
	</tr>
	<tr><td colspan=9 height=5></td></tr>
	</table>
	</fieldset>
	</td>
</tr>
<tr><td height=5></td></tr>
<tr>
	<td>
	<fieldset>
	<legend>�����ʽ</legend>
	<table border=0 cellpadding=0 cellspacing=0>
	<tr><td colspan=9 height=5></td></tr>
	<tr>
		<td width=7></td>
		<td>�߿���ɫ:</td>
		<td width=5></td>
		<td><input type=text id=t_bordercolor size=7 value=""></td>
		<td><img border=0 src="images/rect.gif" width=18 style="cursor:hand" id=s_bordercolor onclick="SelectColor('bordercolor')"></td>
		<td width=40></td>
		<td>�߿���ʽ:</td>
		<td width=5></td>
		<td colspan=2>
			<select id=d_borderstyle size=1 style="width:72px">
			<option value="">Ĭ��</option>
			<option value="solid">ʵ��</option>
			<option value="dotted">����</option>
			<option value="dashed">���ۺ�</option>
			<option value="double">˫��</option>
			<option value="groove">����</option>
			<option value="ridge">͹��</option>
			<option value="inset">Ƕ��</option>
			<option value="outset">����</option>
			</select>
		</td>
		<td width=7></td>
	</tr>
	<tr><td colspan=9 height=5></td></tr>
	<tr>
		<td width=7></td>
		<td>������ɫ:</td>
		<td width=5></td>
		<td><input type=text id=t_bgcolor size=7 value=""></td>
		<td><img border=0 src="images/rect.gif" width=18 style="cursor:hand" id=s_bgcolor onclick="SelectColor('bgcolor')"></td>
		<td width=40></td>
		<td>����ͼƬ:</td>
		<td width=5></td>
		<td colspan='2'><input type=text id=d_image size=10 value=""><input type=hidden id=d_repeat><input type=hidden id=d_attachment></td>
		<td width=7></td>
	</tr>
	<tr><td colspan=9 height=5></td></tr>
	</table>
	</fieldset>
	</td>
</tr>
<tr><td height=5></td></tr>
<tr><td align=right><input type=submit value='  ȷ��  ' id=Ok>&nbsp;&nbsp;<input type=button value='  ȡ��  ' onclick="window.close();"></td></tr>
</table>
</body>
</html>