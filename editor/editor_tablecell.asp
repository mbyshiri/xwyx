<%@language=vbscript codepage=936 %>
<%
option explicit
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
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
body, a, table, div, span, td, th, input, select{font:9pt;font-family: "宋体", Verdana, Arial, Helvetica, sans-serif;}
body {padding:5px}
</style>

<script language="JavaScript">
// 取通过URL传过来的参数 (格式如 ?Param1=Value1&Param2=Value2)
var URLParams = new Object() ;
var aParams = document.location.search.substr(1).split('&') ;
for (i=0 ; i < aParams.length ; i++) {
	var aParam = aParams[i].split('=') ;
	URLParams[aParam[0]] = aParam[1] ;
}
var sAction = URLParams['action'] ;
var sTitle = "";

var oControl;
var oSeletion;
var sRangeType;

var sAlign = "";
var sVAlign = "";
var sWidth = "";
var sHeight = "";
var sBorderColor = "#000000";
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

if (sAction == "row"){
	oControl = getParentObject(oSelection.parentElement(), "TR");
	sAction = "ROW";
	sTitle = "表格行";
}else{
	oControl = getParentObject(oSelection.parentElement(), "TD");
	sAction = "CELL";
	sTitle = "单元格";
}
if (oControl){
	sAlign = oControl.align;
	sVAlign = oControl.vAlign;
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
//=================================================
//过程名：SelectColor
//作  用：显示颜色表
//参  数：what  --- 要获得颜色的参数
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
// 是否有效颜色值
function IsColor(color){
	var temp=color;
	if (temp=="") return true;
	if (temp.length!=7) return false;
	return (temp.search(/\#[a-fA-F0-9]{6}/) != -1);
}
// 只允许输入数字
function IsDigit(){
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}

// 搜索下拉框值与指定值匹配，并选择匹配项
function SearchSelectValue(o_Select, s_Value){
	for (var i=0;i<o_Select.length;i++){
		if (o_Select.options[i].value == s_Value){
			o_Select.selectedIndex = i;
			return true;
		}
	}
	return false;
}
// 返回指定标签的父对象
function getParentObject(obj, tag){
	if (tag == "TD"){
		while(obj!=null && obj.tagName!=tag && obj.tagName!="TH"){
			obj=obj.parentElement;
		}
	}else{
		while(obj!=null && obj.tagName!=tag){
			obj=obj.parentElement;
		}
	}
	return obj;
}

document.write("<title>" + sTitle + "属性</title>");

// 初始值
function InitDocument(){
	SearchSelectValue(d_align, sAlign.toLowerCase());
	SearchSelectValue(d_valign, sVAlign.toLowerCase());
	SearchSelectValue(d_borderstyle, sBorderStyle.toLowerCase());

	// 修改状态时取值
	if ((sWidth == "")||(sWidth==undefined)){
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

// 判断值是否大于0
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

</script>

<SCRIPT event=onclick for=Ok language=JavaScript>
	// 边框颜色的有效性
	sBorderColor = t_bordercolor.value;
	if (!IsColor(sBorderColor)){
		BaseAlert(d_bordercolor,'无效的边框颜色值！');
		return;
	}
	// 背景颜色的有效性
	sBgColor = t_bgcolor.value;
	if (!IsColor(sBgColor)){
		BaseAlert(d_bgcolor,'无效的背景颜色值！');
		return;
	}
	// 宽度有效值性
	var sWidth = "";
	if (d_widthcheck.checked){
		if (!MoreThanOne(d_widthvalue,'无效的宽度！')) return;
		sWidth = d_widthvalue.value + d_widthunit.value;
	}
	// 高度有效值性
	var sHeight = "";
	if (d_heightcheck.checked){
		if (!MoreThanOne(d_heightvalue,'无效的高度！')) return;
		sHeight = d_heightvalue.value + d_heightunit.value;
	}

	sAlign = d_align.options[d_align.selectedIndex].value;
	sVAlign = d_valign.options[d_valign.selectedIndex].value;
	sImage = d_image.value;
	sRepeat = d_repeat.value;
	sAttachment = d_attachment.value;
	sBorderStyle = d_borderstyle.options[d_borderstyle.selectedIndex].value;
	if (sImage!="") {
		sImage = "url(" + sImage + ")";
	}

	if (oControl) {
		try {
			oControl.width = sWidth;
		}
		catch(e) {
			//alert("对不起，请您输入有效的宽度值！\n（如：90%  200  300px  10cm）");
		}
		try {
			oControl.height = sHeight;
		}
		catch(e) {
			//alert("对不起，请您输入有效的高度值！\n（如：90%  200  300px  10cm）");
		}

		oControl.align			= sAlign;
		oControl.vAlign			= sVAlign;
  		oControl.borderColor	= sBorderColor;
  		oControl.bgColor		= sBgColor;
		oControl.style.backgroundImage = sImage;
		oControl.style.backgroundRepeat = sRepeat;
		oControl.style.backgroundAttachment = sAttachment;
		oControl.style.borderStyle = sBorderStyle;
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
	<legend>布局</legend>
	<table border=0 cellpadding=0 cellspacing=0>
	<tr><td colspan=9 height=5></td></tr>
	<tr>
		<td width=7></td>
		<td>水平对齐:</td>
		<td width=5></td>
		<td>
			<select id="d_align" style="width:72px">
			<option value=''>默认</option>
			<option value='left'>左对齐</option>
			<option value='right'>右对齐</option>
			<option value='center'>水平居中</option>
			<option value='right'>两端对齐</option>
			</select>
		</td>
		<td width=40></td>
		<td>垂直对齐:</td>
		<td width=5></td>
		<td>
			<select id="d_valign" style="width:72px">
			<option value=''>默认</option>
			<option value='top'>顶边对齐</option>
			<option value='middle'>垂直居中</option>
			<option value='baseline'>基线</option>
			<option value='bottom'>底边对齐</option>
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
	<legend>尺寸</legend>
	<table border=0 cellpadding=0 cellspacing=0 width='100%'>
	<tr><td colspan=9 height=5></td></tr>
	<tr>
		<td width=7></td>
		<td onclick="d_widthcheck.click()" noWrap valign=middle><input id="d_widthcheck" type="checkbox" onclick="d_widthvalue.disabled=(!this.checked);d_widthunit.disabled=(!this.checked);" value="1"> 指定宽度</td>
		<td align=right width="60%">
			<input name="d_widthvalue" type="text" value="" size="5" ONKEYPRESS="event.returnValue=IsDigit();" maxlength="4">
			<select name="d_widthunit">
			<option value='px'>像素</option><option value='%'>百分比</option>
			</select>
		</td>
		<td width=7></td>
	</tr>
	<tr><td colspan=9 height=5></td></tr>
	<tr>
		<td height=7></td>
		<td onclick="d_heightcheck.click()" noWrap valign=middle><input id="d_heightcheck" type="checkbox" onclick="d_heightvalue.disabled=(!this.checked);d_heightunit.disabled=(!this.checked);" value="1"> 指定高度</td>
		<td align=right height="60%">
			<input name="d_heightvalue" type="text" value="" size="5" ONKEYPRESS="event.returnValue=IsDigit();" maxlength="4">
			<select name="d_heightunit">
			<option value='px'>像素</option><option value='%'>百分比</option>
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
	<legend>样式</legend>
	<table border=0 cellpadding=0 cellspacing=0>
	<tr><td colspan=9 height=5></td></tr>
	<tr>
		<td width=7></td>
		<td>边框颜色:</td>
		<td width=5></td>
		<td><input type=text id=t_bordercolor size=7 value=""></td>
		<td><img border=0 src="images/rect.gif" width=18 style="cursor:hand" id=s_bordercolor onclick="SelectColor('bordercolor')"></td>
		<td width=40></td>
		<td>边框样式:</td>
		<td width=5></td>
		<td colspan=2>
			<select id=d_borderstyle size=1 style="width:70px">
			<option value="">默认</option>
			<option value="solid">实线</option>
			<option value="dotted">虚线</option>
			<option value="dashed">破折号</option>
			<option value="double">双线</option>
			<option value="groove">凹线</option>
			<option value="ridge">凸线</option>
			<option value="inset">嵌入</option>
			<option value="outset">开端</option>
			</select>
		</td>
		<td width=7></td>
	</tr>
	<tr><td colspan=9 height=5></td></tr>
	<tr>
		<td width=7></td>
		<td>背景颜色:</td>
		<td width=5></td>
		<td><input type=text id=t_bgcolor size=7 value=""></td>
		<td><img border=0 src="images/rect.gif" width=18 style="cursor:hand" id=s_bgcolor onclick="SelectColor('bgcolor')"></td>
		<td width=40></td>
		<td>背景图片:</td>
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
<tr><td align=right><input type=submit value='  确定  ' id=Ok>&nbsp;&nbsp;<input type=button value='  取消  ' onclick="window.close();"></td></tr>
</table>
</body>
</html>