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
<HTML>
<HEAD>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>

<STYLE type=text/css>
TD {FONT-SIZE: 10.8pt}
BODY {FONT-SIZE: 10.8pt}
BUTTON {WIDTH: 5em}
</STYLE>



<SCRIPT language=JavaScript>

var URLParams = new Object() ;
var sAction = URLParams['action'] ;
var sTitle = "";
var color = "" ;
var oSelection;
var oControl;
var sRangeType;

// �Ƿ���Ч��ɫֵ
function IsColor(color){
	var temp=color;
	if (temp=="") return true;
	if (temp.length!=7) return false;
	return (temp.search(/\#[a-fA-F0-9]{6}/) != -1);
}

switch (sAction) {
	case "forecolor":	// ����ǰ��ɫ
		sTitle = "����ǰ��ɫ";
		oSelection = dialogArguments.eWebEditor.document.selection.createRange();
		color = oSelection.queryCommandValue("ForeColor");
		if (color) color = N2Color(color);
		break;
	case "backcolor":	// ���屳��ɫ
		sTitle = "���屳��ɫ";
		oSelection = dialogArguments.eWebEditor.document.selection.createRange();
		color = oSelection.queryCommandValue("BackColor");
		if (color) color = N2Color(color);
		break;
	case "bgcolor":		// ���󱳾�ɫ
		sTitle = "���󱳾�ɫ";
		oSelection = dialogArguments.eWebEditor.document.selection.createRange();
		sRangeType = dialogArguments.eWebEditor.document.selection.type;
		if (sRangeType == "Control") {
			oControl = GetControl(oSelection, "TABLE");
		}else{
			oControl = GetParent(oSelection.parentElement());
		}
		if (oControl) {
			switch(oControl.tagName){
			case "TD":
				sTitle += " - ��Ԫ��";
				break;
			case "TR":
			case "TH":
				sTitle += " - ������";
				break;
			default:
				sTitle += " - ����";
				break;
			}
			color = oControl.bgColor;
		}else{
			sTitle += " - ��ҳ";
		}
		break;
	default:			// ������ɫ��
		if (URLParams['color']){
			color = decodeURIComponent(URLParams['color']) ;
		}
		break;
}

document.write("<TITLE>��ɫѡ��" + sTitle + "��</TITLE>");

// Ĭ����ʾֵ
if (!color) color = "#000000";

// �����б�����ɫ���ԵĶ���
function GetParent(obj){
	while(obj!=null && obj.tagName!="TD" && obj.tagName!="TR" && obj.tagName!="TH" && obj.tagName!="TABLE")
		obj=obj.parentElement;
	return obj;
}

// ���ر�ǩ����ѡ���ؼ�
function GetControl(obj, sTag){
	obj=obj.item(0);
	if (obj.tagName==sTag){
		return obj;
	}
	return null;
}

// ��ֵתΪRGB16������ɫ��ʽ
function N2Color(s_Color){
	s_Color = s_Color.toString(16);
	switch (s_Color.length) {
	case 1:
		s_Color = "0" + s_Color + "0000"; 
		break;
	case 2:
		s_Color = s_Color + "0000";
		break;
	case 3:
		s_Color = s_Color.substring(1,3) + "0" + s_Color.substring(0,1) + "00" ;
		break;
	case 4:
		s_Color = s_Color.substring(2,4) + s_Color.substring(0,2) + "00" ;
		break;
	case 5:
		s_Color = s_Color.substring(3,5) + s_Color.substring(1,3) + "0" + s_Color.substring(0,1) ;
		break;
	case 6:
		s_Color = s_Color.substring(4,6) + s_Color.substring(2,4) + s_Color.substring(0,2) ;
		break;
	default:
		s_Color = "";
	}
	return '#' + s_Color;
}

// ��ʼֵ
function InitDocument(){
	ShowColor.bgColor = color;
	RGB.innerHTML = color;
	SelColor.value = color;
}


var SelRGB = color;
var DrRGB = '';
var SelGRAY = '120';

var hexch = new Array('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F');

function ToHex(n) {	
	var h, l;

	n = Math.round(n);
	l = n % 16;
	h = Math.floor((n / 16)) % 16;
	return (hexch[h] + hexch[l]);
}

function DoColor(c, l){
	var r, g, b;

	r = '0x' + c.substring(1, 3);
	g = '0x' + c.substring(3, 5);
	b = '0x' + c.substring(5, 7);

	if(l > 120){
		l = l - 120;

		r = (r * (120 - l) + 255 * l) / 120;
		g = (g * (120 - l) + 255 * l) / 120;
		b = (b * (120 - l) + 255 * l) / 120;
	}else{
		r = (r * l) / 120;
		g = (g * l) / 120;
		b = (b * l) / 120;
	}

	return '#' + ToHex(r) + ToHex(g) + ToHex(b);
}

function EndColor(){
	var i;

	if(DrRGB != SelRGB){
		DrRGB = SelRGB;
		for(i = 0; i <= 30; i ++)
		GrayTable.rows(i).bgColor = DoColor(SelRGB, 240 - i * 8);
	}

	SelColor.value = DoColor(RGB.innerText, GRAY.innerText);
	ShowColor.bgColor = SelColor.value;
}
</SCRIPT>

<SCRIPT event=onclick for=ColorTable language=JavaScript>
	SelRGB = event.srcElement.bgColor;
	EndColor();
</SCRIPT>

<SCRIPT event=onmouseover for=ColorTable language=JavaScript>
	RGB.innerText = event.srcElement.bgColor;
	EndColor();
</SCRIPT>

<SCRIPT event=onmouseout for=ColorTable language=JavaScript>
	RGB.innerText = SelRGB;
	EndColor();
</SCRIPT>

<SCRIPT event=onclick for=GrayTable language=JavaScript>
	SelGRAY = event.srcElement.title;
	EndColor();
</SCRIPT>

<SCRIPT event=onmouseover for=GrayTable language=JavaScript>
	GRAY.innerText = event.srcElement.title;
	EndColor();
</SCRIPT>

<SCRIPT event=onmouseout for=GrayTable language=JavaScript>
	GRAY.innerText = SelGRAY;
	EndColor();
</SCRIPT>

<SCRIPT event=onclick for=Ok language=JavaScript>
	color = SelColor.value;
	if (!IsColor(color)){
		alert('��Ч����ɫֵ��');
		return;
	}

	switch (sAction) {
		case "forecolor":
			dialogArguments.format('ForeColor', color) ;
			window.returnValue = null;
			break;
		case "backcolor":
			dialogArguments.format('BackColor', color) ;
			window.returnValue = null;
			break;
		case "bgcolor":
			if (oControl){
				oControl.bgColor = color;
			}else{
				dialogArguments.setHTML("<table border=0 cellpadding=0 cellspacing=0 width='100%' height='100%'><tr><td valign=top bgcolor='"+color+"'>"+dialogArguments.getHTML()+"</td></tr></table>");
			}
			window.returnValue = null;
			break;
		default:
			window.returnValue = color;
			break;
	}
	window.close();
</SCRIPT>

</HEAD>

<BODY bgColor=#D4D0C8 onload="InitDocument()">
<DIV align=center>
<CENTER>
  <TABLE border=0 cellPadding=0 cellSpacing=10>
   <TBODY>
    <TR>
     <TD>
      <TABLE border=0 cellPadding=0 cellSpacing=0 id=ColorTable style="CURSOR: hand">
        <SCRIPT language=JavaScript>
          function wc(r, g, b, n){
            r = ((r * 16 + r) * 3 * (15 - n) + 0x80 * n) / 15;
	    g = ((g * 16 + g) * 3 * (15 - n) + 0x80 * n) / 15;
	    b = ((b * 16 + b) * 3 * (15 - n) + 0x80 * n) / 15;
             document.write('<TD BGCOLOR=#' + ToHex(r) + ToHex(g) + ToHex(b) + ' height=8 width=8></TD>');
          }

         var cnum = new Array(1, 0, 0, 1, 1, 0, 0, 1, 0, 0, 1, 1, 0, 0, 1, 1, 0, 1, 1, 0, 0);

        for(i = 0; i < 16; i ++){
	   document.write('<TR>');
	   for(j = 0; j < 30; j ++){
		n1 = j % 5;
		n2 = Math.floor(j / 5) * 3;
		n3 = n2 + 3;

		wc((cnum[n3] * n1 + cnum[n2] * (5 - n1)),
		(cnum[n3 + 1] * n1 + cnum[n2 + 1] * (5 - n1)),
		(cnum[n3 + 2] * n1 + cnum[n2 + 2] * (5 - n1)), i);
     	   }
          document.writeln('</TR>');
         }
       </SCRIPT>
       <TBODY></TBODY>
      </TABLE>
     </TD>
     <TD>
      <TABLE border=0 cellPadding=0 cellSpacing=0 id=GrayTable style="CURSOR: hand">
        <SCRIPT language=JavaScript>
          for(i = 255; i >= 0; i -= 8.5)
            document.write('<TR BGCOLOR=#' + ToHex(i) + ToHex(i) + ToHex(i) + '><TD TITLE=' + Math.floor(i * 16 / 17) + ' height=4 width=20></TD></TR>');
        </SCRIPT>
        <TBODY></TBODY>
      </TABLE>
     </TD>
   </TR>
 </TBODY>
</TABLE>
</CENTER>
</DIV>
<DIV align=center>
<CENTER>
 <TABLE border=0 cellPadding=0 cellSpacing=10>
  <TBODY>
   <TR>
     <TD align=middle rowSpan=2>ѡ��ɫ��
      <TABLE border=1 cellPadding=0 cellSpacing=0 height=30 id=ShowColor width=40 bgcolor="">
       <TBODY>
        <TR><TD></TD></TR>
       </TBODY>
      </TABLE>
     </TD>
     <TD rowSpan=2>��ɫ: <SPAN id=RGB></SPAN><BR>����: <SPAN id=GRAY>120</SPAN><BR>����: <INPUT id=SelColor size=7 value=""></TD>
     <TD><BUTTON id=Ok type=submit>ȷ��</BUTTON></TD></TR>
   <TR>
    <TD><BUTTON onclick=window.close();>ȡ��</BUTTON></TD>
   </TR>
  </TBODY>
 </TABLE>
</CENTER>
</DIV>
</BODY>
</HTML>