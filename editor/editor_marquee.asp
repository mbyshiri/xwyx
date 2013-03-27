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
<HTML>
<HEAD>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<style>
BODY {PADDING:5PX}
TD,BODY,SELECT,P,INPUT {FONT-SIZE:9PT}
</style>
<script language=javascript>
var sAction = "INSERT";
var sTitle = "插入";
var el;
var sText = "";
var sBehavior = "";
document.write("<title>滚动文本（" + sTitle + "）</title>");


// 单选的点击事件
function check(){
	sBehavior = event.srcElement.value;
}

// 初始值
function InitDocument() {
	d_text.value = sText;
	switch (sBehavior) {
	case "scroll":
		document.all("d_behavior")[0].checked = true;
		break;
	case "slide":
		document.all("d_behavior")[1].checked = true;
		break;
	default:
		sBehavior = "alternate";
		document.all("d_behavior")[2].checked = true;
		break;
	}

}
</script>


<SCRIPT event=onclick for=Ok language=JavaScript>
	sText = d_text.value;
	if (sAction == "MODI") {
		el.behavior = sBehavior;
		el.innerHTML = sText;
	}else{
              var str1;
              str1="<marquee behavior='"+sBehavior+"'>"+sText+"</marquee>"
	}
              window.returnValue = str1
              window.close();
</script>
</HEAD>

<body bgColor=#D4D0C8 onload="InitDocument()">

<table border=0 cellpadding=0 cellspacing=0 align=center>
  <tr>
   <td>
     <FIELDSET align=left>
	<LEGEND><b>插入滚动文本</b></LEGEND>
	<table width="335" border=0 cellpadding=0 cellspacing=5>
	  <tr valign=middle>
	   <td width="37">文本:</td><td width="191"><input type=text id="d_text" size=30 value=""></td>
	   <td width="87" rowspan="2" align="center"><input type=submit value='  确定  ' id=Ok>
            <br>
            <br>	    
            <input type=button value='  取消  ' onClick="window.close();"></td>
	  </tr>
	  <tr valign=middle><td>表现:</td><td><input onclick="check()" type="radio" name="d_behavior" value="scroll"> 滚动条 <input onclick="check()" type="radio" name="d_behavior" value="slide"> 幻灯片 <input onclick="check()" type="radio" name="d_behavior" value="alternate"> 交替</td>
	  </tr>
	</table>
	</FIELDSET>
   </td>
  </tr>
</table>

</body>
</html>
