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
<title>插入水平线</title>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link rel="stylesheet" type="text/css" href="editor_dialog.css">
<script language="JavaScript">
function OK(){
  var str1;
  str1="<hr color='"+t_color.value+"' size="+size.value+"' "+shadetype.value+" align="+align.value+" width="+width.value+">"
  window.returnValue = str1
  window.close();
}
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
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
</script>
</head>
<BODY bgColor=#D4D0C8 topmargin=15 leftmargin=15 >
<table width=100% border="0" cellpadding="0" cellspacing="2">
  <tr><td>
<FIELDSET align=left>
<LEGEND align=left><strong>输入水平线参数</strong></LEGEND>
      <table border="0" cellpadding="0" cellspacing="3">
        <tr> 
          <td>线条颜色：
            <input name="t_color" id=t_color  size="7" maxlength="7">
	    <img border=0 src="images/rect.gif" width=18 style="cursor:hand" id=s_color onclick="SelectColor('color')">
          </td>

        </tr>
        <tr>
          <td>线条粗度：
            <input name="size"  id=size onKeyPress="event.returnValue=IsDigit();" value="2" size="4" maxlength=3>
必须是数字，范围建议在1-100之间</td>
        </tr>
        <tr> 
          <td> 页面对齐：
            <select name="align"  id=align>
              <option value="left" selected>默认对齐</option>
              <option value="left">左对齐 </option>
              <option value="center">中对齐 </option>
              <option value="right">右对齐 </option>
            </select>
            &nbsp;&nbsp;阴影效果；
            <select name="shadetype"  id=shadetype>
              <option value=noshade selected>无 
              <option value=''>有 
            </select>
          </td>
        </tr>
        <tr> 
          <td> 水平宽度：
            <input name="width" id=width ONKEYPRESS="event.returnValue=IsDigit();" value="400" size="6" maxlength=3>
            必须是数字，范围建议在1-999之间</td>
        </tr>
      </table>
</fieldset></td>
    <td width=80 align="center"><input name="cmdOK" type="button" id="cmdOK" value="  确定  " onClick="OK();">
      <br>
      <br>
      <input name="cmdCancel" type=button id="cmdCancel" onclick="window.close();" value="  取消  "></td>
  </tr></table>
</body>
</html>