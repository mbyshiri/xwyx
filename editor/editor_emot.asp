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
<TITLE>插入表情图标</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>
body, a, table, div, span, td, th, input, select{font:9pt;font-family: "宋体", Verdana, Arial, Helvetica, sans-serif;}
body {padding:5px}
table.content {background-color:#000000;width:100%;}
table.content td {background-color:#ffffff;width:18px;height:18px;text-align:center;vertical-align:middle;cursor:hand;}
.card {cursor:hand;background-color:#3A6EA5;text-align:center;}
</STYLE>
<SCRIPT language=JavaScript>

// 选项卡点击事件
function cardClick(cardID){
	var obj;
	for (var i=1;i<3;i++){
		obj=document.all("card"+i);
		obj.style.backgroundColor="#3A6EA5";
		obj.style.color="#FFFFFF";
	}
	obj=document.all("card"+cardID);
	obj.style.backgroundColor="#FFFFFF";
	obj.style.color="#3A6EA5";

	for (var i=1;i<3;i++){
		obj=document.all("content"+i);
		obj.style.display="none";
	}
	obj=document.all("content"+cardID);
	obj.style.display="";
}
// 点击返回
function SymbolClick(){
  var str1;
  str1="<IMG SRC=" + event.srcElement.src + ">"
  window.returnValue = str1
  window.close();
}

</script>
</HEAD>

<BODY bgColor=#D4D0C8>

<table border=0 cellpadding=0 cellspacing=0><tr valign=top><td>
<fieldset><legend><b>插入小图标</b></legend><br><table border=0 cellpadding=3 cellspacing=0>
<tr align=center>
	<td class="card" onclick="cardClick(1)" id="card1">表情</td>
	<td width=2></td>
	<td class="card" onclick="cardClick(2)" id="card2">心情</td>
</tr>
<tr>
	<td bgcolor=#ffffff align=center valign=middle colspan=11>
	<table border=0 cellpadding=3 cellspacing=1 class="content" id="content1">
	<%
	dim i,s

	For i=0 to 4
	    response.write "<tr>"
		For s=0 to 9
	      Response.write"<td onmouseover=""javascript:yulantu.src='Images/emot/" & i & s &".gif'""><IMG src='Images/emot/" & i & s &".gif' onclick=""SymbolClick()""></td>"
		Next
		Response.write "</tr>"
	Next
	
	%>
	</table>
	<table border=0 cellpadding=3 cellspacing=1 class="content" id="content2">
      <tr>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot1.gif'"><IMG src='Images/emot2/emot1.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot2.gif'"><IMG src='Images/emot2/emot2.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot3.gif'"><IMG src='Images/emot2/emot3.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot4.gif'"><IMG src='Images/emot2/emot4.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot5.gif'"><IMG src='Images/emot2/emot5.gif' onclick="SymbolClick()"></td>
	  </tr>
      <tr>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot6.gif'"><IMG src='Images/emot2/emot6.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot7.gif'"><IMG src='Images/emot2/emot7.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot8.gif'"><IMG src='Images/emot2/emot8.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot9.gif'"><IMG src='Images/emot2/emot9.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot10.gif'"><IMG src='Images/emot2/emot10.gif' onclick="SymbolClick()"></td>
	  </tr>
	  <tr>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot11.gif'"><IMG src='Images/emot2/emot11.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot12.gif'"><IMG src='Images/emot2/emot12.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot13.gif'"><IMG src='Images/emot2/emot13.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot14.gif'"><IMG src='Images/emot2/emot14.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot15.gif'"><IMG src='Images/emot2/emot15.gif' onclick="SymbolClick()"></td>
	  </tr>
      <tr>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot16.gif'"><IMG src='Images/emot2/emot16.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot17.gif'"><IMG src='Images/emot2/emot17.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot18.gif'"><IMG src='Images/emot2/emot18.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot19.gif'"><IMG src='Images/emot2/emot19.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot20.gif'"><IMG src='Images/emot2/emot20.gif' onclick="SymbolClick()"></td>
	  </tr>
	  
	  <tr>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot21.gif'"><IMG src='Images/emot2/emot21.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot22.gif'"><IMG src='Images/emot2/emot22.gif' onclick="SymbolClick()"></td>
		<td onmouseover="javascript:yulantu.src='Images/emot2/emot23.gif'"><IMG src='Images/emot2/emot23.gif' onclick="SymbolClick()"></td>
		<td ></td>
		<td ></td>
	  </tr>
	</table>
   </td>
  </tr>
 </table>
</fieldset>
</td><td width=10></td><td>
<table border=0 cellpadding=0 cellspacing=0>
  <tr><td height=25></td></tr>
  <tr><td align=center>预览</td></tr>
  <tr><td height=10></td></tr>
  <tr>
    <td align=center valign=middle>
      <table border=0 cellpadding=0 cellspacing=1 bgcolor=#000000>
        <tr>
	  <td bgcolor=#ffffff style="font-size:32px;color:#0000ff"  align=center valign=middle width=50 height=50>
	   <IMG SRC="Images/emot/08.gif" id=yulantu BORDER='0'ALT=''>
	  </td>
	</tr>
      </table>
    </td>
  </tr>
  <tr><td height=52></td></tr>
  <tr><td align=center><input type=button value='  取消  ' onclick="window.close();"></td></tr>
</table>
</td></tr></table>
<script language=javascript>
cardClick(1);
</script>

</BODY>
</HTML>