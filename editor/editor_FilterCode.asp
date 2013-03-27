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
<TITLE>字符过滤</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
body, a, table, div, span, td, th, input, select{font-size:9pt;font-family: "宋体", Verdana, Arial, Helvetica, sans-serif;}
body {padding:5px}
</style>

<script language="JavaScript">
function Filtertext(){
    var str1
    str1 = document.myform.Script_Iframe.checked
    str1 += "," + document.myform.Script_Object.checked
    str1 += "," + document.myform.Script_Script.checked
    str1 += "," + document.myform.Script_Class.checked
    str1 += "," + document.myform.Script_Div.checked
    str1 += "," + document.myform.Script_Span.checked
    str1 += "," + document.myform.Script_Table.checked
    str1 += "," + document.myform.Script_Table2.checked
    str1 += "," + document.myform.Script_Img.checked
    str1 += "," + document.myform.Script_Font.checked
    str1 += "," + document.myform.Script_A.checked
    str1 += "," + document.myform.Script_Font2.checked
    str1 += "," + document.myform.FontFilterText.value
    window.returnValue = str1
    window.close();
}
</script>
</HEAD>
<BODY bgColor="#D4D0C8">
<FORM NAME="myform" method="post" action="">
<TABLE CELLSPACING="0" cellpadding="0" border="0">
<TR>
<TD width="500"><fieldset><legend><b>字符过滤设置</b></legend>
  <table CELLSPACING="0" cellpadding="5" border="0">
    <tr class='tdbg'>
       <td height="22">
          <input name="Script_Iframe" type="checkbox" id="Script_Iframe"  value="yes" >Iframe：  &nbsp;过滤内联页。<br>
          <input name="Script_Object" type="checkbox" id="Script_Object"  value="yes" >Object： &nbsp;过滤Falsh广告,控件等。<br>
          <input name="Script_Script" type="checkbox" id="Script_Script"  value="yes" >Script： &nbsp;过滤js、vbs等脚本。<br>
          <input name="Script_Class" type="checkbox" id="Script_Class"  value="yes" >Style： &nbsp;过滤Css 类。<br>
          <input name="Script_Div" type="checkbox" id="Script_Div"  value="yes" >Div： &nbsp;过滤层。<br>
          <input name="Script_Span" type="checkbox" id="Script_Span"  value="yes" >Span： 过滤行内元素Span容器。<br>
          <input name="Script_Table" type="checkbox" id="Script_Table"  value="yes" >Table ：过滤表格及表格里面的所有内容。<br>
          <input name="Script_Tr" type="checkbox" id="Script_Table2"  value="yes" >Table ：仅过滤表格本身内容，表格里面的内容不过滤。<br>
          <input name="Script_Img" type="checkbox" id="Script_Img"  value="yes" >Img：&nbsp;过滤图片。<Font color=blue >注意：不建议过滤</Font><br>
          <input name="Script_Font" type="checkbox" id="Script_Font"  value="yes" >FONT：&nbsp;过滤字体定义。 (字留下样式去掉) <br>
          <input name="Script_Font2" type="checkbox" id="Script_Font2"  value="yes" >过滤带有指定字符的字体：<Input TYPE='Text' Name='FontFilterText' value='' id='id' size='10' maxlength='20'> <br>
          &nbsp;&nbsp;&nbsp;<font color='blue'>注意请在编辑器代码模式下选取字符</font><br>
          &nbsp;&nbsp;&nbsp;<font color='blue'>因为复制后编辑器会转换复制字符的大小写</font><br>
          <input name="Script_A" type="checkbox" id="Script_A"  value="yes" >A：&nbsp;过滤链接 (字留下链接去掉)<br>
        </td>
     </tr>
</table>
</fieldset>
</td>
<td> </td>
<td rowspan="2" valign="top">
  <Input type=button style="width:80px;margin-top:15px" name="btnFind" onClick="Filtertext();" value="过滤"><br>
  <Input type=button style="width:80px;margin-top:5px" name="btnCancel" onClick="window.close();" value="取消"><br>
</td>
</tr>
</table>
</FORM>
</BODY>
</HTML>
