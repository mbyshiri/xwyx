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
<TITLE>�ַ�����</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
body, a, table, div, span, td, th, input, select{font-size:9pt;font-family: "����", Verdana, Arial, Helvetica, sans-serif;}
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
<TD width="500"><fieldset><legend><b>�ַ���������</b></legend>
  <table CELLSPACING="0" cellpadding="5" border="0">
    <tr class='tdbg'>
       <td height="22">
          <input name="Script_Iframe" type="checkbox" id="Script_Iframe"  value="yes" >Iframe��  &nbsp;��������ҳ��<br>
          <input name="Script_Object" type="checkbox" id="Script_Object"  value="yes" >Object�� &nbsp;����Falsh���,�ؼ��ȡ�<br>
          <input name="Script_Script" type="checkbox" id="Script_Script"  value="yes" >Script�� &nbsp;����js��vbs�Ƚű���<br>
          <input name="Script_Class" type="checkbox" id="Script_Class"  value="yes" >Style�� &nbsp;����Css �ࡣ<br>
          <input name="Script_Div" type="checkbox" id="Script_Div"  value="yes" >Div�� &nbsp;���˲㡣<br>
          <input name="Script_Span" type="checkbox" id="Script_Span"  value="yes" >Span�� ��������Ԫ��Span������<br>
          <input name="Script_Table" type="checkbox" id="Script_Table"  value="yes" >Table �����˱�񼰱��������������ݡ�<br>
          <input name="Script_Tr" type="checkbox" id="Script_Table2"  value="yes" >Table �������˱�������ݣ������������ݲ����ˡ�<br>
          <input name="Script_Img" type="checkbox" id="Script_Img"  value="yes" >Img��&nbsp;����ͼƬ��<Font color=blue >ע�⣺���������</Font><br>
          <input name="Script_Font" type="checkbox" id="Script_Font"  value="yes" >FONT��&nbsp;�������嶨�塣 (��������ʽȥ��) <br>
          <input name="Script_Font2" type="checkbox" id="Script_Font2"  value="yes" >���˴���ָ���ַ������壺<Input TYPE='Text' Name='FontFilterText' value='' id='id' size='10' maxlength='20'> <br>
          &nbsp;&nbsp;&nbsp;<font color='blue'>ע�����ڱ༭������ģʽ��ѡȡ�ַ�</font><br>
          &nbsp;&nbsp;&nbsp;<font color='blue'>��Ϊ���ƺ�༭����ת�������ַ��Ĵ�Сд</font><br>
          <input name="Script_A" type="checkbox" id="Script_A"  value="yes" >A��&nbsp;�������� (����������ȥ��)<br>
        </td>
     </tr>
</table>
</fieldset>
</td>
<td> </td>
<td rowspan="2" valign="top">
  <Input type=button style="width:80px;margin-top:15px" name="btnFind" onClick="Filtertext();" value="����"><br>
  <Input type=button style="width:80px;margin-top:5px" name="btnCancel" onClick="window.close();" value="ȡ��"><br>
</td>
</tr>
</table>
</FORM>
</BODY>
</HTML>
