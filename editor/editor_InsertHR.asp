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
<title>����ˮƽ��</title>
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
<LEGEND align=left><strong>����ˮƽ�߲���</strong></LEGEND>
      <table border="0" cellpadding="0" cellspacing="3">
        <tr> 
          <td>������ɫ��
            <input name="t_color" id=t_color  size="7" maxlength="7">
	    <img border=0 src="images/rect.gif" width=18 style="cursor:hand" id=s_color onclick="SelectColor('color')">
          </td>

        </tr>
        <tr>
          <td>�����ֶȣ�
            <input name="size"  id=size onKeyPress="event.returnValue=IsDigit();" value="2" size="4" maxlength=3>
���������֣���Χ������1-100֮��</td>
        </tr>
        <tr> 
          <td> ҳ����룺
            <select name="align"  id=align>
              <option value="left" selected>Ĭ�϶���</option>
              <option value="left">����� </option>
              <option value="center">�ж��� </option>
              <option value="right">�Ҷ��� </option>
            </select>
            &nbsp;&nbsp;��ӰЧ����
            <select name="shadetype"  id=shadetype>
              <option value=noshade selected>�� 
              <option value=''>�� 
            </select>
          </td>
        </tr>
        <tr> 
          <td> ˮƽ��ȣ�
            <input name="width" id=width ONKEYPRESS="event.returnValue=IsDigit();" value="400" size="6" maxlength=3>
            ���������֣���Χ������1-999֮��</td>
        </tr>
      </table>
</fieldset></td>
    <td width=80 align="center"><input name="cmdOK" type="button" id="cmdOK" value="  ȷ��  " onClick="OK();">
      <br>
      <br>
      <input name="cmdCancel" type=button id="cmdCancel" onclick="window.close();" value="  ȡ��  "></td>
  </tr></table>
</body>
</html>