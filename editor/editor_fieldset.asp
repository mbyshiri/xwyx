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
<title>������Ŀ��</title>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link rel="stylesheet" type="text/css" href="editor_dialog.css">
<script language="JavaScript">
function OK(){
    var str1="";
    str1="<FIELDSET align='"+align1.value+"' style='"
    if(t_color.value!='')str1=str1+"color:"+t_color.value+";"
    if(t_backcolor.value!='')str1=str1+"background-color:"+t_backcolor.value+";"
    str1=str1+"'><Legend"
    str1=str1+" align="+align2.value+">"+LegendTitle.value+"</Legend>��������������Ŀ�������</FIELDSET>"
    window.returnValue = str1;
    window.close();
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

<body bgColor=#D4D0C8 topmargin=15 leftmargin=15 >
<table width=100% border="0" cellpadding="0" cellspacing="2">
  <tr><td>
<FIELDSET align=left>
      <LEGEND align=left><strong>������Ŀ��</strong></LEGEND>
      <table border="0" align="center" cellpadding="0" cellspacing="3">
        <tr>
          <td align="right">��Ŀ����뷽ʽ��</td>
          <td><select name="align1" id=align1>
              <option value="left" selected>�����
              <option value="center">����
              <option value="right">�Ҷ���
          </select></td>
        </tr>
        <tr> 
          <td align="right">��Ŀ���⣺            </td>
          <td><input name="LegendTitle" type="text" id="LegendTitle" size="20"></td>
        </tr>
        <tr>
          <td align="right">������뷽ʽ��          </td>
          <td><select name="align2" id=select3>
            <option value="left" selected>�����
            <option value="center">����
            <option value="right">�Ҷ���
          </select></td>
        </tr>
	<tr> 
           <td align="right">�߿���ɫ��          </td>
          <td><input name="t_color" id=t_color  size="7" maxlength="7">
	<img border=0 src="images/rect.gif" width=18 style="cursor:hand" id=s_color onclick="SelectColor('color')">
	  
	  </td>
        </tr>
        <tr>
          <td align="right">������ɫ��          </td>
          <td><input name="t_backcolor" id=t_backcolor size="7" maxlength="7">
	  <img border=0 src="images/rect.gif" width=18 style="cursor:hand" id=s_backcolor onclick="SelectColor('backcolor')">
	  </td>
        </tr>
      </table>
</FIELDSET>
</td><td width=80 align="center"><input name="cmdOK" type="button" id="cmdOK" value="  ȷ��  " onClick="OK();">
<br>
<br><input name="cmdCancel" type=button id="cmdCancel" onclick="window.close();" value='  ȡ��  '></td></tr></table>
</body>
</html>