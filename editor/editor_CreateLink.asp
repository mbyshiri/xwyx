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

dim LinkName
LinkName=trim(request("LinkName"))
%>
<HTML>
<HEAD>
<title>���볬������</title>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link rel="stylesheet" type="text/css" href="editor_dialog.css">
<a href='' target='_blank'></a>
</head>
<BODY bgColor='#D4D0C8' topmargin=15 leftmargin=15 >
<FORM METHOD='POST' name='myform' ACTION=''>
<table width=100% border="0" cellpadding="0" cellspacing="2">
  <tr>
    <td>
     <FIELDSET align=left>
      <LEGEND align=left><strong>����������Ϣ</strong></LEGEND>
       <table border="0" cellpadding="0" cellspacing="3">
        <tr> 
          <td align='right'>�������ͣ�</td>
          <td>
           <select name='LinkType' id='id' onchange='document.myform.linkurl.value=this.value;'>
              <option value='http://' selected>http:</option>
              <option value='ftp://'>ftp:</option>
              <option value='file://'>�ļ�</option>
              <option value='news://'>news:</option>
              <option value='mailto:'>mailto:</option>
           </select>
          </td>
        </tr>
        <tr>
          <td align='right'>���ӵ�ַ��</td>    
          <td>
            <input name="linkurl" value="http://" size="40" >
          </td>
        </tr>
        <tr> 
         <td align='right'>�򿪷�ʽ��</td>    
         <td>
           <select name='Openfashion' id='Openfashion' >
              <option value='0' selected>ԭ�����</option>
              <option value='1'>�´����</option>
           </select>
         </td>
        </tr>
        <tr>
          <td align='right'>title��</td>
          <td>
            <input name="Linktitle"  value="" size="30" ><br><FONT style='font-size:12px' color='blue'>������д����title����</FONT>
          </td>
        </tr>
        <tr> 
          <td align='right'>����������ɫ��</td>
          <td>
            <input name="t_color" id=t_color  size="7" maxlength="7">
            <img border=0 src="images/rect.gif" width=18 style="cursor:hand" id=s_color onclick="SelectColor('color')">
          </td>
        </tr>
        <tr class='tdbg'  class='tdbg5'>
          <td align="right" class='tdbg5'>�����Ƿ�Ӵ֣�</td>
          <td>
            <input type="radio" name="LinkB" value="true">��
            <input type="radio" name="LinkB" value="false" checked>��
          </td>
        </tr>
        <tr class='tdbg'  class='tdbg5'>
          <td align="right"  class='tdbg5'>�����Ƿ���»��ߣ�</td>
          <td>
            <input type="radio" name="LinkX" value="true">��
            <input type="radio" name="LinkX" value="false" checked>��
          </td>
        </tr>
        <tr>
          <td align='right'>������չ��</td>    
          <td>
            <input name="Linkexpand"  value="" size="30" ><br><FONT style='font-size:12px' color='blue'>������д������չ���Ի�ű��¼�</FONT>
          </td>
        </tr>
        <tr class='tdbg'  class='tdbg5'>
          <td align="right"  class='tdbg5'>Ҫ���ӵ����ݣ�</td>
          <td>
            <TEXTAREA style="WIDTH: 240px; HEIGHT: 100px" name="EditTagCode"></TEXTAREA>
          </td>
        </tr>
       </table>
      </FORM>
     </fieldset>
    </td>
    <td width=80 align="center">
      <input name="cmdOK" type="button" id="cmdOK" value="  ȷ��  " onClick="OK();">
      <br><br>
      <input name="cmdCancel" type=button id="cmdCancel" onclick="window.close();" value="  ȡ��  ">
    </td>
   </tr>
  </table>
<script language="JavaScript">
var oControl;
var oSeletion;
var sRangeType;
var LinkType;

oSelection = dialogArguments.HtmlEdit.document.selection.createRange();
sRangeType = dialogArguments.HtmlEdit.document.selection.type;
if (sRangeType == "Control") {
    oControl = oSelection.item(0);
    document.myform.EditTagCode.value=oControl.outerHTML;
}else {
    if (dialogArguments.HtmlEdit!=null) oControl=dialogArguments.HtmlEdit;
    document.myform.EditTagCode.value="<%=LinkName%>";
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
function OK(){
    var str1;
    var str2;
    var LinkName;
    var LinkB;
    var LinkX;
    for (var i=0;i<document.myform.LinkB.length;i++){
        var PowerEasy = document.myform.LinkB[i];
        if (PowerEasy.checked==true)       
            LinkB = PowerEasy.value
        }
    for (var i=0;i<document.myform.LinkX.length;i++){
        var PowerEasy = document.myform.LinkX[i];
        if (PowerEasy.checked==true)       
            LinkX = PowerEasy.value
    }
    LinkName=document.myform.EditTagCode.value;
    if (document.myform.t_color.value!="" ){
        LinkName="<font color="+document.myform.t_color.value+">"+ LinkName +"</font>";
    }
    if (LinkB=="true" ){
        LinkName="<B>"+LinkName+"</B>";
    }
    if (LinkX=="true" ){
        LinkName="<U>"+LinkName+"</U>";
    }
    if (document.myform.Openfashion.value == '1') {
        str2="target=\"_blank\" "
    }else{
        str2=""    
    }
    str1="<a href='"+document.myform.linkurl.value+"' title=\""+document.myform.Linktitle.value+"\" "+str2+document.myform.Linkexpand.value+">"+LinkName+"</a>"
    window.returnValue = str1
    window.close();
}
</script>
</body>
</html>