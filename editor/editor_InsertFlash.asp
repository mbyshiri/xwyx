<!--#include file="editor_ChkPurview.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

%>
<HTML>
<HEAD>
<TITLE>����FLASH�ļ�</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="editor_dialog.css">
<script language="JavaScript">
//=================================================
//��������OK()
//��  �ã��ύ��Ϣ
//=================================================
function OK(){
    var str1="";
    var strurl=document.form1.url.value;
    if (strurl==""||strurl=="http://"){
        alert("��������FLASH�ļ���ַ�������ϴ�FLASH�ļ���");
        document.form1.url.focus();
        return false;
    }else{
        str1 = "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width=" + document.form1.width.value + " height=" + document.form1.height.value + "><param name=movie value=" + document.form1.url.value + "><param name=quality value=high><embed src=" + document.form1.url.value + " pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=" + document.form1.width.value + " height=" + document.form1.height.value + "></embed></object>"
        window.returnValue = str1+"$$$"+document.form1.UpFileName.value;
        window.close();
    }
}
//=================================================
//��������IsDigit()
//��  �ã�����Ϊ����
//=================================================
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
//=================================================
//��������imgwidth
//��  �ã�������ʾFlash���
//��  ����element   --- ���ر�ֵ
//=================================================
function swfModify(){
    if(document.form1.url.value=="http://"){
        document.form1.url.value = "logo3.swf"
    }
    objFiles.innerHTML = "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width=" + document.form1.width.value + " height=" + document.form1.height.value + "><param name=movie value=" + document.form1.url.value + "><param name=quality value=high><embed src=" + document.form1.url.value + " pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=" + document.form1.width.value + " height=" + document.form1.height.value + "></embed></object>"
}
function SelectFile(){
    var arr=showModalDialog('<%=InstallDir & AdminDir%>/Admin_SelectFile.asp?DialogType=Flash&ChannelID=<%=ChannelID%>', '', 'dialogWidth:820px; dialogHeight:600px; help: no; scroll: yes; status: no');
    if(arr!=null){
        var ss=arr.split('|');
        document.form1.url.value=ss[0];
        document.form1.UpFileName.value=ss[0].replace("<%=FilesPath%>", "");
        swfModify();
    }
}
</script>
</head>
<body bgColor=#D4D0C8 topmargin=15 leftmargin=15 >
<form name="form1" method="post" action="">
  <table width=100% border="0" cellpadding="0" cellspacing="2">
    <tr>
      <td>
      <FIELDSET align=left>
      <LEGEND align=left>FLASH��������</LEGEND>
        <table border="0" cellpadding="0" cellspacing="3" >
          <tr>
            <td  height=5></td>
          </tr>
          <tr>
            <td width=350 align='center' id='objFiles'>
            <IMG SRC='../images/filetype_flash.gif'  id=img align='center' width='300' height='200'  BORDER='0' ALT=''>
            </td>
          </tr>
          <tr>
            <td align='center' height='5'></td>
          </tr>
          <tr>
            <td height="17" >��ַ��
             <Input name="url" id=url value="http://"  onChange="javascript:swfModify()" size=45>
            <%if IsUpload=True And AdminName <> "" then %>
             <Input type="button" name="Submit" value="..." title="�����ϴ��ļ���ѡ��" onClick="SelectFile()">
            <%End if%>
            </td>
          </tr>
          <tr>
            <td>��ȣ�
             <Input name="width" id=width ONKEYPRESS="event.returnValue=IsDigit();" onChange="javascript:swfModify()" value=300 size=7 maxlength="4">   �߶ȣ�
             <Input name="height" id=height ONKEYPRESS="event.returnValue=IsDigit();" onChange="javascript:swfModify()" value=200 size=7 maxlength="4">
            </TD>
          </tr>
        </table>
        </fieldset>
      </td>
      <td width=80 align="center">
       <Input name="cmdOK" type="button" id="cmdOK" value="  ȷ��  " onClick="OK();">
      <br><br>
       <Input name="cmdCancel" type=button id="cmdCancel" onclick="window.close();" value='  ȡ��  '>
      </td>
    </tr>
    <%if IsUpload=True then %>
    <tr>
      <td>
      <FIELDSET align=left>
      <LEGEND align=left>�ϴ�����FLASH�ļ�</LEGEND>
      <%
        Response.write "<iframe class=""TBGen"" style=""top:2px"" id=""UploadFiles"" src=""upload.asp?DialogType=flash"
        Response.write "&ChannelID=" & ChannelID
        If PE_CLng(Request(Trim("Anonymous"))) = 1 Then
            Response.write "&Anonymous=1"
        End If		
        If ModuleType=3 Then
            Response.write "&PhotoUpfileType=1"
        End If
        Response.write """ frameborder=0 scrolling=no width=""350"" height=""25""></iframe>"
        Response.write "</fieldset></td>"
        Response.write "</tr>"
    End if 
    %>
     <Input name="UpFileName" type="hidden" id="UpFileName" value="None">
  </table>
  </form>
</body>
</html>
