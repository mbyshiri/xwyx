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
<TITLE>�����ϴ�����</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="editor_dialog.css">
<script language="JavaScript">

function OK(){
    var str="";
    var strurl=document.form1.url.value;
    if (strurl==""||strurl=="http://"){
        alert("���������ϴ��ļ��ĵ�ַ��");
        document.form1.url.focus();
        return false;
    }else if (document.form1.title.value==""){
        alert("�������Ʋ���Ϊ�գ�");
        document.form1.title.focus();
        return false;
    }else{
        str="<a href='"+document.form1.url.value+"' title='"+document.form1.title.value+"'>"+document.form1.title.value+"</a>"
        window.returnValue=str+"$$$"+document.form1.UpFileName.value;
        window.close();
    }
}
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
function SelectFile(){
  var arr=showModalDialog('<%=InstallDir & AdminDir%>/Admin_SelectFile.asp?DialogType=FuJian&ChannelID=<%=ChannelID%>', '', 'dialogWidth:820px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.form1.url.value=ss[0];
    document.form1.UpFileName.value=ss[0].replace("<%=FilesPath%>", "");
  }
}
</script>
</head>
<BODY bgColor=#D4D0C8 topmargin=15 leftmargin=15 >
<form name="form1" method="post" action="">
 <table width=100% border="0" cellpadding="0" cellspacing="2">
   <tr>
     <td>
       <FIELDSET align=left>
       <LEGEND align=left>�ϴ���������</LEGEND>
       <TABLE border="0" cellpadding="0" cellspacing="3" >
        <TR>
     <TD height="17" >��ַ��<INPUT name="url" id=url value="http://" size=40>
        <%if IsUpload=True And AdminName <> "" then %>
             <input type="button" name="Submit" value="..." title="�����ϴ��ļ���ѡ��" onClick="SelectFile()">
        <%End if%>
     </td>
        </TR>
        <TR>
      <TD >�����븽�����ƣ�<INPUT TYPE="text" NAME="title" size="20"></TD></TR>
        <TR>
      <TD align='center'><FONT style='font-size:12px' color='#339900'>'ע��ֻ���ϴ���׺Ϊ '.zip,.doc ,.rar ���ļ�</FONT>
      </TD>
    </TR>
       </TABLE>
       </fieldset>
     </td>
     <td width=80 align="center"><input name="cmdOK" type="button" id="cmdOK" value="  ȷ��  " onClick="OK();">
     <br><br><input name="cmdCancel" type=button id="cmdCancel" onClick="window.close();" value='  ȡ��  '>
     </td>
   </tr>
   <%if IsUpload=True then %>
    <tr>
      <td><fieldset align=left>
      <legend align=left>�ϴ����ظ���</legend>
    <%
        Response.write "<iframe class=""TBGen"" style=""top:2px"" id=""UploadFiles"" src=""upload.asp?DialogType=fujian"
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
    <tr>
      <td height=5></td>
    </tr>
    <input name="UpFileName" type="hidden" id="UpFileName" value="None">
  </table>
</form>
</body>
</html>

