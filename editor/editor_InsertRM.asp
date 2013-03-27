<!--#include file="editor_ChkPurview.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

%>
<HTML>
<HEAD>
<TITLE>插入RealPlay文件</TITLE>
<link rel="stylesheet" type="text/css" href="editor_dialog.css">
<script language="JavaScript">
function OK(){
    var str1="";
    var strurl=document.form1.url.value;
    if (strurl==""||strurl=="http://"||strurl=="rtsp://"){
        alert("请先输入RealPlay文件地址，或者上传RealPlay文件！");
        document.form1.url.focus();
        return false;
    }else{
        str1 = "<object classid='clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA' width="+document.form1.width.value+" height="+document.form1.height.value+"><param name='CONTROLS' value='ImageWindow'><param name='CONSOLE' value='Clip1'><param name='AUTOSTART' value='-1'><param name=src value="+document.form1.url.value+"></object><br><object classid='clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA'  width="+document.form1.width.value+" height=60><param name='CONTROLS' value='ControlPanel,StatusBar'><param name='CONSOLE' value='Clip1'></object>"
        window.returnValue = str1+"$$$"+document.form1.UpFileName.value;
        window.close();
    }
}
//=================================================
//过程名：IsDigit()
//作  用：输入为数字
//=================================================
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
//=================================================
//过程名：ShowRm
//作  用：在线显示RM
//参  数：element   --- 返回表单值
//=================================================
function ShowRm(){
       if(document.form1.url.value=="http://"){
           document.Form1.url.Value = "地址"
       }
      objFiles.innerHTML = "<object classid='clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA' width="+document.form1.width.value+" height="+document.form1.height.value+"><param name='CONTROLS' value='ImageWindow'><param name='CONSOLE' value='Clip1'><param name='AUTOSTART' value='-1'><param name=src value="+document.form1.url.value+"></object><br><object classid='clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA'  width="+document.form1.width.value+" height=60><param name='CONTROLS' value='ControlPanel,StatusBar'><param name='CONSOLE' value='Clip1'></object>"
}
function SelectFile(){
    var arr=showModalDialog('<%=InstallDir & AdminDir%>/Admin_SelectFile.asp?DialogType=rm&ChannelID=<%=ChannelID%>', '', 'dialogWidth:820px; dialogHeight:600px; help: no; scroll: yes; status: no');
    if(arr!=null){
        var ss=arr.split('|');
        document.form1.url.value=ss[0];
        var arrContent=ss[0].split('/');
        document.form1.UpFileName.value=ss[0].replace("<%=FilesPath%>", "");
        ShowRm();
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
    <LEGEND align=left>RealPlay文件参数</LEGEND>
     <TABLE border="0" cellpadding="0" cellspacing="3">
        <tr><td  height=5></td></tr>
        <tr>
          <td width=350 align='center' id='objFiles'>
        <!-- **********    RM开始　********** -->
              <object id="player" name="player" classid="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA" width="300" height="220">
            <param name="CONTROLS" value="Imagewindow">
            <param name="CONSOLE" value="clip1">
            <param name="AUTOSTART" value="0">
            <param name="SRC" value="">
            </object><br>
            <object ID="RP2" CLASSID="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA" WIDTH="300" HEIGHT="60">
            <PARAM NAME="CONTROLS" VALUE="ControlPanel,StatusBar">
            <param name="CONSOLE" value="clip1">
            </object>
        <!-- **********    RM结束　********** -->
          </td>
        </tr>
    <tr><td align='center' height='5'></td></tr>
      <TR>
        <TD >地址：<INPUT name="url" id=url  value="rtsp://" size=40 onChange="javascript:ShowRm()">
        <%if IsUpload=True And AdminName <> "" then %>
             <input type="button" name="Submit" value="..." title="从已上传文件中选择" onClick="SelectFile()">
        <%End if%>
        </td>
      </TR>
      <TR>
       <TD>宽度：<INPUT name="width" id=width ONKEYPRESS="event.returnValue=IsDigit();" value=300 size=7 maxlength="4" onChange="javascript:ShowRm()"> &nbsp;&nbsp;高度：<INPUT name="height" id=height ONKEYPRESS="event.returnValue=IsDigit();" value=200 size=7 maxlength="4" onChange="javascript:ShowRm()">
       </TD>
      </TR>
      <TR>
        <TD align=center>支持格式为：rm、ra、ram</TD>
      </TR>
     </TABLE>
     </fieldset>
    </td>
    <td width=80 align="center"><input name="cmdOK" type="button" id="cmdOK" value="  确定  " onClick="OK();">
    <br>
    <br>  <input name="cmdCancel" type=button id="cmdCancel" onClick="window.close();" value='  取消  '>
    </td>
 </tr>
 <%if IsUpload=True then %>
 <tr>
   <td>
   <FIELDSET align=left>
    <LEGEND align=left>上传本地视频文件</LEGEND>
    <%
        Response.write "<iframe class=""TBGen"" style=""top:2px"" id=""UploadFiles"" src=""upload.asp?DialogType=real"
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
 <input name="UpFileName" type="hidden" id="UpFileName" value="None">
</table>
</form>
</body>
</html>