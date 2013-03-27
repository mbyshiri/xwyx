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
<TITLE>插入视频文件</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="editor_dialog.css">
<script language="JavaScript">
function OK(){
    var str1="";
    var strurl=document.form1.url.value;
    if (strurl==""||strurl=="http://"){
        alert("请先输入视频文件地址，或者上传视频文件！");
        document.form1.url.focus();
        return false;
    }else{
        str1 = "<OBJECT id=MediaPlayer1"
        str1=str1+="    codeBase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701standby=Loading"
        str1=str1+="    type=application/x-oleobject height="+document.form1.height.value+" width="+document.form1.width.value
        str1=str1+="    classid=CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6 VIEWASTEXT>"
        str1=str1+="    <PARAM NAME=\"URL\" value="+document.form1.url.value+">"
        str1=str1+="    <param name=\"AudioStream\" value=\"-1\">"
        str1=str1+="    <param name=\"AutoSize\" value=\"0\">"
        str1=str1+="    <param name=\"AutoStart\" value=\"-1\">"
        str1=str1+="    <param name=\"AnimationAtStart\" value=\"0\">"
        str1=str1+="    <param name=\"AllowScan\" value=\"-1\">"
        str1=str1+="    <param name=\"AllowChangeDisplaySize\" value=\"-1\">"
        str1=str1+="    <param name=\"AutoRewind\" value=\"0\">"
        str1=str1+="    <param name=\"Balance\" value=\"0\">"
        str1=str1+="    <param name=\"BaseURL\" value>"
        str1=str1+="    <param name=\"BufferingTime\" value=\"5\">"
        str1=str1+="    <param name=\"CaptioningID\" value>"
        str1=str1+="    <param name=\"ClickToPlay\" value=\"-1\">"
        str1=str1+="    <param name=\"CursorType\" value=\"0\">"
        str1=str1+="    <param name=\"CurrentPosition\" value=\"-1\">"
        str1=str1+="    <param name=\"CurrentMarker\" value=\"0\">"
        str1=str1+="    <param name=\"DefaultFrame\" value>"
        str1=str1+="    <param name=\"DisplayBackColor\" value=\"0\">"
        str1=str1+="    <param name=\"DisplayForeColor\" value=\"16777215\">"
        str1=str1+="    <param name=\"DisplayMode\" value=\"0\">"
        str1=str1+="    <param name=\"DisplaySize\" value=\"4\">"
        str1=str1+="    <param name=\"Enabled\" value=\"-1\">"
        str1=str1+="    <param name=\"EnableContextMenu\" value=\"-1\">"
        str1=str1+="    <param name=\"EnablePositionControls\" value=\"0\">"
        str1=str1+="    <param name=\"EnableFullScreenControls\" value=\"0\">"
        str1=str1+="    <param name=\"EnableTracker\" value=\"-1\">"
        str1=str1+="    <param name=\"InvokeURLs\" value=\"-1\">"
        str1=str1+="    <param name=\"Language\" value=\"-1\">"
        str1=str1+="    <param name=\"Mute\" value=\"0\">"
        str1=str1+="    <param name=\"PlayCount\" value=\"1\">"
        str1=str1+="    <param name=\"PreviewMode\" value=\"0\">"
        str1=str1+="    <param name=\"Rate\" value=\"1\">"
        str1=str1+="    <param name=\"SAMILang\" value>"
        str1=str1+="    <param name=\"SAMIStyle\" value>"
        str1=str1+="    <param name=\"SAMIFileName\" value>"
        str1=str1+="    <param name=\"SelectionStart\" value=\"-1\">"
        str1=str1+="    <param name=\"SelectionEnd\" value=\"-1\">"
        str1=str1+="    <param name=\"SendOpenStateChangeEvents\" value=\"-1\">"
        str1=str1+="    <param name=\"SendWarningEvents\" value=\"-1\">"
        str1=str1+="    <param name=\"SendErrorEvents\" value=\"-1\">"
        str1=str1+="    <param name=\"SendKeyboardEvents\" value=\"0\">"
        str1=str1+="    <param name=\"SendMouseClickEvents\" value=\"0\">"
        str1=str1+="    <param name=\"SendMouseMoveEvents\" value=\"0\">"
        str1=str1+="    <param name=\"SendPlayStateChangeEvents\" value=\"-1\">"
        str1=str1+="    <param name=\"ShowCaptioning\" value=\"0\">"
        str1=str1+="    <param name=\"ShowControls\" value=\"-1\">"
        str1=str1+="    <param name=\"ShowAudioControls\" value=\"-1\">"
        str1=str1+="    <param name=\"ShowDisplay\" value=\"0\">"
        str1=str1+="    <param name=\"ShowGotoBar\" value=\"0\">"
        str1=str1+="    <param name=\"ShowPositionControls\" value=\"-1\">"
        str1=str1+="    <param name=\"ShowStatusBar\" value=\"-1\">"
        str1=str1+="    <param name=\"ShowTracker\" value=\"-1\">"
        str1=str1+="    <param name=\"TransparentAtStart\" value=\"-1\">"
        str1=str1+="    <param name=\"VideoBorderWidth\" value=\"0\">"
        str1=str1+="    <param name=\"VideoBorderColor\" value=\"0\">"
        str1=str1+="    <param name=\"VideoBorder3D\" value=\"0\">"
        str1=str1+="    <param name=\"Volume\" value=\"70\">"
        str1=str1+="    <param name=\"WindowlessVideo\" value=\"0\">"
        str1=str1+="</OBJECT>"

        window.returnValue = str1+"$$$"+document.form1.UpFileName.value;
        window.close();
    }
}
//=================================================
//过程名：IsDigit()
//作  用：输入为数字
//=================================================
function IsDigit(){
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
//=================================================
//过程名：windowplay
//作  用：在线播放windowplay
//=================================================
function windowplay(){
    if(document.form1.url.value=="http://"){
        document.form1.url.Value = "1.mp3"
    }
    str1 = "<OBJECT id=MediaPlayer1"
    str1=str1+="    codeBase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701standby=Loading"
    str1=str1+="    type=application/x-oleobject height="+document.form1.height.value+" width="+document.form1.width.value
    str1=str1+="    classid=CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6 VIEWASTEXT>"
    str1=str1+="    <PARAM NAME=\"URL\" value="+document.form1.url.value+">"
    str1=str1+="    <param name=\"AudioStream\" value=\"-1\">"
    str1=str1+="    <param name=\"AutoSize\" value=\"0\">"
    str1=str1+="    <param name=\"AutoStart\" value=\"-1\">"
    str1=str1+="    <param name=\"AnimationAtStart\" value=\"0\">"
    str1=str1+="    <param name=\"AllowScan\" value=\"-1\">"
    str1=str1+="    <param name=\"AllowChangeDisplaySize\" value=\"-1\">"
    str1=str1+="    <param name=\"AutoRewind\" value=\"0\">"
    str1=str1+="    <param name=\"Balance\" value=\"0\">"
    str1=str1+="    <param name=\"BaseURL\" value>"
    str1=str1+="    <param name=\"BufferingTime\" value=\"5\">"
    str1=str1+="    <param name=\"CaptioningID\" value>"
    str1=str1+="    <param name=\"ClickToPlay\" value=\"-1\">"
    str1=str1+="    <param name=\"CursorType\" value=\"0\">"
    str1=str1+="    <param name=\"CurrentPosition\" value=\"-1\">"
    str1=str1+="    <param name=\"CurrentMarker\" value=\"0\">"
    str1=str1+="    <param name=\"DefaultFrame\" value>"
    str1=str1+="    <param name=\"DisplayBackColor\" value=\"0\">"
    str1=str1+="    <param name=\"DisplayForeColor\" value=\"16777215\">"
    str1=str1+="    <param name=\"DisplayMode\" value=\"0\">"
    str1=str1+="    <param name=\"DisplaySize\" value=\"4\">"
    str1=str1+="    <param name=\"Enabled\" value=\"-1\">"
    str1=str1+="    <param name=\"EnableContextMenu\" value=\"-1\">"
    str1=str1+="    <param name=\"EnablePositionControls\" value=\"0\">"
    str1=str1+="    <param name=\"EnableFullScreenControls\" value=\"0\">"
    str1=str1+="    <param name=\"EnableTracker\" value=\"-1\">"
    str1=str1+="    <param name=\"InvokeURLs\" value=\"-1\">"
    str1=str1+="    <param name=\"Language\" value=\"-1\">"
    str1=str1+="    <param name=\"Mute\" value=\"0\">"
    str1=str1+="    <param name=\"PlayCount\" value=\"1\">"
    str1=str1+="    <param name=\"PreviewMode\" value=\"0\">"
    str1=str1+="    <param name=\"Rate\" value=\"1\">"
    str1=str1+="    <param name=\"SAMILang\" value>"
    str1=str1+="    <param name=\"SAMIStyle\" value>"
    str1=str1+="    <param name=\"SAMIFileName\" value>"
    str1=str1+="    <param name=\"SelectionStart\" value=\"-1\">"
    str1=str1+="    <param name=\"SelectionEnd\" value=\"-1\">"
    str1=str1+="    <param name=\"SendOpenStateChangeEvents\" value=\"-1\">"
    str1=str1+="    <param name=\"SendWarningEvents\" value=\"-1\">"
    str1=str1+="    <param name=\"SendErrorEvents\" value=\"-1\">"
    str1=str1+="    <param name=\"SendKeyboardEvents\" value=\"0\">"
    str1=str1+="    <param name=\"SendMouseClickEvents\" value=\"0\">"
    str1=str1+="    <param name=\"SendMouseMoveEvents\" value=\"0\">"
    str1=str1+="    <param name=\"SendPlayStateChangeEvents\" value=\"-1\">"
    str1=str1+="    <param name=\"ShowCaptioning\" value=\"0\">"
    str1=str1+="    <param name=\"ShowControls\" value=\"-1\">"
    str1=str1+="    <param name=\"ShowAudioControls\" value=\"-1\">"
    str1=str1+="    <param name=\"ShowDisplay\" value=\"0\">"
    str1=str1+="    <param name=\"ShowGotoBar\" value=\"0\">"
    str1=str1+="    <param name=\"ShowPositionControls\" value=\"-1\">"
    str1=str1+="    <param name=\"ShowStatusBar\" value=\"-1\">"
    str1=str1+="    <param name=\"ShowTracker\" value=\"-1\">"
    str1=str1+="    <param name=\"TransparentAtStart\" value=\"-1\">"
    str1=str1+="    <param name=\"VideoBorderWidth\" value=\"0\">"
    str1=str1+="    <param name=\"VideoBorderColor\" value=\"0\">"
    str1=str1+="    <param name=\"VideoBorder3D\" value=\"0\">"
    str1=str1+="    <param name=\"Volume\" value=\"70\">"
    str1=str1+="    <param name=\"WindowlessVideo\" value=\"0\">"
    str1=str1+="</OBJECT>"
    objFiles.innerHTML = str1
}
function SelectFile(){
    var arr=showModalDialog('<%=InstallDir & AdminDir%>/Admin_SelectFile.asp?DialogType=media&ChannelID=<%=ChannelID%>', '', 'dialogWidth:820px; dialogHeight:600px; help: no; scroll: yes; status: no');
    if(arr!=null){
        var ss=arr.split('|');
        document.form1.url.value=ss[0];
        var arrContent=ss[0].split('/');
        document.form1.UpFileName.value=ss[0].replace("<%=FilesPath%>", "");
        windowplay();
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
   <LEGEND align=left>视频文件参数</LEGEND>
    <TABLE border="0" cellpadding="0" cellspacing="3">
       <tr><td  height=5></td></tr>
     <tr>
      <td width=350 align='center' id='objFiles'>
           <!-- **********    WindowsPaly开始　********** -->
        <OBJECT id=MediaPlayer1
            codeBase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701standby=
            Loading
            type=application/x-oleobject height=300 width=320
            classid=CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6 VIEWASTEXT>
            <PARAM NAME="URL" value="">
            <param name="AudioStream" value="-1">
            <param name="AutoSize" value="0">
            <param name="AutoStart" value="-1">
            <param name="AnimationAtStart" value="0">
            <param name="AllowScan" value="-1">
            <param name="AllowChangeDisplaySize" value="-1">
            <param name="AutoRewind" value="0">
            <param name="Balance" value="0">
            <param name="BaseURL" value>
            <param name="BufferingTime" value="5">
            <param name="CaptioningID" value>
            <param name="ClickToPlay" value="-1">
            <param name="CursorType" value="0">
            <param name="CurrentPosition" value="-1">
            <param name="CurrentMarker" value="0">
            <param name="DefaultFrame" value>
            <param name="DisplayBackColor" value="0">
            <param name="DisplayForeColor" value="16777215">
            <param name="DisplayMode" value="0">
            <param name="DisplaySize" value="4">
            <param name="Enabled" value="-1">
            <param name="EnableContextMenu" value="-1">
            <param name="EnablePositionControls" value="0">
            <param name="EnableFullScreenControls" value="0">
            <param name="EnableTracker" value="-1">
            <param name="InvokeURLs" value="-1">
            <param name="Language" value="-1">
            <param name="Mute" value="0">
            <param name="PlayCount" value="1">
            <param name="PreviewMode" value="0">
            <param name="Rate" value="1">
            <param name="SAMILang" value>
            <param name="SAMIStyle" value>
            <param name="SAMIFileName" value>
            <param name="SelectionStart" value="-1">
            <param name="SelectionEnd" value="-1">
            <param name="SendOpenStateChangeEvents" value="-1">
            <param name="SendWarningEvents" value="-1">
            <param name="SendErrorEvents" value="-1">
            <param name="SendKeyboardEvents" value="0">
            <param name="SendMouseClickEvents" value="0">
            <param name="SendMouseMoveEvents" value="0">
            <param name="SendPlayStateChangeEvents" value="-1">
            <param name="ShowCaptioning" value="0">
            <param name="ShowControls" value="-1">
            <param name="ShowAudioControls" value="-1">
            <param name="ShowDisplay" value="0">
            <param name="ShowGotoBar" value="0">
            <param name="ShowPositionControls" value="-1">
            <param name="ShowStatusBar" value="-1">
            <param name="ShowTracker" value="-1">
            <param name="TransparentAtStart" value="-1">
            <param name="VideoBorderWidth" value="0">
            <param name="VideoBorderColor" value="0">
            <param name="VideoBorder3D" value="0">
            <param name="Volume" value="70">
            <param name="WindowlessVideo" value="0">
        </OBJECT>
          <!-- **********    WindowsPlay结束　********** -->
       </td>
        </tr>
     <tr><td align='center' height='5'></td></tr>
     <TR>
        <TD >地址：<INPUT name="url" id=url  value="http://" size=40 onChange="javascript:windowplay()">
        <%if IsUpload=True And AdminName <> "" then %>
             <input type="button" name="Submit" value="..." title="从已上传文件中选择" onClick="SelectFile()">
        <%End if%>
        </td>
    </TR>
    <TR>
     <TD>宽度：<INPUT name="width" id=width  ONKEYPRESS="event.returnValue=IsDigit();" value=352 size=7 maxlength="4"  onChange="javascript:windowplay()"> &nbsp;&nbsp;高度：<INPUT id=height ONKEYPRESS="event.returnValue=IsDigit();" value=288 size=7 maxlength="4"  onChange="javascript:windowplay()">
     </TD>
    </TR>
    <TR>
     <TD align=center>支持格式为：mp3、avi、wmv、mpg、asf</TD>
    </TR>
 </TABLE>
 </fieldset>
 </td>
 <td width=80 align="center"><input name="cmdOK" type="button" id="cmdOK" value="  确定  " onClick="OK();">
 <br>
 <br>
 <input name="cmdCancel" type=button id="cmdCancel" onClick="window.close();" value='  取消  '></td></tr>
  <%if IsUpload=True then %>
  <tr>
   <td>
   <FIELDSET align=left>
    <LEGEND align=left>上传本地视频文件</LEGEND>
    <%
        Response.write "<iframe class=""TBGen"" style=""top:2px"" id=""UploadFiles"" src=""upload.asp?DialogType=media"
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

