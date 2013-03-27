<!--#include file="CommonCode.asp"-->
<%
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

Dim DownloadUrl, FileExt
DownloadUrl = PE_Content.GetDownloadUrl()
DownloadUrl = Replace(DownloadUrl,  "&nbsp;", " ")
Set PE_Content = Nothing
Call CloseConn
If DownloadUrl = "ErrorDownloadUrl" Then Response.End
FileExt = LCase(Mid(DownloadUrl, InStrRev(DownloadUrl, ".") + 1))
If InStr(DownloadUrl, "://") <= 0 Then
    DownloadUrl = "http://" & Trim(Request.ServerVariables("HTTP_HOST")) & DownloadUrl
End If

Select Case FileExt
Case "wmv", "mpg", "asf", "mp3", "mpeg", "avi"
    ShowMediaPlayer (DownloadUrl)
Case "rm", "ra", "ram"
    ShowRealPlayer (DownloadUrl)
Case Else
    '方法一可以解决下载地址中的中文名在另存为对框中的乱码问题，但只能直接点击下载
    'Response.write "<meta http-equiv='refresh' content=""0;url='" & DownloadUrl & "'"">"
    
    '方法二可以使用“目标另存为”，但会有中文乱码问题
    Response.Redirect DownloadUrl
End Select


Function ShowMediaPlayer(strUrl)
%>
<object id=MediaPlayer1 codeBase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701standby=Loading type=application/x-oleobject height=300 width=320 classid=CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6 VIEWASTEXT>
  <PARAM NAME="URL" value="<%=strUrl%>">
  <param name="AudioStream" value="-1">
  <param name="AutoSize" value="0">
  <param name="AutoStart" value="-1">
  <param name="AnimationAtStart" value="-1">
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
</object>
<%
End Function

Function ShowRealPlayer(strUrl)
%>
<object classid='clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA' width='300' height='220'>
  <param name='CONTROLS' value='ImageWindow'>
  <param name='CONSOLE' value='Clip1'>
  <param name='AUTOSTART' value='-1'>
  <param name='src' value="<%=strUrl%>">
</object>
<br>
<object classid='clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA' width='300' height='60'>
  <param name='CONTROLS' value='ControlPanel,StatusBar'>
  <param name='CONSOLE' value='Clip1'>
</object>
<%
End Function
%>
