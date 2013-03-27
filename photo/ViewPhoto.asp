<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>图片内容</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%'>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td height='380' align='center' valign='middle'>" & vbCrLf

Dim PhotoUrl
PhotoUrl = LCase(Trim(request("PhotoUrl")))
If PhotoUrl = "" Then
    PhotoUrl = "images/nopic.gif"
End If
Select Case LCase(Mid(PhotoUrl, InStrRev(PhotoUrl, ".") + 1))
Case "swf"
    Response.Write "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' width='640' height='400' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0'   name='images1'><param name='movie' value='" & PhotoUrl & "'><param name='quality' value='high'><embed src='" & PhotoUrl & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='550' height='400'></embed></object>"
Case "gif", "jpg", "jpeg", "jpe", "bmp", "png"
    Response.Write "<img name='images1' src='" & PhotoUrl & "' border='0'>"
Case Else
    Response.Redirect PhotoUrl
End Select

Response.Write "      <input type='hidden' name='PhotoUrl' value='" & PhotoUrl & "'>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case LCase(Mid(PhotoUrl, InStrRev(PhotoUrl, ".") + 1))
Case "swf"
    Response.Write "<div id='hiddenPic' style='position:absolute; left:0px; top:0px; width:100px; height:100px; z-index:-1; visibility: hidden;'>"
    Response.Write "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' width='640' height='400' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0'   name='images2'><param name='movie' value='" & PhotoUrl & "'><param name='quality' value='high'><embed src='" & PhotoUrl & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='550' height='400'></embed></object>"
    Response.Write "</div>"
Case "gif", "jpg", "jpeg", "jpe", "bmp", "png"
    Response.Write "<div id='hiddenPic' style='position:absolute; left:0px; top:0px; width:100px; height:100px; z-index:-1; visibility: hidden;'>"
    Response.Write "<img name='images2' src='" & PhotoUrl & "' border='0'>"
    Response.Write "</div>"
End Select

Response.Write "</body>" & vbCrLf
Response.Write "</html>" & vbCrLf
%>
