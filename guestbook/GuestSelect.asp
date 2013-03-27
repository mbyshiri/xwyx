<html><head><title>请选择头像</title></head>
<script>
window.focus();
function changeimage(imagename){
  window.opener.document.myform.GuestImages.value=imagename;
  window.opener.document.myform.showimages.src='../guestbook/images/face/'+imagename+'.gif';
}
</script>
<body>

<table align="center" width="95%" cellpadding="5"><tr><td>
<%
Dim i
For i = 1 To 22
    Response.Write "<img src='images/face/"
    If i < 10 Then
        i = "0" & i
    End If
    Response.Write i & ".gif' border=0 onclick=""changeimage('" & i & "') "" style=cursor:hand>"
    If i Mod 5 = 0 Then
        Response.Write "<br>"
    End If
Next
%>
</td></tr></table>

<div align='center'><font size='2'>【<a href='javascript:window.close();'>关闭窗口</a>】</font></div>
</body>
</html>