<%@language=vbscript codepage=936 %>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>动易目录管理工具</title>
</head>
<frameset framespacing="0" border="false" rows="60,*" frameborder="0" scrolling="yes">
    <frame name="UploadFile_top" scrolling="no" src="Admin_UploadFile_Top.asp?ChannelID=<%=request("ChannelID")%>&UploadDir=<%=request("UploadDir")%>">
    <frameset rows="*" cols="0,*" framespacing="0" frameborder="0" border="false" id="frame" scrolling="yes">
        <frame name="UploadFile_left" scrolling="auto" src="Admin_UploadFile_Left.asp?ChannelID=<%=request("ChannelID")%>&UploadDir=<%=request("UploadDir")%>">
        <frame name="UploadFile_Main" scrolling="auto" src="Admin_UploadFile_Main.asp?ChannelID=<%=request("ChannelID")%>&UploadDir=<%=request("UploadDir")%>">
    </frameset>
</frameset>
<noframes>
  <body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
  <p>你的浏览器版本过低！！！本系统要求IE5及以上版本才能使用本系统。</p>
  </body>
</noframes>
</html>
