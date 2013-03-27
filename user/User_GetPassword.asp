<!--#include file="User_GetPassword_Code.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SiteName & " >> 会员中心 >> 找回密码"%></title>
<link href="../Skin/DefaultSkin.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="756" border="0" align="center" cellpadding="0" cellspacing="0" class="user_border">
  <tr>
    <td valign="top">
<%
Call Execute
%>
    </td>
  </tr>
</table>
</body>
</html>
<%
Call CloseConn
%>