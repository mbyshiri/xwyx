<!--#include file="User_RechargeLog_Code.asp"-->
<%
If ShowUserModel = False Then
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SiteName & " >> 会员中心"%></title>
<link href="../Skin/DefaultSkin.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="top.asp"-->
<table width="756" border="0" align="center" cellpadding="0" cellspacing="0" class="user_border">
  <tr>
    <td valign="top">
      <table width="100%" border="0" cellpadding="5" cellspacing="0" class="user_box">
        <tr>
          <td class="user_righttitle"><img src="Images/point2.gif" align="absmiddle"><%
Response.Write "您现在的位置：<a href='../'>" & SiteName & "</a> >> <a href='Index.asp'>会员中心</a> >> 有效期充值明细查询"
          %></td>
        </tr>
        <tr>
          <td height="200" valign='top'>
<%
Call Main
%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<!--#include file="bottom.asp"-->
</body>
</html>
<% Call CloseConn %>
<%
Else
    Dim arrstrHtml
    ChannelID = 0
    SkinID = DefaultSkinID
    PageTitle = SiteName & " >> 会员中心"
    strPageTitle = "您现在的位置：<a href='../'>" & SiteName & "</a> >> <a href='Index.asp'>会员中心</a> >> 有效期充值明细查询"
    strHtml = GetTemplate(0, 102, 0)
    If strHtml = XmlText("BaseText", "TemplateErr", "找不到模板") Then
        Response.Write strHtml
        Response.End	        
    End If	
    strHtml = Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))
    strHtml = Replace(strHtml, "{$MenuJS}", GetMenuJS("", False))
    strHtml = Replace(strHtml, "{$ShowMessageBox}", ShowMessageBox())

    Call ReplaceCommonLabel
    strHtml = Replace(strHtml, "{$PageTitle}", SiteTitle & " >> " & PageTitle)
    strHtml = Replace(strHtml, "{$UserMenu}", UserMenu())
    strHtml = Replace(strHtml, "{$ShowPath}", strPageTitle)
    arrstrHtml = Split(strHtml,"{$MainContent}")
    Response.Write arrstrHtml(0)
    Call Main
    Response.Write arrstrHtml(1)
    Call CloseConn
End If
%>