<!--#include file="User_Message_Code.asp"-->
<%
If ShowUserModel = False Then
%>
<html>
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
Response.Write "您现在的位置：<a href='../'>" & SiteName & "</a> >> <a href='Index.asp'>会员中心</a> >> 站内短消息管理"
          %></td>
        </tr>
        <tr>
          <td align='center'>
            <table align='center'><tr align='center' valign='top'>
            <td width='80'><a href='User_Message.asp?Action=New'><img src='images/m_new.gif' border='0' title='撰写短消息'><br>撰写短消息</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=Outbox'><img src='images/m_draft.gif' border='0' title='草稿箱'><br>草稿箱</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=Inbox'><img src='images/m_box_in.gif' border='0' title='收件箱'><br>收件箱</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=IsSend'><img src='images/m_box_out.gif' border='0' title='发件箱'><br>发件箱</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=Recycle'><img src='images/m_box_recycle.gif' border='0' title='废件箱'><br>废件箱</a></td>
            </tr></table>
          </td>
        </tr>
        <tr>
          <td height="200" valign='top'>
<%
Call Execute
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
    strPageTitle = "您现在的位置：<a href='../'>" & SiteName & "</a> >> <a href='Index.asp'>会员中心</a> >> 信息管理"
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
    If Instr(strHtml,"{$MainContent}") = 0 Then 
        Response.Write "会员中心通用模板里面必须包含模板内容'{$MainContent}'，请参考默认模板。"
        Response.End		
    End If	
    Response.Write arrstrHtml(0)
%>
            <table align='center'><tr align='center' valign='top'>
            <td width='80'><a href='User_Message.asp?Action=New'><img src='images/m_new.gif' border='0' title='撰写短消息'><br>撰写短消息</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=Outbox'><img src='images/m_draft.gif' border='0' title='草稿箱'><br>草稿箱</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=Inbox'><img src='images/m_box_in.gif' border='0' title='收件箱'><br>收件箱</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=IsSend'><img src='images/m_box_out.gif' border='0' title='发件箱'><br>发件箱</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=Recycle'><img src='images/m_box_recycle.gif' border='0' title='废件箱'><br>废件箱</a></td>
            </tr></table>
<%	
    call Execute
    Response.Write arrstrHtml(1)
    Call CloseConn
End If
%>