<!--#include file="User_Message_Code.asp"-->
<%
If ShowUserModel = False Then
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SiteName & " >> ��Ա����"%></title>
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
Response.Write "�����ڵ�λ�ã�<a href='../'>" & SiteName & "</a> >> <a href='Index.asp'>��Ա����</a> >> վ�ڶ���Ϣ����"
          %></td>
        </tr>
        <tr>
          <td align='center'>
            <table align='center'><tr align='center' valign='top'>
            <td width='80'><a href='User_Message.asp?Action=New'><img src='images/m_new.gif' border='0' title='׫д����Ϣ'><br>׫д����Ϣ</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=Outbox'><img src='images/m_draft.gif' border='0' title='�ݸ���'><br>�ݸ���</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=Inbox'><img src='images/m_box_in.gif' border='0' title='�ռ���'><br>�ռ���</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=IsSend'><img src='images/m_box_out.gif' border='0' title='������'><br>������</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=Recycle'><img src='images/m_box_recycle.gif' border='0' title='�ϼ���'><br>�ϼ���</a></td>
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
    PageTitle = SiteName & " >> ��Ա����"
    strPageTitle = "�����ڵ�λ�ã�<a href='../'>" & SiteName & "</a> >> <a href='Index.asp'>��Ա����</a> >> ��Ϣ����"
    strHtml = GetTemplate(0, 102, 0)
    If strHtml = XmlText("BaseText", "TemplateErr", "�Ҳ���ģ��") Then
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
        Response.Write "��Ա����ͨ��ģ������������ģ������'{$MainContent}'����ο�Ĭ��ģ�塣"
        Response.End		
    End If	
    Response.Write arrstrHtml(0)
%>
            <table align='center'><tr align='center' valign='top'>
            <td width='80'><a href='User_Message.asp?Action=New'><img src='images/m_new.gif' border='0' title='׫д����Ϣ'><br>׫д����Ϣ</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=Outbox'><img src='images/m_draft.gif' border='0' title='�ݸ���'><br>�ݸ���</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=Inbox'><img src='images/m_box_in.gif' border='0' title='�ռ���'><br>�ռ���</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=IsSend'><img src='images/m_box_out.gif' border='0' title='������'><br>������</a></td>
            <td width='80'><a href='User_Message.asp?Action=Manage&ManageType=Recycle'><img src='images/m_box_recycle.gif' border='0' title='�ϼ���'><br>�ϼ���</a></td>
            </tr></table>
<%	
    call Execute
    Response.Write arrstrHtml(1)
    Call CloseConn
End If
%>