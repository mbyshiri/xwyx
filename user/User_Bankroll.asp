<!--#include file="User_Bankroll_Code.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

If ShowUserModel = False Then
%>
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
          <td class="user_righttitle"><img src="Images/point2.gif" align="absmiddle">
            <%
            Response.Write "�����ڵ�λ�ã�<a href='../'>" & SiteName & "</a> >> <a href='Index.asp'>��Ա����</a> >> �ʽ���ϸ��ѯ >> "
            ShowType = PE_CLng(Trim(Request("ShowType")))
            Select Case ShowType
            Case 0
                Response.Write "������ϸ��¼"
            Case 1
                Response.Write "���������¼"
            Case 2
                Response.Write "����֧����¼"
            End Select
            %>
          </td>
        </tr>
        <tr>
          <td height="200" valign='top'>
            <br>
            <p align='center'>
            <a href='User_Bankroll.asp'><img src='images/detail_all.jpg' border='0' title='������ϸ��¼'></a>&nbsp;&nbsp;&nbsp;&nbsp;
            <a href='User_Bankroll.asp?ShowType=1'><img src='images/detail_income.jpg' border='0' title='���������¼'></a>&nbsp;&nbsp;&nbsp;&nbsp;
            <a href='User_Bankroll.asp?ShowType=2'><img src='images/detail_payout.jpg' border='0' title='����֧����¼'></a>
            </p>
			<%Call Main()%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<!--#include file="bottom.asp"-->
</body>
</html>
<%
Call CloseConn
%>
<%
Else
    ShowType = PE_CLng(Trim(Request("ShowType")))
    Dim arrstrHtml
    ChannelID = 0
    SkinID = DefaultSkinID
    PageTitle = SiteName & " >> ��Ա����"
    strPageTitle = "�����ڵ�λ�ã�<a href='../'>" & SiteName & "</a> >> <a href='Index.asp'>��Ա����</a> >> �ʽ���ϸ��ѯ >> "
    Select Case ShowType
    Case 0
        strPageTitle = strPageTitle & "������ϸ��¼"
    Case 1
        strPageTitle = strPageTitle & "���������¼"
    Case 2
        strPageTitle = strPageTitle & "����֧����¼"
    End Select
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
    Response.Write arrstrHtml(0)
%>

            <br>
            <p align='center'>
            <a href='User_Bankroll.asp'><img src='images/detail_all.jpg' border='0' title='������ϸ��¼'></a>&nbsp;&nbsp;&nbsp;&nbsp;
            <a href='User_Bankroll.asp?ShowType=1'><img src='images/detail_income.jpg' border='0' title='���������¼'></a>&nbsp;&nbsp;&nbsp;&nbsp;
            <a href='User_Bankroll.asp?ShowType=2'><img src='images/detail_payout.jpg' border='0' title='����֧����¼'></a>
            </p>
			<%Call Main()%>

<%
    Response.Write arrstrHtml(1)
    Call CloseConn
End If
%>