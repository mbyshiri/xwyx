<!--#include file="../Include/PowerEasy.Common.Manage.asp"-->
<!--#include file="../Include/PowerEasy.UserInfo.asp"-->

　　<%=UserName%>，您好！欢迎您进入本网站客户自助管理系统！为使您正常使用本系统功能及获得最佳界面浏览效果，请您将IE浏览器升级为6.0，屏幕分辨率设置为1024*768。在这里您可以在本站发布相关信息或接收您的短消息。如果您的信息有所变更，请您及时<a href="User_Info.asp?Action=Modify">修改您的相关信息</a>，以便我们能及时与您取得联系，更好地为您服务！<br>

<%
On Error Resume Next

If UserType = 1 Or UserType = 2 Then
    Call ShowInfo(UserID, True)
Else
    Call ShowInfo(UserID, False)
End If
Response.Write "<br><table width='100%' height='60'><tr align='center'><td>"
Response.Write "    <input class='button1' type='button' name='Submit' value='  修改密码  ' onClick=""window.location.href='User_Info.asp?Action=ModifyPwd'"">"
Response.Write "    <input class='button1' type='button' name='Submit' value='  修改信息  ' onClick=""window.location.href='User_Info.asp?Action=Modify'"">"
If UserType > 0 Then
	If UserType = 1 Then
		If ClientID > 0 Then
			Response.Write "    <input class='button1' type='button' name='Submit' value='查看企业成员' onClick='ShowTabs(5)'>"
		Else
			Response.Write "    <input class='button1' type='button' name='Submit' value='  注销企业  ' onClick=""if(confirm('确定要注销您的企业吗？一旦注销，所有成员都将变成个人会员。')){window.location.href='User_Info.asp?Action=DelCompany'}"">"
		End If
	Else
		Response.Write "    <input class='button1' type='button' name='Submit' value='  退出企业  ' onClick=""if (confirm('确定要退出当前企业吗？')){window.location.href='User_Info.asp?Action=Exit'}"">"
	End If
Else
	Response.Write "    <input class='button1' type='button' name='Submit' value='注册我的企业' onClick=""window.location.href='User_Info.asp?Action=RegCompany'"">"
End If
If UserSetting(18) = 1 Then
    Response.Write "    <input class='button1' type='button' name='Submit' value='  兑换" & PointName & "  ' onClick=""window.location.href='User_Exchange.asp?Action=Exchange'"">"
Else
    Response.Write "    <input class='button1' type='button' name='Submit' value='  兑换" & PointName & "  ' disabled>"
End If
If UserSetting(19) = 1 Then
    Response.Write "    <input class='button1' type='button' name='Submit' value=' 兑换有效期 ' onClick=""window.location.href='User_Exchange.asp?Action=Valid'"">"
Else
    Response.Write "    <input class='button1' type='button' name='Submit' value=' 兑换有效期 ' disabled>"
End If
If NoShow_Shop = False Then
    Response.Write "    <input class='button1' type='button' name='Submit' value=' 充值卡充值 ' onClick=""window.location.href='User_Exchange.asp?Action=Recharge'"">"
Else
    Response.Write "    <input class='button1' type='button' name='Submit' value=' 在线支付 ' onClick=""window.location.href='../PayOnline/PayOnline.asp'"">"
End If
Response.Write "    <input class='button1' type='button' name='Submit' value=' 我的短信息 ' onClick=""window.location.href='User_Message.asp?Action=Manage&ManageType=Inbox'"">"
Response.Write "</td></tr><tr align='center'><td>"
Response.Write "    <input class='button1' type='button' name='Submit' value='  文章签收  ' onClick=""window.location.href='User_Article.asp?Action=Manage&ManageType=Receive&Passed=All'"">"
If NoShow_Shop = False Then
    Response.Write "    <input class='button1' type='button' name='Submit' value='查看订单信息' onClick=""window.location.href='User_Order.asp'"">"
End If
Response.Write "    <input class='button1' type='button' name='Submit' value='查看资金明细' onClick=""window.location.href='User_Bankroll.asp'"">"
Response.Write "    <input class='button1' type='button' name='Submit' value='查看" & PointName & "明细' onClick=""window.location.href='User_ConsumeLog.asp'"">"
Response.Write "    <input class='button1' type='button' name='Submit' value=' 有效期明细 ' onClick=""window.location.href='User_RechargeLog.asp'"">"
Response.Write "    <input class='button1' type='button' name='Submit' value='查看收藏明细' onClick=""window.location.href='User_Favorite.asp'"">"
Response.Write "    <input class='button1' type='button' name='Submit' value='查看在线支付' onClick=""window.location.href='User_Payment.asp'"">"
Response.Write "    </td>"
Response.Write "  </tr>"
Response.Write "</table>"
Response.Write "<br><br><br>"
%>
