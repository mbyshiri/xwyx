<!--#include file="../Include/PowerEasy.Common.Manage.asp"-->
<!--#include file="../Include/PowerEasy.UserInfo.asp"-->

����<%=UserName%>�����ã���ӭ�����뱾��վ�ͻ���������ϵͳ��Ϊʹ������ʹ�ñ�ϵͳ���ܼ������ѽ������Ч����������IE���������Ϊ6.0����Ļ�ֱ�������Ϊ1024*768���������������ڱ�վ���������Ϣ��������Ķ���Ϣ�����������Ϣ���������������ʱ<a href="User_Info.asp?Action=Modify">�޸����������Ϣ</a>���Ա������ܼ�ʱ����ȡ����ϵ�����õ�Ϊ������<br>

<%
On Error Resume Next

If UserType = 1 Or UserType = 2 Then
    Call ShowInfo(UserID, True)
Else
    Call ShowInfo(UserID, False)
End If
Response.Write "<br><table width='100%' height='60'><tr align='center'><td>"
Response.Write "    <input class='button1' type='button' name='Submit' value='  �޸�����  ' onClick=""window.location.href='User_Info.asp?Action=ModifyPwd'"">"
Response.Write "    <input class='button1' type='button' name='Submit' value='  �޸���Ϣ  ' onClick=""window.location.href='User_Info.asp?Action=Modify'"">"
If UserType > 0 Then
	If UserType = 1 Then
		If ClientID > 0 Then
			Response.Write "    <input class='button1' type='button' name='Submit' value='�鿴��ҵ��Ա' onClick='ShowTabs(5)'>"
		Else
			Response.Write "    <input class='button1' type='button' name='Submit' value='  ע����ҵ  ' onClick=""if(confirm('ȷ��Ҫע��������ҵ��һ��ע�������г�Ա������ɸ��˻�Ա��')){window.location.href='User_Info.asp?Action=DelCompany'}"">"
		End If
	Else
		Response.Write "    <input class='button1' type='button' name='Submit' value='  �˳���ҵ  ' onClick=""if (confirm('ȷ��Ҫ�˳���ǰ��ҵ��')){window.location.href='User_Info.asp?Action=Exit'}"">"
	End If
Else
	Response.Write "    <input class='button1' type='button' name='Submit' value='ע���ҵ���ҵ' onClick=""window.location.href='User_Info.asp?Action=RegCompany'"">"
End If
If UserSetting(18) = 1 Then
    Response.Write "    <input class='button1' type='button' name='Submit' value='  �һ�" & PointName & "  ' onClick=""window.location.href='User_Exchange.asp?Action=Exchange'"">"
Else
    Response.Write "    <input class='button1' type='button' name='Submit' value='  �һ�" & PointName & "  ' disabled>"
End If
If UserSetting(19) = 1 Then
    Response.Write "    <input class='button1' type='button' name='Submit' value=' �һ���Ч�� ' onClick=""window.location.href='User_Exchange.asp?Action=Valid'"">"
Else
    Response.Write "    <input class='button1' type='button' name='Submit' value=' �һ���Ч�� ' disabled>"
End If
If NoShow_Shop = False Then
    Response.Write "    <input class='button1' type='button' name='Submit' value=' ��ֵ����ֵ ' onClick=""window.location.href='User_Exchange.asp?Action=Recharge'"">"
Else
    Response.Write "    <input class='button1' type='button' name='Submit' value=' ����֧�� ' onClick=""window.location.href='../PayOnline/PayOnline.asp'"">"
End If
Response.Write "    <input class='button1' type='button' name='Submit' value=' �ҵĶ���Ϣ ' onClick=""window.location.href='User_Message.asp?Action=Manage&ManageType=Inbox'"">"
Response.Write "</td></tr><tr align='center'><td>"
Response.Write "    <input class='button1' type='button' name='Submit' value='  ����ǩ��  ' onClick=""window.location.href='User_Article.asp?Action=Manage&ManageType=Receive&Passed=All'"">"
If NoShow_Shop = False Then
    Response.Write "    <input class='button1' type='button' name='Submit' value='�鿴������Ϣ' onClick=""window.location.href='User_Order.asp'"">"
End If
Response.Write "    <input class='button1' type='button' name='Submit' value='�鿴�ʽ���ϸ' onClick=""window.location.href='User_Bankroll.asp'"">"
Response.Write "    <input class='button1' type='button' name='Submit' value='�鿴" & PointName & "��ϸ' onClick=""window.location.href='User_ConsumeLog.asp'"">"
Response.Write "    <input class='button1' type='button' name='Submit' value=' ��Ч����ϸ ' onClick=""window.location.href='User_RechargeLog.asp'"">"
Response.Write "    <input class='button1' type='button' name='Submit' value='�鿴�ղ���ϸ' onClick=""window.location.href='User_Favorite.asp'"">"
Response.Write "    <input class='button1' type='button' name='Submit' value='�鿴����֧��' onClick=""window.location.href='User_Payment.asp'"">"
Response.Write "    </td>"
Response.Write "  </tr>"
Response.Write "</table>"
Response.Write "<br><br><br>"
%>
