<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 0      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

strFileName = "Admin_ShowPurview.asp"

Response.Write "<html><head><title>�鿴����Ȩ��</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Response.Write "  <tr class='topbg'>" & vbCrLf
Response.Write "    <td height='22' colspan='2' align='center'><strong>�� �� �� �� Ȩ ��</strong></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>��������</strong></td>" & vbCrLf
Response.Write "    <td> <a href='Admin_ShowPurview.asp'>����Ȩ����ҳ</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
Response.Write "  <tr class='title'>"
Response.Write "    <td height='22'>" & GetChannelList() & "</td>"
Response.Write "  </tr>"
Response.Write "</table><br>"
Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
Response.Write "  <tr>"
Response.Write "    <td height='22'>" & GetManagePath() & "</td>"
Response.Write "  </tr>"
Response.Write "</table>"


If ChannelID = 0 Then
    Call ShowAllPurview
ElseIf ChannelID = 4 Then
    Call ShowGuestBookPurview
Else
    Call ShowChannelPurview
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub ShowAllPurview()

    Dim rsChannel, sqlChannel, rsAdmin, Channel_Purview
    sqlChannel = "select * from PE_Channel where ChannelType<=1 and ChannelID<>4 and Disabled=" & PE_False & " order by OrderID"
    Set rsChannel = Server.CreateObject("adodb.recordset")
    rsChannel.Open sqlChannel, Conn, 1, 1
    Do While Not rsChannel.EOF
        Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Set rsAdmin = Conn.Execute("select AdminPurview_" & rsChannel("ChannelDir") & " from PE_Admin where AdminName='" & AdminName & "'")
        If Not (rsAdmin.BOF And rsAdmin.EOF) Then
            Channel_Purview = rsAdmin(0)
        End If
        rsAdmin.Close
        Set rsAdmin = Nothing
        Response.Write "  <tr class='title' height='22'>"
        Response.Write "    <td colspan='4'><strong>" & rsChannel("ChannelName") & "</strong> "
        If Channel_Purview = 1 Then Response.Write "��Ƶ������Ա��"
        If Channel_Purview = 2 Then Response.Write "����Ŀ�ܱࣩ"
        If Channel_Purview = 3 Then Response.Write "����Ŀ����Ա��"
        If Channel_Purview = 4 Then Response.Write "����Ȩ�ޣ�"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='30%'>����Ŀ" & rsChannel("ChannelShortName") & "¼�롢��ˡ�����Ȩ��</td>"
        Response.Write "    <td align='center' width='20%'>"
        If Channel_Purview <= 2 Then
            Response.Write "<font color=blue>ȫ��Ȩ��</font>"
        ElseIf Channel_Purview = 3 Then
            Response.Write "<a href='Admin_ShowPurview.asp?iChannelID=" & rsChannel("ChannelID") & "'><font color=blue>����Ȩ��</font></a>"
        Else
            Response.Write "<font color=red>��Ȩ��</font>"
        End If
        Response.Write "    </td>"
        Response.Write "    <td width='30%'>ר��" & rsChannel("ChannelShortName") & "����Ȩ��</td>"
        Response.Write "    <td align='center' width='20%'>"
        If AdminPurview = 1 Or Channel_Purview <= 2 Then
            Response.Write "<font color=blue>��</font>"
        Else
            Response.Write "<font color=red>��</font>"
        End If
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='30%'>��Ŀ����ר��������ɹ���Ȩ��</td>"
        Response.Write "    <td align='center' width='20%'>"
        If AdminPurview = 1 Or Channel_Purview = 1 Then
            Response.Write "<font color=blue>��</font>"
        Else
            Response.Write "<font color=red>��</font>"
        End If
        Response.Write "    </td>"
        Response.Write "    <td width='30%'>" & rsChannel("ChannelShortName") & "���ۡ�����վ����������Ȩ��</td>"
        Response.Write "    <td align='center' width='20%'>"
        If AdminPurview = 1 Or Channel_Purview = 1 Then
            Response.Write "<font color=blue>��</font>"
        Else
            Response.Write "<font color=red>��</font>"
        End If
        Response.Write "    </td>"
        Response.Write "  </tr>"

        Response.Write "  <tr>"
        Response.Write "    <td class='tdbg' colspan='4'>"
        Response.Write "<b>����Ȩ�ޣ�</b><br>"
        Response.Write "ģ�����&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Template_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        Response.Write "JS�ļ�����&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "JsFile_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        Response.Write "�����˵�&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Keyword_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        Response.Write "�ؼ��ֹ���&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Template_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        If rsChannel("ModuleType") = 5 Then
            Response.Write "���̹���&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Producer_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
            Response.Write "Ʒ�ƹ���&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Trademark_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        Else
            Response.Write "���߹���&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Author_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
            Response.Write "��Դ����&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Copyfrom_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        End If
        Response.Write "����XML&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "XML_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        Response.Write "�Զ����ֶ�&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Field_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        Response.Write "������&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "AD_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "</table>"
        Response.Write "<br>"
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing

    Dim rsGuestBook, sqlGuestBook, rsAdminGuest, GuestBook_Purview
    sqlGuestBook = "select * from PE_Channel where ChannelType<=1 and ChannelID=4 and Disabled=" & PE_False & ""
    Set rsGuestBook = Server.CreateObject("adodb.recordset")
    rsGuestBook.Open sqlGuestBook, Conn, 1, 1
    If Not (rsGuestBook.EOF And rsGuestBook.BOF) Then
        Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Set rsAdminGuest = Conn.Execute("select AdminPurview_GuestBook from PE_Admin where AdminName='" & AdminName & "'")
        If Not (rsAdminGuest.BOF And rsAdminGuest.EOF) Then
            GuestBook_Purview = rsAdminGuest(0)
        End If
        rsAdminGuest.Close
        Set rsAdminGuest = Nothing
        Response.Write "  <tr class='title' height='22'>"
        Response.Write "    <td colspan='4'><strong>" & rsGuestBook("ChannelName") & "</strong> "
        If GuestBook_Purview = 1 Then Response.Write "��Ƶ������Ա��"
        If GuestBook_Purview = 2 Then Response.Write "����Ŀ�ܱࣩ"
        If GuestBook_Purview = 3 Then Response.Write "����Ŀ����Ա��"
        If GuestBook_Purview = 4 Then Response.Write "����Ȩ�ޣ�"
        Response.Write "    </td>"
        Response.Write "  </tr>"

        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='30%'>����Ŀ" & rsGuestBook("ChannelShortName") & "�޸ġ�ɾ�����ƶ�����ˡ��������̶����ظ�Ȩ��</td>"
        Response.Write "    <td align='center' width='20%'>"
        If GuestBook_Purview <= 2 Then
            Response.Write "<font color=blue>ȫ��Ȩ��</font>"
        ElseIf GuestBook_Purview = 3 Then
            Response.Write "<a href='Admin_ShowPurview.asp?iChannelID=" & rsGuestBook("ChannelID") & "'><font color=blue>����Ȩ��</font></a>"
        Else
            Response.Write "<font color=red>��Ȩ��</font>"
        End If
        Response.Write "    </td>"
        Response.Write "    <td width='30%'>���Թ���" & rsGuestBook("ChannelShortName") & "���</td>"
        Response.Write "    <td align='center' width='20%'>"
        If AdminPurview = 1 Or GuestBook_Purview <= 2 Then
            Response.Write "<font color=blue>��</font>"
        Else
            Response.Write "<font color=red>��</font>"
        End If
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='30%'>��Ŀ��������ִ����ҳǶ���������</td>"
        Response.Write "    <td align='center' width='20%'>"
        If AdminPurview = 1 Or GuestBook_Purview = 1 Then
            Response.Write "<font color=blue>��</font>"
        Else
            Response.Write "<font color=red>��</font>"
        End If
        Response.Write "    </td>"
        Response.Write "    <td width='30%'>" & rsGuestBook("ChannelShortName") & "��������Ȩ��</td>"
        Response.Write "    <td align='center' width='20%'>"
        If AdminPurview = 1 Or GuestBook_Purview = 1 Then
            Response.Write "<font color=blue>��</font>"
        Else
            Response.Write "<font color=red>��</font>"
        End If
        Response.Write "    </td>"
        Response.Write "  </tr>"

        Response.Write "  <tr>"
        Response.Write "    <td class='tdbg' colspan='4'>"
        Response.Write "<b>����Ȩ�ޣ�</b><br>"
        Response.Write "������&nbsp;" & ShowChannelOtherPurview(GuestBook_Purview, "AD_" & rsGuestBook("ChannelDir")) & "&nbsp;&nbsp;"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "</table>"
        Response.Write "<br>"
    End If
    rsGuestBook.Close
    Set rsGuestBook = Nothing
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td colspan='6' height='22'><strong>������վ����Ȩ��</strong><strong> </strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>�޸��Լ����� " & ShowPurview("ModifyPwd") & "</td>"
    Response.Write "    <td width='16%'>��վƵ������ " & ShowPurview("Channel") & "</td>"
    Response.Write "    <td width='16%'>�ɼ����� " & ShowPurview("Collection") & "</td>"
    Response.Write "    <td width='16%'>����Ϣ���� " & ShowPurview("Message") & "</td>"
    Response.Write "    <td width='16%'>�ʼ��б���� " & ShowPurview("MailList") & "</td>"
    Response.Write "    <td width='16%'>��վ������ " & ShowPurview("AD") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>�������ӹ��� " & ShowPurview("FriendSite") & "</td>"
    Response.Write "    <td width='16%'>��վ������� " & ShowPurview("Announce") & "</td>"
    Response.Write "    <td width='16%'>��վ������� " & ShowPurview("Vote") & "</td>"
    Response.Write "    <td width='16%'>��վͳ�ƹ��� " & ShowPurview("Counter") & "</td>"
    Response.Write "    <td width='16%'>��վ������ " & ShowPurview("Skin") & "</td>"
    Response.Write "    <td width='16%'>ͨ��ģ����� " & ShowPurview("Template") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>�Զ����ǩ���� " & ShowPurview("Label") & "</td>"
    Response.Write "    <td width='16%'>��վ������� " & ShowPurview("Cache") & "</td>"
    Response.Write "    <td width='16%'>վ�����ӹ��� " & ShowPurview("KeyLink") & "</td>"
    Response.Write "    <td width='16%'>�ַ����˹��� " & ShowPurview("Rtext") & "</td>"
    Response.Write "    <td width='16%'>��Ա����� " & ShowPurview("UserGroup") & "</td>"
    Response.Write "    <td width='16%'>��ֵ������ " & ShowPurview("Card") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>�ҳ��Ǽǹ��� " & ShowPurview("Equipment") & "</td>"
    Response.Write "    <td width='16%'>ѧ����Ϣ���� " & ShowPurview("InfoManage") & "</td>"
    Response.Write "    <td width='16%'>ѧ���ɼ����� " & ShowPurview("ScoreManage") & "</td>"
    Response.Write "    <td width='16%'>���Թ��� " & ShowPurview("TestManage") & "</td>"
    Response.Write "    <td width='16%'></td>"
    Response.Write "    <td width='16%'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td colspan='6' height='22'><strong>��Ա����Ȩ��</strong><strong> </strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>�鿴��Ա��Ϣ " & ShowPurview("User_View") & "</td>"
    Response.Write "    <td width='16%'>�޸Ļ�Ա��Ϣ " & ShowPurview("User_ModifyInfo") & "</td>"
    Response.Write "    <td width='16%'>�޸Ļ�ԱȨ�� " & ShowPurview("User_MofidyPurview") & "</td>"
    Response.Write "    <td width='16%'>��ס/������Ա " & ShowPurview("User_Lock") & "</td>"
    Response.Write "    <td width='16%'>ɾ����Ա " & ShowPurview("User_Del") & "</td>"
    Response.Write "    <td width='16%'>����Ϊ�ͻ� " & ShowPurview("User_Update") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>��Ա�ʽ���� " & ShowPurview("User_Money") & "</td>"
    Response.Write "    <td width='16%'>��Ա��ȯ���� " & ShowPurview("User_Point") & "</td>"
    Response.Write "    <td width='16%'>��Ա��Ч�ڹ��� " & ShowPurview("User_Valid") & "</td>"
    Response.Write "    <td width='16%'>��Ա������ϸ " & ShowPurview("ConsumeLog") & "</td>"
    Response.Write "    <td width='16%'>��Ա��Ч����ϸ " & ShowPurview("RechargeLog") & "</td>"
    Response.Write "    <td width='16%'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td colspan='6' height='22'><strong>�̳��ճ���������Ȩ��</strong><strong> </strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>�鿴���� " & ShowPurview("Order_View") & "</td>"
    Response.Write "    <td width='16%'>ȷ�϶��� " & ShowPurview("Order_Confirm") & "</td>"
    Response.Write "    <td width='16%'>�޸Ķ��� " & ShowPurview("Order_Modify") & "</td>"
    Response.Write "    <td width='16%'>ɾ������ " & ShowPurview("Order_Del") & "</td>"
    Response.Write "    <td width='16%'>�տ�� " & ShowPurview("Order_Payment") & "</td>"
    Response.Write "    <td width='16%'>����Ʊ " & ShowPurview("Order_Invoice") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>�������ͣ�ʵ� " & ShowPurview("Order_Deliver") & "</td>"
    Response.Write "    <td width='16%'>�������ͣ������ " & ShowPurview("Order_Download") & "</td>"
    Response.Write "    <td width='16%'>�������ͣ��㿨�� " & ShowPurview("Order_SendCard") & "</td>"
    Response.Write "    <td width='16%'>���嶩�� " & ShowPurview("Order_End") & "</td>"
    Response.Write "    <td width='16%'>�������� " & ShowPurview("Order_Transfer") & "</td>"
    Response.Write "    <td width='16%'>������ӡ " & ShowPurview("Order_Print") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>����ͳ�� " & ShowPurview("Order_Count") & "</td>"
    Response.Write "    <td width='16%'>������ϸ��� " & ShowPurview("Order_OrderItem") & "</td>"
    Response.Write "    <td width='16%'>����ͳ��/���� " & ShowPurview("Order_SaleCount") & "</td>"
    Response.Write "    <td width='16%'>����֧������ " & ShowPurview("Payment") & "</td>"
    Response.Write "    <td width='16%'>�ʽ���ϸ��ѯ " & ShowPurview("Bankroll") & "</td>"
    Response.Write "    <td width='16%'>���˻���¼ " & ShowPurview("Deliver") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>����������¼ " & ShowPurview("Transfer") & "</td>"
    Response.Write "    <td width='16%'>������������ " & ShowPurview("PresentProject") & "</td>"
    Response.Write "    <td width='16%'>���ʽ���� " & ShowPurview("PaymentType") & "</td>"
    Response.Write "    <td width='16%'>�ͻ���ʽ���� " & ShowPurview("DeliverType") & "</td>"
    Response.Write "    <td width='16%'>�����ʻ����� " & ShowPurview("Bank") & "</td>"
    Response.Write "    <td width='16%'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td colspan='5' height='22'><strong>�ͻ���ϵ����Ȩ��</strong><strong> </strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='17%'>�鿴�ͻ���Ϣ " & ShowPurview("Client_View") & "</td>"
    Response.Write "    <td width='17%'>��ӿͻ� " & ShowPurview("Client_Add") & "</td>"
    Response.Write "    <td width='25%'>�޸������Լ��Ŀͻ���Ϣ " & ShowPurview("Client_ModifyOwn") & "</td>"
    Response.Write "    <td width='25%'>�޸����пͻ���Ϣ " & ShowPurview("Client_ModifyAll") & "</td>"
    Response.Write "    <td width='17%'>ɾ���ͻ� " & ShowPurview("Client_Del") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='17%'>�鿴�����¼ " & ShowPurview("Service_View") & "</td>"
    Response.Write "    <td width='17%'>��ӷ����¼ " & ShowPurview("Service_Add") & "</td>"
    Response.Write "    <td width='25%'>�޸��Լ���ӵķ����¼ " & ShowPurview("Service_ModifyOwn") & "</td>"
    Response.Write "    <td width='25%'>�޸����з����¼ " & ShowPurview("Service_ModifyAll") & "</td>"
    Response.Write "    <td width='17%'>ɾ�������¼ " & ShowPurview("Service_Del") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='17%'>�鿴Ͷ�߼�¼ " & ShowPurview("Complain_View") & "</td>"
    Response.Write "    <td width='17%'>���Ͷ�߼�¼ " & ShowPurview("Complain_Add") & "</td>"
    Response.Write "    <td width='25%'>�޸��Լ���ӵ�Ͷ�߼�¼ " & ShowPurview("Complain_ModifyOwn") & "</td>"
    Response.Write "    <td width='25%'>�޸�����Ͷ�߼�¼ " & ShowPurview("Complain_ModifyAll") & "</td>"
    Response.Write "    <td width='17%'>ɾ��Ͷ�߼�¼ " & ShowPurview("Complain_Del") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='17%'>�鿴�طü�¼ " & ShowPurview("Call_View") & "</td>"
    Response.Write "    <td width='17%'>��ӻطü�¼ " & ShowPurview("Call_Add") & "</td>"
    Response.Write "    <td width='25%'>�޸��Լ���ӵĻطü�¼ " & ShowPurview("Call_ModifyOwn") & "</td>"
    Response.Write "    <td width='25%'>�޸����лطü�¼ " & ShowPurview("Call_ModifyAll") & "</td>"
    Response.Write "    <td width='17%'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

End Sub

Sub ShowGuestBookPurview()
    Dim rsAdminGuest, GuestBook_Purview, arrKind_GuestBook
    Dim rsGuestBook, GuestBookDir, GuestBookName, GuestBookShortName

    Set rsGuestBook = Conn.Execute("select * from PE_Channel where ChannelID=4")
    If Not (rsGuestBook.BOF Or rsGuestBook.EOF) Then
        GuestBookDir = rsGuestBook("ChannelDir")
        GuestBookName = rsGuestBook("ChannelName")
        GuestBookShortName = rsGuestBook("ChannelShortName")
        Set rsAdminGuest = Conn.Execute("select AdminPurview_GuestBook,arrClass_GuestBook from PE_Admin where AdminName='" & AdminName & "'")
        If Not (rsAdminGuest.BOF And rsAdminGuest.EOF) Then
            GuestBook_Purview = rsAdminGuest(0)
            arrKind_GuestBook = Split(rsAdminGuest(1), "|||")
        End If
        rsAdminGuest.Close
        Set rsAdminGuest = Nothing
    End If
    rsGuestBook.Close
    Set rsGuestBook = Nothing

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td colspan='2' height='22'><strong>" & GuestBookName & "</strong> "
    If GuestBook_Purview = 1 Then Response.Write "��Ƶ������Ա��"
    If GuestBook_Purview = 2 Then Response.Write "����Ŀ�ܱࣩ"
    If GuestBook_Purview = 3 Then Response.Write "����Ŀ����Ա��"
    If GuestBook_Purview = 4 Then Response.Write "����Ȩ�ޣ�"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>��Ŀ����Ȩ�ޣ�����ִ����ҳǶ���������</td>"
    Response.Write "    <td align='center' width='20%'>"
    If AdminPurview = 1 Or GuestBook_Purview = 1 Then
        Response.Write "<font color=blue>��</font>"
    Else
        Response.Write "<font color=red>��</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>���Թ���" & GuestBookShortName & "���</td>"
    Response.Write "    <td align='center' width='20%'>"
    If AdminPurview = 1 Or GuestBook_Purview <= 2 Then
        Response.Write "<font color=blue>��</font>"
    Else
        Response.Write "<font color=red>��</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>" & GuestBookShortName & "��������Ȩ��</td>"
    Response.Write "    <td align='center' width='20%'>"
    If AdminPurview = 1 Or GuestBook_Purview = 1 Then
        Response.Write "<font color=blue>��</font>"
    Else
        Response.Write "<font color=red>��</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>����Ŀ" & GuestBookShortName & "�޸ġ�ɾ�����ƶ�����ˡ��������̶����ظ�Ȩ��</td>"
    Response.Write "    <td align='center' width='20%'>"
    If GuestBook_Purview <= 2 Then
        Response.Write "<font color=blue>ȫ��Ȩ��</font>"
    ElseIf GuestBook_Purview = 3 Then
        Response.Write "<font color=blue>����Ȩ��</font>"
    Else
        Response.Write "<font color=red>��Ȩ��</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22'><strong>����Ȩ�ޣ�</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td>"
    Response.Write "������&nbsp;" & ShowPurview("AD_" & GuestBookDir) & "&nbsp;&nbsp;"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

    If GuestBook_Purview = 3 Then
        Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
        Response.Write "  <tr align='center' class='title'>"
        Response.Write "    <td height='22'><strong>��Ŀ����</strong></td>"
        Response.Write "    <td width='70'><strong>�޸�</strong></td>"
        Response.Write "    <td width='70'><strong>ɾ��</strong></td>"
        Response.Write "    <td width='70'><strong>�ƶ�</strong></td>"
        Response.Write "    <td width='70'><strong>���</strong></td>"
        Response.Write "    <td width='70'><strong>����</strong></td>"
        Response.Write "    <td width='70'><strong>�̶�</strong></td>"
        Response.Write "    <td width='70'><strong>�ظ�</strong></td>"
        Response.Write "  </tr>"
        Dim rsGuestKind
        Set rsGuestKind = Conn.Execute("select * from PE_GuestKind order by OrderID,KindID")
        Do While Not rsGuestKind.EOF
            Response.Write "  <tr class='tdbg'>"
            Response.Write "    <td align='center'>" & rsGuestKind("KindName") & "</td>"
            Response.Write "    <td align='center'>"
            If FoundInArr(arrKind_GuestBook(0), rsGuestKind("KindID"), ",") = True Then
                Response.Write "<font color=blue>��</font>"
            Else
                Response.Write "<font color=red>��</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrKind_GuestBook(1), rsGuestKind("KindID"), ",") = True Then
                Response.Write "<font color=blue>��</font>"
            Else
                Response.Write "<font color=red>��</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrKind_GuestBook(2), rsGuestKind("KindID"), ",") = True Then
                Response.Write "<font color=blue>��</font>"
            Else
                Response.Write "<font color=red>��</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrKind_GuestBook(3), rsGuestKind("KindID"), ",") = True Then
                Response.Write "<font color=blue>��</font>"
            Else
                Response.Write "<font color=red>��</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrKind_GuestBook(4), rsGuestKind("KindID"), ",") = True Then
                Response.Write "<font color=blue>��</font>"
            Else
                Response.Write "<font color=red>��</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrKind_GuestBook(5), rsGuestKind("KindID"), ",") = True Then
                Response.Write "<font color=blue>��</font>"
            Else
                Response.Write "<font color=red>��</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrKind_GuestBook(6), rsGuestKind("KindID"), ",") = True Then
                Response.Write "<font color=blue>��</font>"
            Else
                Response.Write "<font color=red>��</font>"
            End If
            Response.Write "</td>"
            Response.Write "  </tr>"
            rsGuestKind.MoveNext
        Loop
        Set rsGuestKind = Nothing
        Response.Write "</table>"
    End If
End Sub

Sub ShowChannelPurview()
    Dim rsAdmin, Channel_Purview, arrClass_View, arrClass_Input, arrClass_Check, arrClass_Manage
    Dim rsChannel, ChannelDir, ChannelName, ChannelShortName, ModuleType

    If ChannelID > 0 Then
        Set rsChannel = Conn.Execute("select * from PE_Channel where ChannelID=" & ChannelID)
        If Not (rsChannel.BOF Or rsChannel.EOF) Then
            ChannelDir = rsChannel("ChannelDir")
            ChannelName = rsChannel("ChannelName")
            ChannelShortName = rsChannel("ChannelShortName")
            ModuleType = rsChannel("ModuleType")
            Set rsAdmin = Conn.Execute("select AdminPurview_" & ChannelDir & ",arrClass_View,arrClass_Input,arrClass_Check,arrClass_Manage from PE_Admin where AdminName='" & AdminName & "'")
            If Not (rsAdmin.BOF And rsAdmin.EOF) Then
                Channel_Purview = rsAdmin(0)
                arrClass_View = rsAdmin("arrClass_View")
                arrClass_Input = rsAdmin("arrClass_Input")
                arrClass_Check = rsAdmin("arrClass_Check")
                arrClass_Manage = rsAdmin("arrClass_Manage")
            End If
            rsAdmin.Close
            Set rsAdmin = Nothing
        End If
        rsChannel.Close
        Set rsChannel = Nothing
    End If

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td colspan='2' height='22'><strong>" & ChannelName & "</strong> "
    If Channel_Purview = 1 Then Response.Write "��Ƶ������Ա��"
    If Channel_Purview = 2 Then Response.Write "����Ŀ�ܱࣩ"
    If Channel_Purview = 3 Then Response.Write "����Ŀ����Ա��"
    If Channel_Purview = 4 Then Response.Write "����Ȩ�ޣ�"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>��Ŀ����ר��������ɹ���Ȩ��</td>"
    Response.Write "    <td align='center' width='20%'>"
    If AdminPurview = 1 Or Channel_Purview = 1 Then
        Response.Write "<font color=blue>��</font>"
    Else
        Response.Write "<font color=red>��</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>ר��" & ChannelShortName & "����Ȩ��</td>"
    Response.Write "    <td align='center' width='20%'>"
    If AdminPurview = 1 Or Channel_Purview <= 2 Then
        Response.Write "<font color=blue>��</font>"
    Else
        Response.Write "<font color=red>��</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>" & ChannelShortName & "���ۡ�����վ����������Ȩ��</td>"
    Response.Write "    <td align='center' width='20%'>"
    If AdminPurview = 1 Or Channel_Purview = 1 Then
        Response.Write "<font color=blue>��</font>"
    Else
        Response.Write "<font color=red>��</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>����Ŀ" & ChannelShortName & "¼�롢��ˡ�����Ȩ��</td>"
    Response.Write "    <td align='center' width='20%'>"
    If Channel_Purview <= 2 Then
        Response.Write "<font color=blue>ȫ��Ȩ��</font>"
    ElseIf Channel_Purview = 3 Then
        Response.Write "<font color=blue>����Ȩ��</font>"
    Else
        Response.Write "<font color=red>��Ȩ��</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22'><strong>����Ȩ�ޣ�</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td>"
    Response.Write "ģ�����&nbsp;" & ShowPurview("Template_" & ChannelDir) & "&nbsp;&nbsp;"
    Response.Write "JS�ļ�����&nbsp;" & ShowPurview("JsFile_" & ChannelDir) & "&nbsp;&nbsp;"
    Response.Write "�����˵�&nbsp;" & ShowPurview("Keyword_" & ChannelDir) & "&nbsp;&nbsp;"
    Response.Write "�ؼ��ֹ���&nbsp;" & ShowPurview("Template_" & ChannelDir) & "&nbsp;&nbsp;"
    If ModuleType = 5 Then
        Response.Write "���̹���&nbsp;" & ShowPurview("Producer_" & ChannelDir) & "&nbsp;&nbsp;"
        Response.Write "Ʒ�ƹ���&nbsp;" & ShowPurview("Trademark_" & ChannelDir) & "&nbsp;&nbsp;"
    Else
        Response.Write "���߹���&nbsp;" & ShowPurview("Author_" & ChannelDir) & "&nbsp;&nbsp;"
        Response.Write "��Դ����&nbsp;" & ShowPurview("Copyfrom_" & ChannelDir) & "&nbsp;&nbsp;"
    End If
    Response.Write "����XML&nbsp;" & ShowPurview("XML_" & ChannelDir) & "&nbsp;&nbsp;"
    Response.Write "�Զ����ֶ�&nbsp;" & ShowPurview("Field_" & ChannelDir) & "&nbsp;&nbsp;"
    Response.Write "������&nbsp;" & ShowPurview("AD_" & ChannelDir) & "&nbsp;&nbsp;"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

    If Channel_Purview = 3 Then
        Dim arrShowLine(20)
        Dim sqlClass, rsClass, i, iDepth
        For i = 0 To UBound(arrShowLine)
            arrShowLine(i) = False
        Next
        sqlClass = "select * from PE_Class where ChannelID=" & ChannelID & " order by RootID,OrderID"
        Set rsClass = Server.CreateObject("adodb.recordset")
        rsClass.Open sqlClass, Conn, 1, 1

        Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
        Response.Write "  <tr align='center' class='title'>"
        Response.Write "    <td height='22'><strong>��Ŀ����</strong></td>"
        Response.Write "    <td width='100'><strong>�鿴</strong></td>"
        Response.Write "    <td width='100'><strong>¼��</strong></td>"
        Response.Write "    <td width='100'><strong>���</strong></td>"
        Response.Write "    <td width='100'><strong>����</strong></td>"
        Response.Write "  </tr>"
        Do While Not rsClass.EOF
            Response.Write "     <tr class='tdbg'><td>"
            iDepth = rsClass("Depth")
            If rsClass("NextID") > 0 Then
                arrShowLine(iDepth) = True
            Else
                arrShowLine(iDepth) = False
            End If
            If iDepth > 0 Then
                For i = 1 To iDepth
                    If i = iDepth Then
                        If rsClass("NextID") > 0 Then
                            Response.Write "<img src='../images/tree_line1.gif' width='17' height='16' valign='abvmiddle'>"
                        Else
                            Response.Write "<img src='../images/tree_line2.gif' width='17' height='16' valign='abvmiddle'>"
                        End If
                    Else
                        If arrShowLine(i) = True Then
                            Response.Write "<img src='../images/tree_line3.gif' width='17' height='16' valign='abvmiddle'>"
                        Else
                            Response.Write "<img src='../images/tree_line4.gif' width='17' height='16' valign='abvmiddle'>"
                        End If
                    End If
                Next
            End If
            If rsClass("Child") > 0 Then
                Response.Write "<img src='../images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>"
            Else
                Response.Write "<img src='../images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>"
            End If
            If rsClass("Depth") = 0 Then
                Response.Write "<b>"
            End If
            Response.Write rsClass("ClassName")
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrClass_View, rsClass("ClassID"), ",") = True Then
                Response.Write "<font color=blue>��</font>"
            Else
                Response.Write "<font color=red>��</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrClass_Input, rsClass("ClassID"), ",") = True Then
                Response.Write "<font color=blue>��</font>"
            Else
                Response.Write "<font color=red>��</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrClass_Check, rsClass("ClassID"), ",") = True Then
                Response.Write "<font color=blue>��</font>"
            Else
                Response.Write "<font color=red>��</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrClass_Manage, rsClass("ClassID"), ",") = True Then
                Response.Write "<font color=blue>��</font>"
            Else
                Response.Write "<font color=red>��</font>"
            End If
            Response.Write "</td></tr>"
        rsClass.MoveNext
        Loop
        rsClass.Close
        Set rsClass = Nothing
        Response.Write "</table>"
    End If
End Sub

Function ShowPurview(strPurview)
    If CheckPurview_Other(AdminPurview_Others, strPurview) = True Then
        ShowPurview = "<font color=blue>��</font>"
    Else
        ShowPurview = "<font color=red>��</font>"
    End If
End Function

Function ShowChannelOtherPurview(Channel_Purview, strPurview)
    If ChannelPurview = 1 And CheckPurview_Other(AdminPurview_Others, strPurview) = True Then
        ShowChannelOtherPurview = "<font color=blue>��</font>"
    Else
        ShowChannelOtherPurview = "<font color=red>��</font>"
    End If
End Function

Function GetChannelList()
    Dim rsChannel, sqlChannel, strChannel, i
    If ChannelID = 0 Then
        strChannel = "<a href='" & strFileName & "?iChannelID=0'><font color=red>���й���Ȩ��</font></a> | "
    Else
        strChannel = "<a href='" & strFileName & "?iChannelID=0'>���й���Ȩ��</a> | "
    End If
    sqlChannel = "select * from PE_Channel where ChannelType<=1 and Disabled=" & PE_False & " order by OrderID"
    Set rsChannel = Server.CreateObject("adodb.recordset")
    rsChannel.Open sqlChannel, Conn, 1, 1
    If rsChannel.BOF And rsChannel.EOF Then
        strChannel = strChannel & "û���κ�Ƶ��"
    Else
        i = 1
        Do While Not rsChannel.EOF
            If rsChannel("ChannelID") = ChannelID Then
                strChannel = strChannel & "<a href='" & strFileName & "?iChannelID=" & ChannelID & "'><font color=red>" & rsChannel("ChannelName") & "Ȩ��</font></a>"
            Else
                strChannel = strChannel & "<a href='" & strFileName & "?iChannelID=" & rsChannel("ChannelID") & "'>" & rsChannel("ChannelName") & "Ȩ��</a>"
            End If
            strChannel = strChannel & " | "
            i = i + 1
            If i Mod 10 = 0 Then
                strChannel = strChannel & "<br>"
            End If
            rsChannel.MoveNext
        Loop
    End If
    rsChannel.Close
    Set rsChannel = Nothing
    GetChannelList = strChannel
End Function

Function GetManagePath()
    Dim strPath, sqlPath, rsPath
    strPath = "�����ڵ�λ�ã��鿴����Ȩ��&nbsp;&gt;&gt;&nbsp;"
    If ChannelID = 0 Then
        strPath = strPath & "���й���Ȩ��"
    Else
        sqlPath = "select ChannelID,ChannelName from PE_Channel where ChannelID=" & ChannelID
        Set rsPath = Server.CreateObject("adodb.recordset")
        rsPath.Open sqlPath, Conn, 1, 1
        If rsPath.BOF And rsPath.EOF Then
            strPath = strPath & "�����Ƶ������"
        Else
            strPath = strPath & "<a href='" & strFileName & "?iChannelID=" & rsPath(0) & "'>" & rsPath(1) & "Ȩ��</a>"
        End If
        rsPath.Close
        Set rsPath = Nothing
    End If
    GetManagePath = strPath
End Function
%>
