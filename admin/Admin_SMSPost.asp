<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.MD5_New.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

Dim SendTo
SendTo = Trim(Request("SendTo"))
If AdminPurview > 1 Then
    Select Case SendTo
    Case "Member"
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "SendSMSToMember")
    Case "Contacter"
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "SendSMSToContacter")
    Case "Consignee"
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "SendSMSToConsignee")
    Case "Other"
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "SendSMSToOther")
    End Select
    If PurviewPassed = False Then
        Response.write "<br><p align='center'><font color='red' style='font-size:9pt'>�Բ�����û�д��������Ȩ�ޡ�</font></p>"
        Response.End
    End If
End If
Dim arrMobile()
regEx.Pattern = "[^0-9]"

Dim mSendNum     '���ն��ŵ��ֻ���
Dim mContent     '��������
Dim mSendTiming  '�Ƿ�ʱ���ͣ�0Ϊ��Ĭ�ϣ���1Ϊ��ʱ
Dim mSendTime    '��ʱ����ʱ��
Dim MD5String   'MD5У���ַ���MD5�ַ��������������ֶμ����û����룬Ȼ����MD5�������ܵõ����ַ������ֶ�˳�����£�
                 'SMSUserName & SMSKey & mSendNum & mContent & mSendTiming & mSendTime
Dim RecieverCount
RecieverCount = 0
mContent = Trim(Request("Content"))
mSendTiming = PE_CLng(Trim(Request("SendTiming")))
Select Case SendTo
Case "SendToMember"
    mSendNum = GetReciever_Member()
Case "SendToContacter"
    mSendNum = GetReciever_Contacter()
Case "SendToConsignee"
    mSendNum = GetReciever_Consignee()
Case "SendToOther"
    mSendNum = Trim(Request("Receiver"))
    Dim arrReceiver
    arrReceiver = Split(mSendNum, vbCrLf)
    RecieverCount = UBound(arrReceiver) + 1
End Select
If mSendNum = "" Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>�Ҳ��������������ֻ�����</li>"
End If
mSendTime = Trim(Request("SendDate")) & " " & Trim(Request("SendTime_Hour")) & ":" & Trim(Request("SendTime_Minute")) & ":00"

If mContent = "" Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>�������������</li>"
End If
If mSendTiming = 1 And IsDate(mSendTime) = False Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>��ʱ����ʱ��ĸ�ʽ���ԣ�</li>"
End If

If FoundErr = True Then
    Response.write ErrMsg
    Response.End
End If

Dim ID, ranNum, dtNow
Randomize
dtNow = Now()
ranNum = Int(900 * Rnd) + 100
ID = Year(dtNow) & Right("0" & Month(dtNow), 2) & Right("0" & Day(dtNow), 2) & Right("0" & Hour(dtNow), 2) & Right("0" & Minute(dtNow), 2) & Right("0" & Second(dtNow), 2) & ranNum

Dim PE_MD5
Set PE_MD5 = New Md5_Class
MD5String = UCase(Trim(PE_MD5.MD5(ID & SMSUserName & SMSKey & mSendNum & mContent & mSendTiming & mSendTime)))
Set PE_MD5 = Nothing

Dim MessageCount
MessageCount = ((Len(mContent) \ 70) + 1) * RecieverCount
%>
<html>
<head>
<Title>����Ԥ��</Title>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312' />
<link href='Admin_Style.css' rel='stylesheet' type='text/css'>
</head>
<body>
<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>
  <tr class='topbg'>
    <td height='22' colspan='10'><table width='100%'><tr class='topbg'><td align='center'><b>�� �� �� �� �� ��</b></td><td width='60' align='right'><a href='http://go.powereasy.net/go.aspx?UrlID=10047' target='_blank'><img src='images/help.gif' border='0'></a></td></tr></table></td>
  </tr>
  <tr class='tdbg'>
    <td width='70' height='30'><strong>��������</strong></td>
    <td><a href='Admin_SMS.asp?SendTo=Member'>����Ա���Ͷ���</a>&nbsp;|&nbsp;    <a href='Admin_SMS.asp?SendTo=Contacter'>����ϵ�˷��Ͷ���</a>&nbsp;|&nbsp;    <a href='Admin_SMS.asp?SendTo=Consignee'>�������е��ջ��˷��Ͷ���</a>&nbsp;|&nbsp;    <a href='Admin_SMS.asp?SendTo=Other'>�������˷��Ͷ���</a>    </td>
  </tr>
</table>
<form name='message' method='post' action='http://sms.powereasy.net/MessageGate2/MessageGate.aspx'>
<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>
<tr class='title'><td colspan='2' align='center'>Ԥ �� �� ��</td></tr>
<tr class='tdbg' valign='top'>
<td width='300'>�����Ǹ�����ָ�����������ҵ��Ľ����ˣ�<br><textarea name="SendNum" rows='20' cols='40' readonly><%= mSendNum %></textarea></td>
<td><b>�������ݣ�</b><br><textarea name="Content" rows='5' cols='60' readonly><%= mContent %></textarea><br><br><br><br><br><br><b>����ͳ�ƣ�</b><br>��Ҫ�� <%=RecieverCount%> �����뷢�� <%=MessageCount%> ������<br><br><b>˵����</b><br>��Ϊÿ�����Ų��ܳ���70���֣����Զ��������ܻ���ں�������<br>��Ϊ���������еı����滻��ԭ�򣬿��ܻᵼ��ʵ�ʷ��͵Ķ������ᳬ���������Ķ����������ս���Զ��׶���ͨƽ̨�ϵ�ʵ�ʷ�����ĿΪ׼��</td></tr>
<tr class='tdbg'><td colspan='2' height='50' align='center'><input type='submit' name='submit' value='�ύ�����ŷ�����'>
<input type="hidden" name="UserName" value="<%= SMSUserName %>" />
<input type="hidden" name="SendTiming" value="<%= mSendTiming %>" />
<input type="hidden" name="SendTime" value="<%= mSendTime %>" />
<input type="hidden" name="MD5String" value="<%= MD5String %>" />
<input type="hidden" name="Reserve" value="����ԱID��<%= AdminID %>" />
<input type="hidden" name="ID" value="<%= ID %>" />
</td></tr></table>
</form>
</body>
</html>
<%
Function FoundMobile(sMobile, iTemp)
    Dim i, bl
    bl = False
    For i = iTemp To 0 Step -1
        If Trim(arrMobile(i)) = Trim(sMobile) Then
            bl = True
            Exit For
        End If
    Next
    FoundMobile = bl
End Function

Function GetValidNumber(sNumber)
    Dim strTemp, l
    strTemp = regEx.Replace(sNumber, "")
    l = Len(strTemp)
    If (l = 11 Or l = 12) And (Left(strTemp, 1) = "0" Or Left(strTemp, 1) = "1") Then
        GetValidNumber = strTemp
    Else
        GetValidNumber = ""
    End If
End Function

Function GetReciever_Member()
    Dim InceptType, GroupID, inceptUser
    Dim sqlUser, strReciever, strMobile
    Dim BeginID, EndID
    inceptUser = Replace(ReplaceBadChar(Request("InceptUser")), ",", "','")
    BeginID = PE_CLng(Trim(Request("BeginID")))
    EndID = PE_CLng(Trim(Request("EndID")))
    strReciever = ""
    sqlUser = "select U.UserName,C.TrueName,C.Mobile,C.PHS from PE_User U left join PE_Contacter C on U.ContacterID=C.ContacterID where (C.Mobile<>'' or C.PHS<>'')"
    InceptType = PE_CLng(Trim(Request("InceptType")))
    GroupID = Trim(Request("GroupID"))
    If IsValidID(GroupID) = False Then
        GroupID = ""
    End If

    If GroupID <> "" Then
        sqlUser = sqlUser & " and U.GroupID in (" & GroupID & ")"
    End If

    If inceptUser <> "" Then
        sqlUser = sqlUser & " and U.UserName in ('" & inceptUser & "')"
    End If
    If BeginID > 0 Then
        sqlUser = sqlUser & " And U.UserID>=" & BeginID
    End If
    If EndID > 0 Then
        sqlUser = sqlUser & " And U.UserID<=" & EndID
    End If
    sqlUser = sqlUser & " order by C.Mobile,C.PHS"
    Dim rsUser, strMoblie
    Set rsUser = server.CreateObject("adodb.recordset")
    rsUser.open sqlUser, Conn, 1, 1
    ReDim arrMobile(rsUser.recordcount)
    Do While Not rsUser.EOF
        strMobile = rsUser("Mobile")
        If strMobile = "" Then strMobile = rsUser("PHS")
        strMobile = GetValidNumber(strMobile)
        If strMobile <> "" Then
            If FoundMobile(strMobile, RecieverCount) = False Then
                arrMobile(RecieverCount) = strMobile
                If strReciever <> "" Then
                    strReciever = strReciever & vbCrLf
                End If
                strReciever = strReciever & strMobile & "," & rsUser("UserName")
                If rsUser("TrueName") <> "" Then
                    strReciever = strReciever & "," & rsUser("TrueName")
                End If
                RecieverCount = RecieverCount + 1
            End If
        End If
        rsUser.movenext
    Loop
    rsUser.Close
    Set rsUser = Nothing
    GetReciever_Member = strReciever
End Function

Function GetReciever_Contacter()
    Dim InceptType, GroupID, TrueName
    Dim sqlContacter, strReciever, strMobile
    strReciever = ""
    sqlContacter = "select TrueName,Mobile,PHS from PE_Contacter where (Mobile<>'' or PHS<>'')"
    InceptType = PE_CLng(Trim(Request("InceptType")))
    Select Case InceptType
    Case 0  '���л�Ա
        
    Case 1  'ָ������
        sqlContacter = sqlContacter & " and Country='" & ReplaceBadChar(Request("Country")) & "'"
        sqlContacter = sqlContacter & " and Province='" & ReplaceBadChar(Request("Province")) & "'"
        sqlContacter = sqlContacter & " and City='" & ReplaceBadChar(Request("City")) & "'"
    Case 2  'ָ����ϵ��
        TrueName = Replace(ReplaceBadChar(Request("TrueName")), ",", "','")
        sqlContacter = sqlContacter & " and TrueName in ('" & TrueName & "')"
    End Select
    sqlContacter = sqlContacter & " order by Mobile,PHS"
    Dim rsContacter, strMoblie
    Set rsContacter = server.CreateObject("adodb.recordset")
    rsContacter.open sqlContacter, Conn, 1, 1
    ReDim arrMobile(rsContacter.recordcount)
    Do While Not rsContacter.EOF
        strMobile = rsContacter("Mobile")
        If strMobile = "" Then strMobile = rsContacter("PHS")
        strMobile = GetValidNumber(strMobile)
        If strMobile <> "" Then
            If FoundMobile(strMobile, RecieverCount) = False Then
                arrMobile(RecieverCount) = strMobile
                If strReciever <> "" Then
                    strReciever = strReciever & vbCrLf
                End If
                strReciever = strReciever & strMobile & "," & rsContacter("TrueName")
                RecieverCount = RecieverCount + 1
            End If
        End If
        rsContacter.movenext
    Loop
    rsContacter.Close
    Set rsContacter = Nothing
    GetReciever_Contacter = strReciever
End Function

Function GetReciever_Consignee()
    Dim InceptType, GroupID, OrderFormID
    Dim sqlOrder, strReciever, strMobile
    strReciever = ""
    sqlOrder = "select OrderFormNum,ContacterName,Mobile,Phone from PE_OrderForm where (Mobile<>'' or Phone<>'')"
    InceptType = PE_CLng(Trim(Request("InceptType")))
    Select Case InceptType
    Case 0  '���ж���
        
    Case 1  '��ѯ����
        Dim BeginID, EndID, BeginDate, EndDate, MinMoney, MaxMoney, OrderStatus, PayStatus, DeliverStatus, OrderFormNum, ClientName, UserName, AgentName, ContacterName, Address, Phone, Mobile, Remark, ProductName
        BeginID = PE_CLng(Trim(Request("BeginID")))
        EndID = PE_CLng(Trim(Request("EndID")))
        BeginDate = Trim(Request("BeginDate"))
        EndDate = Trim(Request("EndDate"))
        MinMoney = PE_CDbl(Trim(Request("MinMoney")))
        MaxMoney = PE_CDbl(Trim(Request("MaxMoney")))
        OrderStatus = PE_CLng(Trim(Request("OrderStatus")))
        PayStatus = PE_CLng(Trim(Request("PayStatus")))
        DeliverStatus = PE_CLng(Trim(Request("DeliverStatus")))
        OrderFormNum = ReplaceBadChar(Trim(Request("OrderFormNum")))
        ClientName = ReplaceBadChar(Trim(Request("ClientName")))
        UserName = ReplaceBadChar(Trim(Request("UserName")))
        AgentName = ReplaceBadChar(Trim(Request("AgentName")))
        ContacterName = ReplaceBadChar(Trim(Request("ContacterName")))
        Address = ReplaceBadChar(Trim(Request("Address")))
        Phone = ReplaceBadChar(Trim(Request("Phone")))
        Mobile = ReplaceBadChar(Trim(Request("Mobile")))
        Remark = ReplaceBadChar(Trim(Request("Remark")))
        ProductName = ReplaceBadChar(Trim(Request("ProductName")))

        If BeginID > 0 Then
            sqlOrder = sqlOrder & " And OrderFormID>=" & BeginID
        End If
        If EndID > 0 Then
            sqlOrder = sqlOrder & " And OrderFormID<=" & EndID
        End If

        If BeginDate <> "" Then
            BeginDate = PE_CDate(BeginDate)
            If SystemDatabaseType = "SQL" Then
                sqlOrder = sqlOrder & " And InputTime>='" & BeginDate & "'"
            Else
                sqlOrder = sqlOrder & " And InputTime>=#" & BeginDate & "#"
            End If
        End If
        If EndDate <> "" Then
            EndDate = PE_CDate(EndDate)
            If SystemDatabaseType = "SQL" Then
                sqlOrder = sqlOrder & " And InputTime<='" & EndDate & "'"
            Else
                sqlOrder = sqlOrder & " And InputTime<=#" & EndDate & "#"
            End If
        End If
        If MinMoney > 0 Then
            sqlOrder = sqlOrder & " And MoneyTotal>=" & MinMoney
        End If
        If MaxMoney > 0 Then
            sqlOrder = sqlOrder & " And MoneyTotal<=" & MaxMoney
        End If
        If OrderStatus >= 0 Then
            sqlOrder = sqlOrder & " And Status=" & OrderStatus
        End If
        If PayStatus >= 0 Then
            Select Case PayStatus
            Case 0
                sqlOrder = sqlOrder & " And MoneyTotal>0 And MoneyReceipt=0"
            Case 1
                sqlOrder = sqlOrder & " And MoneyTotal>MoneyReceipt And MoneyReceipt>0"
            Case 2
                sqlOrder = sqlOrder & " And MoneyTotal<=MoneyReceipt"
            End Select
        End If
        If DeliverStatus >= 0 Then
            sqlOrder = sqlOrder & " And DeliverStatus=" & DeliverStatus
        End If
        If OrderFormNum <> "" Then
            sqlOrder = sqlOrder & " And OrderFormNum like '%" & OrderFormNum & "%'"
        End If
        If ClientName <> "" Then
            sqlOrder = "select O.OrderFormNum,O.ContacterName,O.Mobile,O.Phone from PE_OrderForm O left join PE_Client C on O.ClientID=C.ClientID where (O.Mobile<>'' or O.Phone<>'')"
            sqlOrder = sqlOrder & " And C.ClientName like '%" & ClientName & "%'"
        End If
        If UserName <> "" Then
            sqlOrder = sqlOrder & " And UserName like '%" & UserName & "%'"
        End If
        If AgentName <> "" Then
            sqlOrder = sqlOrder & " And AgentName='" & AgentName & "'"
        End If
        If ContacterName <> "" Then
            sqlOrder = sqlOrder & " And ContacterName like '%" & ContacterName & "%'"
        End If
        If Address <> "" Then
            sqlOrder = sqlOrder & " And Address like '%" & Address & "%'"
        End If
        If Phone <> "" Then
            sqlOrder = sqlOrder & " And Phone like '%" & Phone & "%'"
        End If
        If Mobile <> "" Then
            sqlOrder = sqlOrder & " And Mobile like '%" & Mobile & "%'"
        End If
        If Remark <> "" Then
            sqlOrder = sqlOrder & " And Remark like '%" & Remark & "%'"
        End If
        If ProductName <> "" Then
            sqlOrder = "select O.OrderFormNum,O.ContacterName,O.Mobile,O.Phone from PE_Product P inner join (Pe_OrderFormItem I inner join (PE_OrderForm O left join PE_Client C On O.ClientID = C.ClientID) on I.OrderFormID = O.OrderFormID) on P.ProductID = I.ProductID where (O.Mobile<>'' or O.Phone<>'')"
            sqlOrder = sqlOrder & " And P.ProductName like '%" & ProductName & "%'"
        End If

    Case 2  'ָ������ID
        OrderFormID = Trim(Request("OrderFormID"))
        If IsValidID(OrderFormID) = False Then
            OrderFormID = ""
        End If

        If OrderFormID = "" Then OrderFormID = "0"
        sqlOrder = sqlOrder & " and OrderFormID in (" & OrderFormID & ")"
    End Select
    sqlOrder = sqlOrder & " order by Mobile,Phone"
    Dim rsOrder, strMoblie
    Set rsOrder = server.CreateObject("adodb.recordset")
    rsOrder.open sqlOrder, Conn, 1, 1
    ReDim arrMobile(rsOrder.recordcount)
    Do While Not rsOrder.EOF
        strMobile = rsOrder("Mobile")
        If strMobile = "" Then strMobile = rsOrder("Phone")
        strMobile = GetValidNumber(strMobile)
        If strMobile <> "" Then
            If FoundMobile(strMobile, RecieverCount) = False Then
                arrMobile(RecieverCount) = strMobile
                If strReciever <> "" Then
                    strReciever = strReciever & vbCrLf
                End If
                strReciever = strReciever & strMobile & "," & rsOrder("ContacterName") & "," & rsOrder("OrderFormNum")
                RecieverCount = RecieverCount + 1
            End If
        End If
        rsOrder.movenext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
    GetReciever_Consignee = strReciever
End Function
%>
