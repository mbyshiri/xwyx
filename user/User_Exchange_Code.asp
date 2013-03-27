<!--#include file="CommonCode.asp"-->
<!--#include file="../Include/PowerEasy.Bankroll.asp"-->
<!--#include file="../Include/PowerEasy.ConsumeLog.asp"-->
<!--#include file="../Include/PowerEasy.RechargeLog.asp"-->
<!--#include file="../Include/PowerEasy.Base64.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Sub Execute()
    Select Case Action
    Case "Exchange"
        Call Exchange
    Case "SaveExchange"
        Call SaveExchange
    Case "Valid"
        Call Valid
    Case "SaveValid"
        Call SaveValid
    Case "Recharge"
        Call Recharge
    Case "SaveRecharge"
        Call SaveRecharge
    Case "GetCard"
        Call GetCard
    Case "SendPoint"
        Call SendPoint
    Case "SaveSendPoint"
        Call SaveSendPoint
    Case Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ĳ���</li>"
    End Select
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
End Sub

Sub ShowUserInfo()
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>�û�����</td>"
    Response.Write "      <td>" & UserName & "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>�ʽ���</td>"
    Response.Write "      <td>" & Balance & " Ԫ</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>���û��֣�</td>"
    Response.Write "      <td>" & UserExp & " ��</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>����" & PointName & "����</td>"
    Response.Write "      <td>" & UserPoint & " " & PointUnit & "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>��Ч������</td>"
    Response.Write "      <td>��ʼ�������ڣ�" & FormatDateTime(BeginTime, 2) & "&nbsp;&nbsp;&nbsp;&nbsp;��Ч�ڣ�"
    If ValidNum = -1 Then
        Response.Write "������<br>"
    Else
        Response.Write ValidNum & arrCardUnit(ValidUnit) & "<br>"
        If ValidDays >= 0 Then
            Response.Write "���� <font color=blue>" & ValidDays & "</font> �쵽��"
        Else
            Response.Write "�Ѿ����� <font color=red>" & Abs(ValidDays) & "</font> ��"
        End If
    End If
    Response.Write "      </td>"
    Response.Write "    </tr>"
End Sub

Sub Exchange()
    If UserSetting(18) = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��������������һ�" & PointName & "��</li>"
        Exit Sub
    End If
    Response.Write "<form name='myform' action='User_Exchange.asp' method='post'>"
    Response.Write "  <table width='100%' border='0' cellspacing='1' cellpadding='2' align='center' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height=22 colSpan=2 align='center'><b>�һ�" & PointName & "</b></td>"
    Response.Write "    </tr>"
    Call ShowUserInfo
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>�һ�" & PointName & "��</td>"
    Response.Write "      <td>"
    Response.Write "        <input type='radio' name='ChangeType' value='1' checked>ʹ���ʽ���"
    Response.Write "        �� <input name='ChangeMoney' type='text' value='10' size='6' maxlength='8' style='text-align:center'> Ԫ�һ���" & PointName
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;�һ����ʣ�" & FormatNumber(MoneyExchangePoint, 2, vbTrue, vbFalse, vbTrue) & "Ԫ:1" & PointUnit
    Response.Write "        <br>"
    Response.Write "        <input type='radio' name='ChangeType' value='2'>ʹ�þ�����֣�"
    Response.Write "        �� <input name='ChangeExp' type='text' value='10' size='6' maxlength='8' style='text-align:center'> �ֶһ���" & PointName
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;�һ����ʣ�" & FormatNumber(UserExpExchangePoint, 2, vbTrue, vbFalse, vbTrue) & "��:1" & PointUnit
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveExchange'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value='ִ�жһ�'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub Valid()
    If UserSetting(19) = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��������������һ���Ч�ڣ�</li>"
        Exit Sub
    End If
    Response.Write "<form name='myform' action='User_Exchange.asp' method='post'>"
    Response.Write "  <table width='100%' border='0' cellspacing='1' cellpadding='2' align='center' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height=22 colSpan=2 align='center'><b>�� �� �� Ч ��</b></td>"
    Response.Write "    </tr>"
    Call ShowUserInfo
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>�һ���Ч�ڣ�</td>"
    Response.Write "      <td>"
    Response.Write "        <input type='radio' name='ChangeType' value='1' checked>ʹ���ʽ���"
    Response.Write "        �� <input name='ChangeMoney' type='text' value='10' size='6' maxlength='8' style='text-align:center'> Ԫ�һ�����Ч��"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;�һ����ʣ�" & MoneyExchangeValidDay & "Ԫ:1��"
    Response.Write "        <br>"
    Response.Write "        <input type='radio' name='ChangeType' value='2'>ʹ�þ�����֣�"
    Response.Write "        �� <input name='ChangeExp' type='text' value='10' size='6' maxlength='8' style='text-align:center'> �ֶһ�����Ч��"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;�һ����ʣ�" & UserExpExchangeValidDay & "��:1��"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveValid'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value='ִ�жһ�'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub Recharge()
    Response.Write "<form name='myform' action='User_Exchange.asp' method='post'>"
    Response.Write "  <table width='500' border='0' cellspacing='1' cellpadding='2' align='center' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height=22 colSpan=2 align='center'><b>�� ֵ �� �� ֵ</b></td>"
    Response.Write "    </tr>"
    Call ShowUserInfo
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>��ֵ�����ţ�</td>"
    Response.Write "      <td><input name='CardNum' type='text' value='' size='30' maxlength='30'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>��ֵ�����룺</td>"
    Response.Write "      <td><input name='Password' type='text' value='' size='30' maxlength='30'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'><input name='Action' type='hidden' id='Action' value='SaveRecharge'>"
    Response.Write "        <input name=Submit   type=submit id='Submit' value=' ȷ �� '></td>"
    Response.Write "    </tr>"
    Response.Write "  </TABLE>"
    Response.Write "</form>"
End Sub

Sub SendPoint()
    If UserSetting(20) = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������" & PointName & "���͸����ˣ�</li>"
        Exit Sub
    End If
    Response.Write "<form name='myform' action='User_Exchange.asp' method='post'>" & vbCrLf
    Response.Write "  <table width='500' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr height='22' align='center' class='title'>" & vbCrLf
    Response.Write "      <td colSpan='2'><b>����" & PointName & "</b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='120' align='right' class='tdbg5'>�� �� ����</td>" & vbCrLf
    Response.Write "      <td>" & UserName & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='120' align='right' class='tdbg5'>��ǰ" & PointName & "����</td>" & vbCrLf
    Response.Write "      <td>" & UserPoint & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='120' align='right' class='tdbg5'>�����˵��û�����</td>" & vbCrLf
    Response.Write "      <td> <input name='SendObject' type='text' size='30'> </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='120' align='right' class='tdbg5'>���͵�" & PointName & "����</td>" & vbCrLf
    Response.Write "      <td> <input name='SendPoint' type='text' maxLength='16' size='30'> </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <td height='40' colspan='2'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveSendPoint'>" & vbCrLf
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' ���� '>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

Sub GetCard()
    Response.Write "<br><table width='100%' cellspacing='1' cellpadding='2'  class='border'><tr class='title'><td align='center'>��ȡ�����ֵ��</td></tr>"
    Response.Write "<tr><td height='100'>"
    Dim rsOrderItem, rsCard, sqlCard, i, strCardInfo
    Set rsOrderItem = Conn.Execute("select O.OrderFormID,I.ItemID,P.ProductID,P.ProductName,P.ProductKind,I.Amount from PE_OrderForm O inner join (PE_OrderFormItem I inner join PE_Product P on I.ProductID=P.ProductID) on I.OrderFormID=O.OrderFormID where O.UserName='" & UserName & "' and P.ProductKind=3 order by I.ItemID")
    If rsOrderItem.BOF And rsOrderItem.EOF Then
        Response.Write "����û�й����κε㿨����Ʒ��"
    Else
        Response.Write "<br><br><table width='80%' align='center' cellspacing='1' cellpadding='2'>"
        Response.Write "<tr class='title' align='center'><td>��Ʒ����</td><td>��ֵ������</td><td>��ֵ������</td><td>��ֵ������</td><td>��ֵ����ֵ</td><td>��ֵ������</td><td>��ֵ��ֹ����</td></tr>"
        Do While Not rsOrderItem.EOF
            Set rsCard = Conn.Execute("select * from PE_Card where ProductID=" & rsOrderItem("ProductID") & " and OrderFormItemID=" & rsOrderItem("ItemID") & "")
            If rsCard.BOF And rsCard.EOF Then
                Response.Write "<tr class='tdbg' align='center'><td>" & rsOrderItem("Productname") & "</td><td colspan='10' align='center'>��û�н������ź����룬������������ϵ��</td></tr>"
            Else
                i = 0
                Do While Not rsCard.EOF
                    If rsCard("UserName") = "" Then
                        Response.Write "<tr class='tdbg' align='center'><td>" & rsOrderItem("Productname") & "</td>"
                        Response.Write "<td>"
                        If rsCard("CardType") = 0 Then
                            Response.Write "��վ��ֵ��"
                        Else
                            Response.Write "<font color='blue'>������˾��</font>"
                        End If
                        Response.Write "</td>"
                        Response.Write "<td>" & rsCard("CardNum") & "</td>"
                        Response.Write "<td>" & Base64decode(rsCard("Password")) & "</td>"
                        Response.Write "<td>" & rsCard("Money") & "</td>"
                        Response.Write "<td>" & GetValidNum(rsCard("ValidNum"), rsCard("ValidUnit")) & arrCardUnit(rsCard("ValidUnit")) & "</td>"
                        Response.Write "<td>" & rsCard("EndDate") & "</td></tr>"
                        i = i + 1
                    End If
                    rsCard.MoveNext
                Loop
                If i = 0 Then
                    Response.Write "<tr class='tdbg' align='center'><td>" & rsOrderItem("Productname") & "</td><td colspan='10' height='50' align='center'>����������г�ֵ�����Ѿ�ʹ�á�</td></tr>"
                End If
            End If
            Set rsCard = Nothing
            rsOrderItem.MoveNext
        Loop
        Response.Write "</table><br><br>"
    End If
    Set rsOrderItem = Nothing
    Response.Write "</td></tr>"
    Response.Write "<tr class='tdbg'><td><font color='red'>ע�⣺</font><br>����ֻ��ʾ�˻�δʹ�õĳ�ֵ���Ŀ��ż����롣Ϊ�˰�ȫ�������������ʹ�ã�<br><br>�����������Ǳ�վ�ĳ�ֵ��������ֱ�ӵ������ֵ����ֵ�����ӽ��г�ֵ��<br>������������������˾�Ŀ����뾡��ȥ��ع�˾����վ�ĳ�ֵ��ڽ��г�ֵ��</td></tr>"
    Response.Write "</table>"
End Sub

Sub SaveExchange()
    If UserSetting(18) = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��������������һ�" & PointName & "��</li>"
        Exit Sub
    End If

    Dim rsUser, sqlUser
    Dim ChangeType, ChangeMoney, ChangeExp, GetPoint
    ChangeType = Abs(PE_CLng(Trim(Request("ChangeType"))))
    ChangeMoney = Abs(PE_CDbl(Trim(Request("ChangeMoney"))))
    ChangeExp = Abs(PE_CLng(Trim(Request("ChangeExp"))))

    If ChangeType = 1 Then 'ʹ�û���
        If ChangeMoney = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>������Ҫ�һ����ʽ�����</li>"
        Else
            If ChangeMoney > Balance Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>������ʽ������������ʽ���</li>"
            Else
                If Fix(ChangeMoney / MoneyExchangePoint) < 1 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>������ʽ��������Զһ� 1 " & PointUnit & PointName & "��</li>"
                End If
            End If
        End If
    Else  'ʹ�û���
        If ChangeExp = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>������Ҫ��ȥ�Ļ�������</li>"
        Else
            If ChangeExp > UserExp Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>����Ļ������������Ŀ��û��֣�</li>"
            Else
                If Fix(ChangeExp / UserExpExchangePoint) < 1 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>����Ļ����������Զһ� 1 " & PointUnit & PointName & "��</li>"
                End If
            End If
        End If
    End If

    If FoundErr = True Then
        Exit Sub
    End If

    Set rsUser = Server.CreateObject("Adodb.RecordSet")
    sqlUser = "select * from PE_User where UserID=" & UserID
    rsUser.Open sqlUser, Conn, 1, 3

    If ChangeType = 1 Then
        GetPoint = Fix(ChangeMoney / MoneyExchangePoint)
        rsUser("Balance") = rsUser("Balance") - ChangeMoney
        rsUser("UserPoint") = rsUser("UserPoint") + GetPoint
        Call AddBankrollItem("System", UserName, ClientID, ChangeMoney, 4, "", 0, 2, 0, 0, "���ڶһ� " & GetPoint & " " & PointUnit & PointName, Now())
        Call AddConsumeLog("System", 0, UserName, 0, GetPoint, 1, "�� " & ChangeMoney & " Ԫ�ʽ�һ��� " & GetPoint & " " & PointUnit & PointName)
        Call WriteSuccessMsg("�ɹ��� " & ChangeMoney & " Ԫ�ʽ�һ��� " & GetPoint & " " & PointUnit & PointName & " ��", ComeUrl)
    Else
        GetPoint = Fix(ChangeExp / UserExpExchangePoint)
        rsUser("UserExp") = rsUser("UserExp") - ChangeExp
        rsUser("UserPoint") = rsUser("UserPoint") + GetPoint
        Call AddConsumeLog("System", 0, UserName, 0, GetPoint, 1, "�� " & ChangeExp & " �ֻ��ֶһ��� " & GetPoint & " " & PointUnit & PointName)
        Call WriteSuccessMsg("�ɹ��� " & ChangeExp & " �ֻ��ֶһ��� " & GetPoint & " " & PointUnit & PointName & " ��", ComeUrl)
    End If

    rsUser.Update
    rsUser.Close
    Set rsUser = Nothing
End Sub

Sub SaveValid()
    If UserSetting(19) = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��������������һ���Ч�ڣ�</li>"
        Exit Sub
    End If

    Dim rsUser, sqlUser
    Dim ChangeType, ChangeMoney, ChangeExp, GetValidDay
    ChangeType = Abs(PE_CLng(Trim(Request("ChangeType"))))
    ChangeMoney = Abs(PE_CDbl(Trim(Request("ChangeMoney"))))
    ChangeExp = Abs(PE_CLng(Trim(Request("ChangeExp"))))

    If ChangeType = 1 Then 'ʹ�û���
        If ChangeMoney = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>������Ҫ�һ����ʽ�����</li>"
        Else
            If ChangeMoney > Balance Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>������ʽ������������ʽ���</li>"
            Else
                If Fix(ChangeMoney / MoneyExchangeValidDay) < 1 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>������ʽ��������Զһ� 1 ����Ч�ڣ�</li>"
                End If
            End If
        End If
    Else  'ʹ�û���
        If ChangeExp = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>������Ҫ��ȥ�Ļ�������</li>"
        Else
            If ChangeExp > UserExp Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>����Ļ������������Ŀ��û��֣�</li>"
            Else
                If Fix(ChangeExp / UserExpExchangeValidDay) < 1 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>����Ļ����������Զһ� 1 ����Ч�ڣ�</li>"
                End If
            End If
        End If
    End If

    If FoundErr = True Then
        Exit Sub
    End If
    
    Set rsUser = Server.CreateObject("Adodb.RecordSet")
    sqlUser = "select * from PE_User where UserID=" & UserID
    rsUser.Open sqlUser, Conn, 1, 3

    If rsUser("ValidNum") = -1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������Ч��Ϊ�������ڡ�������һ���Ч�ڡ�"
    Else
        If ChangeType = 1 Then
            GetValidDay = Fix(ChangeMoney / MoneyExchangeValidDay)
            rsUser("Balance") = rsUser("Balance") - ChangeMoney
            Call AddBankrollItem("System", UserName, ClientID, ChangeMoney, 4, "", 0, 2, 0, 0, "���ڶһ� " & GetValidDay & " ����Ч��", Now())
        Else
            GetValidDay = Fix(ChangeExp / UserExpExchangeValidDay)
            rsUser("UserExp") = rsUser("UserExp") - ChangeExp
        End If

        If ValidDays > 0 Then
            If rsUser("ValidUnit") = 1 Then
                rsUser("ValidNum") = rsUser("ValidNum") + GetValidDay
                rsUser.Update
            Else
                rsUser("ValidNum") = ValidNumToValidDays(rsUser("ValidNum"), rsUser("ValidUnit"), rsUser("BeginTime")) + GetValidDay
                rsUser("ValidUnit") = 1
                rsUser.Update
                Call AddRechargeLog("System", UserName, 0, 0, 0, "�һ���Ч��ʱ������Ч�ڼƷѵ�λ")
            End If
        Else
            rsUser("BeginTime") = Now()
            rsUser("ValidNum") = GetValidDay
            rsUser("ValidUnit") = 1
            rsUser.Update
            Call AddRechargeLog("System", UserName, 0, 0, 0, "�һ���Ч��ʱ��ԭ�����ڵ���Ч�����¼���")
        End If

        If ChangeType = 1 Then
            Call AddRechargeLog("System", UserName, GetValidDay, 1, 1, "�� " & ChangeMoney & " Ԫ�ʽ�һ��� " & GetValidDay & " ����Ч��")
            Call WriteSuccessMsg("�ɹ��� " & ChangeMoney & " Ԫ�ʽ�һ��� " & GetValidDay & " ����Ч�ڣ�", ComeUrl)
        Else
            Call AddRechargeLog("System", UserName, GetValidDay, 1, 1, "�� " & ChangeExp & " �ֻ��ֶһ��� " & GetValidDay & " ����Ч��")
            Call WriteSuccessMsg("�ɹ��� " & ChangeExp & " �ֻ��ֶһ��� " & GetValidDay & " ����Ч�ڣ�", ComeUrl)
        End If
    End If
    rsUser.Close
    Set rsUser = Nothing
End Sub


Sub SaveRecharge()
    Dim CardNum, Password
    Dim rsCard
    CardNum = ReplaceBadChar(Trim(Request("CardNum")))
    Password = ReplaceBadChar(Trim(Request("Password")))
    If CardNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������ֵ�����ţ�</li>"
    End If
    If Password = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������ֵ�����룡</li>"
    Else
        Password = Base64encode(Password)
    End If
    If FoundErr = True Then Exit Sub
    
    Set rsCard = Server.CreateObject("Adodb.Recordset")
    rsCard.Open "select * from PE_Card where CardNum='" & CardNum & "' and Password='" & Password & "'", Conn, 1, 3
    If rsCard.BOF And rsCard.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���Ż��������</li>"
    Else
        If rsCard("CardType") <> 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>������ĳ�ֵ����������˾�Ŀ��������ڱ�վ���г�ֵ���뾡��ȥ�йع�˾����վ�ĳ�ֵ��ڽ��г�ֵ��</li>"
        End If
        If rsCard("UserName") <> "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>������ĳ�ֵ���Ѿ�ʹ�ù��ˣ�</li>"
        End If
        If rsCard("EndDate") < Date Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>������ĳ�ֵ���Ѿ�ʧЧ���˿��ĳ�ֵ��ֹ����Ϊ��" & rsCard("EndDate")
        End If
    End If
    If FoundErr = True Then
        rsCard.Close
        Set rsCard = Nothing
        Exit Sub
    End If
    

    Dim strMsg
    strMsg = "��ֵ�ɹ���"
    If rsCard("ValidUnit") = 5 Then
        strMsg = strMsg & "&nbsp;&nbsp;&nbsp;<font color='red'>��ϲ���������� ��" & GetValidNum(rsCard("ValidNum"), rsCard("ValidUnit")) & "��</font>"
    End If
    strMsg = strMsg & "<br><br>��ֵ�����ţ�" & rsCard("CardNum") & "<br>"
    strMsg = strMsg & "��ֵ����ֵ��" & rsCard("Money") & "Ԫ" & "<br>"
    If rsCard("ValidUnit") = 5 Then
        strMsg = strMsg & "��Ա����"
    Else
        strMsg = strMsg & "��ֵ��������"
    End If
        strMsg = strMsg & GetValidNum(rsCard("ValidNum"), rsCard("ValidUnit")) & arrCardUnit(rsCard("ValidUnit")) & "<br>"
    strMsg = strMsg & "��ֵ��ֹ���ڣ�" & rsCard("EndDate") & "<br><br>"
    
    Dim rsUser, sqlUser
    Set rsUser = Server.CreateObject("Adodb.RecordSet")
    sqlUser = "select * from PE_User where UserID=" & UserID
    rsUser.Open sqlUser, Conn, 1, 3
    Select Case rsCard("ValidUnit")
    Case 0    '����
        strMsg = strMsg & "����ֵǰ��" & PointName & "����" & rsUser("UserPoint") & "<br>"
        rsUser("UserPoint") = rsUser("UserPoint") + rsCard("ValidNum")
        rsUser.Update
        strMsg = strMsg & "����ֵ���" & PointName & "����" & rsUser("UserPoint") & "<br>"
        Call AddConsumeLog("System", 0, UserName, 0, rsCard("ValidNum"), 1, "��ֵ����ֵ�����ţ�" & rsCard("CardNum") & "")
    Case 4    'Ԫ
        strMsg = strMsg & "����ֵǰ���ʽ����Ϊ�� " & rsUser("Balance") & " Ԫ<br>"
        rsUser("Balance") = rsUser("Balance") + rsCard("ValidNum")
        rsUser.Update
        strMsg = strMsg & "����ֵ����ʽ����Ϊ�� " & rsUser("Balance") & " Ԫ<br>"
        
        Call AddBankrollItem("System", UserName, ClientID, rsCard("ValidNum"), 4, "", 0, 1, 0, 0, "��ֵ����ֵ�����ţ�" & rsCard("CardNum") & "", Now())

    Case 5  '��Ա��
        Conn.Execute ("Update PE_User Set GroupID = " & rsCard("ValidNum") & " where UserName='" & UserName & "'")

    Case Else    '��Ч��
        If rsUser("ValidNum") = -1 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>������Ч��Ϊ�������ڡ��������ֵ��"
        Else
            If ValidDays > 0 Then
                strMsg = strMsg & "����ֵǰ����Ч�ڣ�" & rsUser("ValidNum") & arrCardUnit(rsUser("ValidUnit")) & "<br>"
                If rsUser("ValidUnit") = rsCard("ValidUnit") Then
                    rsUser("ValidNum") = rsUser("ValidNum") + rsCard("ValidNum")
                    rsUser.Update
                ElseIf rsUser("ValidUnit") < rsCard("ValidUnit") Then
                    If rsUser("ValidUnit") = 1 Then
                        If rsCard("ValidUnit") = 2 Then
                            rsUser("ValidNum") = rsUser("ValidNum") + rsCard("ValidNum") * 30
                        Else
                            rsUser("ValidNum") = rsUser("ValidNum") + rsCard("ValidNum") * 365
                        End If
                    Else
                        rsUser("ValidNum") = rsUser("ValidNum") + rsCard("ValidNum") * 12
                    End If
                    rsUser.Update
                Else
                    If rsCard("ValidUnit") = 1 Then
                        If rsUser("ValidUnit") = 2 Then
                            rsUser("ValidNum") = rsCard("ValidNum") + rsUser("ValidNum") * 30
                        Else
                            rsUser("ValidNum") = rsCard("ValidNum") + rsUser("ValidNum") * 365
                        End If
                    Else
                        rsUser("ValidNum") = rsCard("ValidNum") + rsUser("ValidNum") * 12
                    End If
                    rsUser("ValidUnit") = rsCard("ValidUnit")
                    rsUser.Update

                    Call AddRechargeLog("System", UserName, 0, 0, 0, "��ֵ����ֵʱ������Ч�ڼƷѵ�λ�����ţ�" & rsCard("CardNum") & "")
                End If
                strMsg = strMsg & "����ֵ�����Ч�ڣ�" & rsUser("ValidNum") & arrCardUnit(rsUser("ValidUnit")) & "<br>"
            Else
                strMsg = strMsg & "����ֵǰ��Ч���Ѿ����� " & Abs(ValidDays) & " ��<br>"
                rsUser("BeginTime") = Now()
                rsUser("ValidNum") = rsCard("ValidNum")
                rsUser("ValidUnit") = rsCard("ValidUnit")
                rsUser.Update
                strMsg = strMsg & "����ֵ�����Ч�ڣ�" & rsUser("ValidNum") & arrCardUnit(rsUser("ValidUnit")) & "����ʼ�������ڣ�" & Date & "<br>"
                Call AddRechargeLog("System", UserName, 0, 0, 0, "��ֵ����ֵʱ��ԭ�����ڵ���Ч�����¼��㡣���ţ�" & rsCard("CardNum") & "")
            End If
            Call AddRechargeLog("System", UserName, rsCard("ValidNum"), rsCard("ValidUnit"), 1, "��ֵ����ֵ�����ţ�" & rsCard("CardNum") & "")
        End If
    End Select
    
    If FoundErr = False Then
        rsCard("UserName") = UserName
        rsCard("UseTime") = Now()
        rsCard.Update
        Call WriteSuccessMsg(strMsg, "")
    End If
    rsUser.Close
    Set rsUser = Nothing
    rsCard.Close
    Set rsCard = Nothing
End Sub

Sub SaveSendPoint()
    If UserSetting(20) = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������" & PointName & "���͸����ˣ�</li>"
        Exit Sub
    End If
    Dim SendObject, SendPoint, i, j
    Dim arrSendObject
    Dim rsUser, rsObject
    
    SendObject = Trim(Request("SendObject"))
    SendPoint = PE_CLng(Trim(Request("SendPoint")))
    If SendObject = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������Է����û�����</li>"
    Else
        If CheckBadChar(SendObject) = False Then
            ErrMsg = ErrMsg + "<li>�û����к��зǷ��ַ�</li>"
            FoundErr = True
        End If
    End If
    If SendPoint <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����δ����" & PointName & "���������" & PointName & "���д��ڷǷ��ַ���</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    j = 0
    arrSendObject = Split(SendObject, ",")
    Set rsUser = Conn.Execute("select * from PE_User where UserID=" & UserID & "")
    If rsUser("UserPoint") - SendPoint * (UBound(arrSendObject) + 1) < 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����" & PointName & "������</li>"
        Exit Sub
    Else
        For i = 0 To UBound(arrSendObject)
            Set rsObject = Conn.Execute("select UserID from PE_User where UserName='" & arrSendObject(i) & "'")
            If Not rsObject.EOF Then
                Conn.Execute "Update PE_User set UserPoint=UserPoint + " & SendPoint & "  where UserName='" & arrSendObject(i) & "'"
                Conn.Execute "Update PE_User set UserPoint=UserPoint - " & SendPoint & " where UserID=" & UserID
                Call AddConsumeLog("System", 0, UserName, 0, SendPoint, 2, "��" & arrSendObject(i) & "�û�����" & PointName & "")
                Call AddConsumeLog("System", 0, arrSendObject(i), 0, SendPoint, 1, "���" & UserName & "�û����͵�" & PointName & "")
                Conn.Execute "Insert into PE_Message (Incept,Sender,Title,IsSend,Content,Flag) values('" & arrSendObject(i) & "','" & UserName & "','����" & PointName & "',1,'" & UserName & "������" & PointName & "" & SendPoint & PointUnit & "',0)"
            Else
                j = j + 1
            End If
            Set rsObject = Nothing
        Next
        If j = 0 Then
           Call WriteSuccessMsg(PointName & "���ͳɹ���", ComeUrl)
        Else
           Call WriteSuccessMsg("��" & UBound(arrSendObject) - j + 1 & "λ�û����ͳɹ���������" & j & "λ�û������ڣ�", ComeUrl)
        End If
    End If
    rsUser.Close
    Set rsUser = Nothing
End Sub

Function GetValidNum(intValidNum, intValidUnit)
    If intValidUnit = 5 Then
        Dim rsGroupList
        Set rsGroupList = Conn.Execute("Select GroupName from PE_UserGroup where GroupID = " & intValidNum)
        If Not (rsGroupList.EOF And rsGroupList.BOF) Then
            GetValidNum = rsGroupList("GroupName")
        Else
            GetValidNum = intValidNum
        End If
        rsGroupList.Close
        Set rsGroupList = Nothing
    Else
        GetValidNum = intValidNum
    End If
End Function


'**************************************************
'��������ValidNumToValidDays
'��  �ã�ת����Ч��Ϊ��Ч����
'��  ����iValidNum ----��Ч��
'        iValidUnit ----��Ч�ڵ�λ
'        iBeginTime ---- ��ʼ��������
'����ֵ����Ч����
'**************************************************
Function ValidNumToValidDays(iValidNum, iValidUnit, iBeginTime)
    If (iValidNum = "" Or IsNumeric(iValidNum) = False Or iValidUnit = "" Or IsNumeric(iValidUnit) = False Or iBeginTime = "" Or IsDate(iBeginTime) = False) Then
        ValidNumToValidDays = 0
        Exit Function
    End If
    Dim tmpDate, arrInterval
    arrInterval = Array("h", "D", "m", "yyyy")
    If iValidNum = -1 Then
        ValidNumToValidDays = 99999
    Else
        ValidNumToValidDays = DateDiff("D", iBeginTime, DateAdd(arrInterval(iValidUnit), iValidNum, iBeginTime))
    End If
End Function
%>
