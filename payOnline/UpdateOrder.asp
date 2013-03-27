<!--#include file="../Include/PowerEasy.Bankroll.asp"-->
<!--#include file="../Include/PowerEasy.Base64.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim PlatformName, AccountsID, MD5Key, Rate
Dim rsPayPlatform
Set rsPayPlatform = Conn.Execute("select * from PE_PayPlatform where PlatformID=" & PlatformID & "")
If rsPayPlatform.BOF And rsPayPlatform.EOF Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>�Ҳ���ָ��������֧��ƽ̨��</li>"
Else
    If PE_CLng(rsPayPlatform("IsDisabled")) = -1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��֧��ƽ̨δ���ã�</li>"
    Else
        Select Case PlatformID
        Case 1 
            If rsPayPlatform("MD5Key") = "sldkfjkalsdjfasdf" Then 
                FoundErr = True
                ErrMsg = ErrMsg & "<li>Ϊ�������׵İ�ȫ,�벻Ҫʹ��ϵͳĬ�ϵ�MD5��Կ��</li>"                
            End If						          									
        Case Else
            If rsPayPlatform("MD5Key") = "aaaaaaaaaa" Then 
                FoundErr = True
                ErrMsg = ErrMsg & "<li>Ϊ�������׵İ�ȫ,�벻Ҫʹ��ϵͳĬ�ϵ�MD5��Կ��</li>"                
            End If			
        End Select		       					    
    End If	
    PlatformName = rsPayPlatform("ShowName")
    AccountsID = rsPayPlatform("AccountsID")
    MD5Key = rsPayPlatform("MD5Key")
    Rate = rsPayPlatform("Rate")
End If
Set rsPayPlatform = Nothing

Sub CheckPlatformID(thePlatformID)
    Dim rsCheck
    Set rsCheck = Conn.Execute("select * from PE_PayPlatform where PlatformID=" & thePlatformID & "")
    If rsCheck.BOF And rsCheck.EOF Then
        FoundErr = True
        ErrMsg = "<li>�Ҳ���ָ��������֧��ƽ̨��</li>"
    Else
        If PE_CLng(rsCheck("IsDisabled")) = -1 Then
            FoundErr = True
            ErrMsg = "<li>��֧��ƽ̨δ���ã�</li>"
        Else
            Select Case PlatformID
            Case 1 
                If rsCheck("MD5Key") = "sldkfjkalsdjfasdf" Then 
                    FoundErr = True
                    ErrMsg = "<li>Ϊ�������׵İ�ȫ,�벻Ҫʹ��ϵͳĬ�ϵ�MD5��Կ��</li>"                
                End If						          									
            Case else
                If rsCheck("MD5Key") = "aaaaaaaaaa" Then 
                    FoundErr = True
                    ErrMsg = "<li>Ϊ�������׵İ�ȫ,�벻Ҫʹ��ϵͳĬ�ϵ�MD5��Կ��</li>"                
                End If			
            End Select		       					    
        End If	
    End If
    Set rsCheck = Nothing	
    If FoundErr = True Then 
        Call WriteErrMsg(ErrMsg,ComeUrl)
        Response.end
    End If		
End Sub

Sub UpdateOrder(ByVal PaymentNum, ByVal amount, ByVal eBankInfo, ByVal Remark, Status, UpdateDeliverStatus, UpdateOrderStatus)
    Dim PaymentID, OrderFormID, MoneyReceipt, MoneyPayout, eBankID
    Dim sqlPayment, rsPayment
    Dim DoUpdate

    PaymentNum = ReplaceBadChar(PaymentNum)
    sqlPayment = "select * from PE_Payment where PaymentNum='" & PaymentNum & "'"
    Set rsPayment = Server.CreateObject("Adodb.RecordSet")
    rsPayment.Open sqlPayment, Conn, 1, 3
    If rsPayment.BOF And rsPayment.EOF Then
        FoundErr = True
        If IsMessageShow = True Then
          Response.Write "�Ҳ���ָ����֧������"
        End If
    Else
        If rsPayment("MoneyTrue") <> CCur(amount) Then
            FoundErr = True
            If IsMessageShow = True Then
              Response.Write "<li>֧�����ԣ�</li>"
            End If
        Else
            PaymentID = rsPayment("PaymentID")   '֧��ID
            UserName = rsPayment("UserName")
            OrderFormID = rsPayment("OrderFormID")   '����ID
            MoneyReceipt = rsPayment("MoneyPay")  '֧�����
            rsPayment("Status") = Status
            rsPayment("eBankInfo") = eBankInfo    '֧��������Ϣ
            rsPayment("Remark") = Remark
            rsPayment.Update
        End If
    End If
    rsPayment.Close
    Set rsPayment = Nothing
    
    If FoundErr = True Or UpdateDeliverStatus = False Then Exit Sub
    
    Dim sqlTrs, rsTrs, trs
    Set rsTrs = Conn.Execute("select * from PE_BankrollItem where PaymentID=" & PaymentID & "")
    If rsTrs.BOF And rsTrs.EOF Then
        DoUpdate = True
    Else
        DoUpdate = False
    End If
    Set rsTrs = Nothing
    
    Set trs = Conn.Execute("select ClientID from PE_User where UserName='" & UserName & "'")
    If trs.BOF And trs.EOF Then
        ClientID = 0
    Else
        ClientID = trs(0)
    End If
    Set trs = Nothing
    
    If DoUpdate = True And UpdateOrderStatus = True Then
        Conn.Execute ("update PE_User set Balance=Balance+" & MoneyReceipt & " where UserName='" & UserName & "'")
        Call AddBankrollItem("System", UserName, ClientID, MoneyReceipt, 3, "", PlatformID, 1, 0, PaymentID, "����֧�����ţ�" & PaymentNum, Now())
    End If

    If OrderFormID > 0 Then
        Dim rsOrder
        Dim strCardInfo
        strCardInfo = ""
        Set rsOrder = Server.CreateObject("adodb.recordset")
        rsOrder.Open "select * from PE_OrderForm where OrderFormID=" & OrderFormID & "", Conn, 1, 3
        If Not (rsOrder.BOF And rsOrder.EOF) Then
            If DoUpdate = True Then
                If UpdateDeliverStatus = True And rsOrder("MoneyTotal") - rsOrder("MoneyReceipt") <= MoneyReceipt Then
                    rsOrder("EnableDownload") = True
                    rsOrder.Update
                End If
                If rsOrder("MoneyReceipt") < rsOrder("MoneyTotal") And UpdateOrderStatus = True Then   'MoneyTotal:�����ܽ��
                    If rsOrder("MoneyTotal") - rsOrder("MoneyReceipt") > MoneyReceipt Then
                        MoneyPayout = MoneyReceipt                        ' MoneyReceipt:'֧�����
                        rsOrder("MoneyReceipt") = rsOrder("MoneyReceipt") + MoneyReceipt
                    Else
                        MoneyPayout = rsOrder("MoneyTotal") - rsOrder("MoneyReceipt")
                        rsOrder("MoneyReceipt") = rsOrder("MoneyTotal")
                        
                    End If
                    If rsOrder("Status") <= 2 Then
                        rsOrder("Status") = 3
                    End If
                    rsOrder.Update
                    '���ʽ���ϸ�������֧����¼
                    Call AddBankrollItem("System", UserName, ClientID, MoneyPayout, 4, "", 0, 2, OrderFormID, 0, "֧���������ã������ţ�" & rsOrder("OrderFormNum"), Now())
                                    
                    '���ʽ�����п۳�֧������
                    Conn.Execute ("update PE_User set Balance=Balance-" & MoneyPayout & " where UserName='" & UserName & "'")
                    If IsMessageShow = True Then
                        Response.Write "ͬʱ�Ѿ�Ϊ���Ķ������Ϊ " & rsOrder("OrderFormNum") & " �Ķ���֧���� " & FormatNumber(MoneyPayout, 2, vbTrue, vbFalse, vbTrue) & "Ԫ��<br>"
                    End If
                End If
            End If

            Dim rsOrderItem
            Dim HaveOtherProduct
            Set rsOrderItem = Conn.Execute("select I.ItemID,P.ProductID,P.ProductName,P.ProductKind,I.Amount from PE_OrderFormItem I inner join PE_Product P on I.ProductID=P.ProductID where I.OrderFormID=" & OrderFormID & " order by I.ItemID")
            Do While Not rsOrderItem.EOF
                If rsOrderItem("ProductKind") = 3 Then
                    Set trs = Conn.Execute("select * from PE_Card where ProductID=" & rsOrderItem("ProductID") & " and OrderFormItemID=" & rsOrderItem("ItemID") & "")
                    If trs.BOF And trs.EOF Then
                        If DoUpdate = True Then
                            Dim rsCard, sqlCard
                            Set rsCard = Server.CreateObject("Adodb.Recordset")
                            sqlCard = "select top " & rsOrderItem("Amount") & " CardID,ProductID,OrderFormItemID,CardNum,Password from PE_Card where ProductID=" & rsOrderItem("ProductID") & " and OrderFormItemID=0 order by CardID"
                            rsCard.Open sqlCard, Conn, 1, 3
                            If rsCard.RecordCount >= rsOrderItem("Amount") Then
                                If IsMessageShow = True Then
                                    Response.Write "<br><br>������ĳ�ֵ������Ϣ���£���������ʹ�ã��Է���ֵ��������ʹ�ã�<br>"
                                End If
                                Do While Not rsCard.EOF
                                    If IsMessageShow = True Then
                                        Response.Write "<br>���ţ�" & rsCard("CardNum") & "&nbsp;&nbsp;&nbsp;&nbsp;���룺" & Base64decode(rsCard("Password"))
                                    End If
                                    rsCard(2) = rsOrderItem(0)
                                    rsCard.Update
                                    rsCard.MoveNext
                                Loop
                                Conn.Execute ("update PE_Product set Stocks=Stocks-" & rsOrderItem("Amount") & ",OrderNum=OrderNum-" & rsOrderItem("Amount") & " where ProductID=" & rsOrderItem("ProductID") & "")
                                If IsMessageShow = True Then
                                    Response.Write "<br><br><a href='../User/User_Exchange.asp?Action=Recharge'>ʹ�ó�ֵ����ֵ</a>&nbsp;&nbsp;&nbsp;&nbsp;"
                                End If
                            Else
                                HaveOtherProduct = True
                            End If
                            rsCard.Close
                            Set rsCard = Nothing
                        End If
                    Else
                        If IsMessageShow = True Then
                            Response.Write "<br><br>������ĳ�ֵ������Ϣ���£���������ʹ�ã��Է���ֵ��������ʹ�ã�<br>"
                            Do While Not trs.EOF
                                Response.Write "<br>���ţ�" & trs("CardNum") & "&nbsp;&nbsp;&nbsp;&nbsp;���룺" & Base64decode(trs("Password"))
                                trs.MoveNext
                            Loop
                            Response.Write "<br><br><a href='../User/User_Exchange.asp?Action=Recharge'>ʹ�ó�ֵ����ֵ</a>&nbsp;&nbsp;&nbsp;&nbsp;"
                        End If
                    End If
                    Set trs = Nothing
                Else
                    HaveOtherProduct = True
                End If
                rsOrderItem.MoveNext
            Loop
            Set rsOrderItem = Nothing

            If HaveOtherProduct = False And DoUpdate = True Then  '����ö�����ȫ����������Ʒ�Ļ�,����״̬Ϊ�ѷ���
                rsOrder("DeliverStatus") = 2
                rsOrder.Update
            End If

            If IsMessageShow = True Then
                Response.Write "<a href='../User/User_Order.asp?Action=ShowOrder&OrderFormID=" & OrderFormID & "'>��˲鿴������Ϣ</a>"
            End If
        End If
        rsOrder.Close
        Set rsOrder = Nothing
    End If
End Sub
%>

