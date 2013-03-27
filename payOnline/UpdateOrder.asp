<!--#include file="../Include/PowerEasy.Bankroll.asp"-->
<!--#include file="../Include/PowerEasy.Base64.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim PlatformName, AccountsID, MD5Key, Rate
Dim rsPayPlatform
Set rsPayPlatform = Conn.Execute("select * from PE_PayPlatform where PlatformID=" & PlatformID & "")
If rsPayPlatform.BOF And rsPayPlatform.EOF Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>找不到指定的在线支付平台！</li>"
Else
    If PE_CLng(rsPayPlatform("IsDisabled")) = -1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>该支付平台未启用！</li>"
    Else
        Select Case PlatformID
        Case 1 
            If rsPayPlatform("MD5Key") = "sldkfjkalsdjfasdf" Then 
                FoundErr = True
                ErrMsg = ErrMsg & "<li>为了您交易的安全,请不要使用系统默认的MD5密钥！</li>"                
            End If						          									
        Case Else
            If rsPayPlatform("MD5Key") = "aaaaaaaaaa" Then 
                FoundErr = True
                ErrMsg = ErrMsg & "<li>为了您交易的安全,请不要使用系统默认的MD5密钥！</li>"                
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
        ErrMsg = "<li>找不到指定的在线支付平台！</li>"
    Else
        If PE_CLng(rsCheck("IsDisabled")) = -1 Then
            FoundErr = True
            ErrMsg = "<li>该支付平台未启用！</li>"
        Else
            Select Case PlatformID
            Case 1 
                If rsCheck("MD5Key") = "sldkfjkalsdjfasdf" Then 
                    FoundErr = True
                    ErrMsg = "<li>为了您交易的安全,请不要使用系统默认的MD5密钥！</li>"                
                End If						          									
            Case else
                If rsCheck("MD5Key") = "aaaaaaaaaa" Then 
                    FoundErr = True
                    ErrMsg = "<li>为了您交易的安全,请不要使用系统默认的MD5密钥！</li>"                
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
          Response.Write "找不到指定的支付单！"
        End If
    Else
        If rsPayment("MoneyTrue") <> CCur(amount) Then
            FoundErr = True
            If IsMessageShow = True Then
              Response.Write "<li>支付金额不对！</li>"
            End If
        Else
            PaymentID = rsPayment("PaymentID")   '支户ID
            UserName = rsPayment("UserName")
            OrderFormID = rsPayment("OrderFormID")   '定单ID
            MoneyReceipt = rsPayment("MoneyPay")  '支付金额
            rsPayment("Status") = Status
            rsPayment("eBankInfo") = eBankInfo    '支付银行信息
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
        Call AddBankrollItem("System", UserName, ClientID, MoneyReceipt, 3, "", PlatformID, 1, 0, PaymentID, "在线支付单号：" & PaymentNum, Now())
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
                If rsOrder("MoneyReceipt") < rsOrder("MoneyTotal") And UpdateOrderStatus = True Then   'MoneyTotal:订单总金额
                    If rsOrder("MoneyTotal") - rsOrder("MoneyReceipt") > MoneyReceipt Then
                        MoneyPayout = MoneyReceipt                        ' MoneyReceipt:'支付金额
                        rsOrder("MoneyReceipt") = rsOrder("MoneyReceipt") + MoneyReceipt
                    Else
                        MoneyPayout = rsOrder("MoneyTotal") - rsOrder("MoneyReceipt")
                        rsOrder("MoneyReceipt") = rsOrder("MoneyTotal")
                        
                    End If
                    If rsOrder("Status") <= 2 Then
                        rsOrder("Status") = 3
                    End If
                    rsOrder.Update
                    '向资金明细表中添加支付记录
                    Call AddBankrollItem("System", UserName, ClientID, MoneyPayout, 4, "", 0, 2, OrderFormID, 0, "支付订单费用，订单号：" & rsOrder("OrderFormNum"), Now())
                                    
                    '从资金余额中扣除支付费用
                    Conn.Execute ("update PE_User set Balance=Balance-" & MoneyPayout & " where UserName='" & UserName & "'")
                    If IsMessageShow = True Then
                        Response.Write "同时已经为您的订单编号为 " & rsOrder("OrderFormNum") & " 的订单支付了 " & FormatNumber(MoneyPayout, 2, vbTrue, vbFalse, vbTrue) & "元。<br>"
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
                                    Response.Write "<br><br>您购买的充值卡的信息如下，请您尽快使用，以防充值卡被他人使用！<br>"
                                End If
                                Do While Not rsCard.EOF
                                    If IsMessageShow = True Then
                                        Response.Write "<br>卡号：" & rsCard("CardNum") & "&nbsp;&nbsp;&nbsp;&nbsp;密码：" & Base64decode(rsCard("Password"))
                                    End If
                                    rsCard(2) = rsOrderItem(0)
                                    rsCard.Update
                                    rsCard.MoveNext
                                Loop
                                Conn.Execute ("update PE_Product set Stocks=Stocks-" & rsOrderItem("Amount") & ",OrderNum=OrderNum-" & rsOrderItem("Amount") & " where ProductID=" & rsOrderItem("ProductID") & "")
                                If IsMessageShow = True Then
                                    Response.Write "<br><br><a href='../User/User_Exchange.asp?Action=Recharge'>使用充值卡充值</a>&nbsp;&nbsp;&nbsp;&nbsp;"
                                End If
                            Else
                                HaveOtherProduct = True
                            End If
                            rsCard.Close
                            Set rsCard = Nothing
                        End If
                    Else
                        If IsMessageShow = True Then
                            Response.Write "<br><br>您购买的充值卡的信息如下，请您尽快使用，以防充值卡被他人使用！<br>"
                            Do While Not trs.EOF
                                Response.Write "<br>卡号：" & trs("CardNum") & "&nbsp;&nbsp;&nbsp;&nbsp;密码：" & Base64decode(trs("Password"))
                                trs.MoveNext
                            Loop
                            Response.Write "<br><br><a href='../User/User_Exchange.asp?Action=Recharge'>使用充值卡充值</a>&nbsp;&nbsp;&nbsp;&nbsp;"
                        End If
                    End If
                    Set trs = Nothing
                Else
                    HaveOtherProduct = True
                End If
                rsOrderItem.MoveNext
            Loop
            Set rsOrderItem = Nothing

            If HaveOtherProduct = False And DoUpdate = True Then  '如果该定单中全部是虚拟物品的话,物流状态为已发货
                rsOrder("DeliverStatus") = 2
                rsOrder.Update
            End If

            If IsMessageShow = True Then
                Response.Write "<a href='../User/User_Order.asp?Action=ShowOrder&OrderFormID=" & OrderFormID & "'>点此查看订单信息</a>"
            End If
        End If
        rsOrder.Close
        Set rsOrder = Nothing
    End If
End Sub
%>

