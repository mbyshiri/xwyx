<!--#include file="CommonCode.asp"-->
<!--#include file="../Include/PowerEasy.Bankroll.asp"-->
<!--#include file="../Include/PowerEasy.ConsumeLog.asp"-->
<!--#include file="../Include/PowerEasy.RechargeLog.asp"-->
<!--#include file="../Include/PowerEasy.Base64.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
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
        ErrMsg = ErrMsg & "<li>错误的参数</li>"
    End Select
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
End Sub

Sub ShowUserInfo()
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>用户名：</td>"
    Response.Write "      <td>" & UserName & "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>资金余额：</td>"
    Response.Write "      <td>" & Balance & " 元</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>可用积分：</td>"
    Response.Write "      <td>" & UserExp & " 分</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>可用" & PointName & "数：</td>"
    Response.Write "      <td>" & UserPoint & " " & PointUnit & "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>有效天数：</td>"
    Response.Write "      <td>开始计算日期：" & FormatDateTime(BeginTime, 2) & "&nbsp;&nbsp;&nbsp;&nbsp;有效期："
    If ValidNum = -1 Then
        Response.Write "无限期<br>"
    Else
        Response.Write ValidNum & arrCardUnit(ValidUnit) & "<br>"
        If ValidDays >= 0 Then
            Response.Write "尚有 <font color=blue>" & ValidDays & "</font> 天到期"
        Else
            Response.Write "已经过期 <font color=red>" & Abs(ValidDays) & "</font> 天"
        End If
    End If
    Response.Write "      </td>"
    Response.Write "    </tr>"
End Sub

Sub Exchange()
    If UserSetting(18) = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>不允许进行自助兑换" & PointName & "！</li>"
        Exit Sub
    End If
    Response.Write "<form name='myform' action='User_Exchange.asp' method='post'>"
    Response.Write "  <table width='100%' border='0' cellspacing='1' cellpadding='2' align='center' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height=22 colSpan=2 align='center'><b>兑换" & PointName & "</b></td>"
    Response.Write "    </tr>"
    Call ShowUserInfo
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>兑换" & PointName & "：</td>"
    Response.Write "      <td>"
    Response.Write "        <input type='radio' name='ChangeType' value='1' checked>使用资金余额："
    Response.Write "        将 <input name='ChangeMoney' type='text' value='10' size='6' maxlength='8' style='text-align:center'> 元兑换成" & PointName
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;兑换比率：" & FormatNumber(MoneyExchangePoint, 2, vbTrue, vbFalse, vbTrue) & "元:1" & PointUnit
    Response.Write "        <br>"
    Response.Write "        <input type='radio' name='ChangeType' value='2'>使用经验积分："
    Response.Write "        将 <input name='ChangeExp' type='text' value='10' size='6' maxlength='8' style='text-align:center'> 分兑换成" & PointName
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;兑换比率：" & FormatNumber(UserExpExchangePoint, 2, vbTrue, vbFalse, vbTrue) & "分:1" & PointUnit
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveExchange'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value='执行兑换'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub Valid()
    If UserSetting(19) = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>不允许进行自助兑换有效期！</li>"
        Exit Sub
    End If
    Response.Write "<form name='myform' action='User_Exchange.asp' method='post'>"
    Response.Write "  <table width='100%' border='0' cellspacing='1' cellpadding='2' align='center' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height=22 colSpan=2 align='center'><b>兑 换 有 效 期</b></td>"
    Response.Write "    </tr>"
    Call ShowUserInfo
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>兑换有效期：</td>"
    Response.Write "      <td>"
    Response.Write "        <input type='radio' name='ChangeType' value='1' checked>使用资金余额："
    Response.Write "        将 <input name='ChangeMoney' type='text' value='10' size='6' maxlength='8' style='text-align:center'> 元兑换成有效期"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;兑换比率：" & MoneyExchangeValidDay & "元:1天"
    Response.Write "        <br>"
    Response.Write "        <input type='radio' name='ChangeType' value='2'>使用经验积分："
    Response.Write "        将 <input name='ChangeExp' type='text' value='10' size='6' maxlength='8' style='text-align:center'> 分兑换成有效期"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;兑换比率：" & UserExpExchangeValidDay & "分:1天"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveValid'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value='执行兑换'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub Recharge()
    Response.Write "<form name='myform' action='User_Exchange.asp' method='post'>"
    Response.Write "  <table width='500' border='0' cellspacing='1' cellpadding='2' align='center' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height=22 colSpan=2 align='center'><b>充 值 卡 充 值</b></td>"
    Response.Write "    </tr>"
    Call ShowUserInfo
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>充值卡卡号：</td>"
    Response.Write "      <td><input name='CardNum' type='text' value='' size='30' maxlength='30'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120' align='right' class='tdbg5'>充值卡密码：</td>"
    Response.Write "      <td><input name='Password' type='text' value='' size='30' maxlength='30'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'><input name='Action' type='hidden' id='Action' value='SaveRecharge'>"
    Response.Write "        <input name=Submit   type=submit id='Submit' value=' 确 定 '></td>"
    Response.Write "    </tr>"
    Response.Write "  </TABLE>"
    Response.Write "</form>"
End Sub

Sub SendPoint()
    If UserSetting(20) = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>不允许将" & PointName & "赠送给他人！</li>"
        Exit Sub
    End If
    Response.Write "<form name='myform' action='User_Exchange.asp' method='post'>" & vbCrLf
    Response.Write "  <table width='500' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr height='22' align='center' class='title'>" & vbCrLf
    Response.Write "      <td colSpan='2'><b>赠送" & PointName & "</b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='120' align='right' class='tdbg5'>用 户 名：</td>" & vbCrLf
    Response.Write "      <td>" & UserName & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='120' align='right' class='tdbg5'>当前" & PointName & "数：</td>" & vbCrLf
    Response.Write "      <td>" & UserPoint & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='120' align='right' class='tdbg5'>获赠人的用户名：</td>" & vbCrLf
    Response.Write "      <td> <input name='SendObject' type='text' size='30'> </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='120' align='right' class='tdbg5'>赠送的" & PointName & "数：</td>" & vbCrLf
    Response.Write "      <td> <input name='SendPoint' type='text' maxLength='16' size='30'> </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <td height='40' colspan='2'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveSendPoint'>" & vbCrLf
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 赠送 '>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

Sub GetCard()
    Response.Write "<br><table width='100%' cellspacing='1' cellpadding='2'  class='border'><tr class='title'><td align='center'>获取虚拟充值卡</td></tr>"
    Response.Write "<tr><td height='100'>"
    Dim rsOrderItem, rsCard, sqlCard, i, strCardInfo
    Set rsOrderItem = Conn.Execute("select O.OrderFormID,I.ItemID,P.ProductID,P.ProductName,P.ProductKind,I.Amount from PE_OrderForm O inner join (PE_OrderFormItem I inner join PE_Product P on I.ProductID=P.ProductID) on I.OrderFormID=O.OrderFormID where O.UserName='" & UserName & "' and P.ProductKind=3 order by I.ItemID")
    If rsOrderItem.BOF And rsOrderItem.EOF Then
        Response.Write "您还没有购买任何点卡类商品！"
    Else
        Response.Write "<br><br><table width='80%' align='center' cellspacing='1' cellpadding='2'>"
        Response.Write "<tr class='title' align='center'><td>商品名称</td><td>充值卡类型</td><td>充值卡卡号</td><td>充值卡密码</td><td>充值卡面值</td><td>充值卡点数</td><td>充值截止日期</td></tr>"
        Do While Not rsOrderItem.EOF
            Set rsCard = Conn.Execute("select * from PE_Card where ProductID=" & rsOrderItem("ProductID") & " and OrderFormItemID=" & rsOrderItem("ItemID") & "")
            If rsCard.BOF And rsCard.EOF Then
                Response.Write "<tr class='tdbg' align='center'><td>" & rsOrderItem("Productname") & "</td><td colspan='10' align='center'>尚没有交付卡号和密码，请您与我们联系。</td></tr>"
            Else
                i = 0
                Do While Not rsCard.EOF
                    If rsCard("UserName") = "" Then
                        Response.Write "<tr class='tdbg' align='center'><td>" & rsOrderItem("Productname") & "</td>"
                        Response.Write "<td>"
                        If rsCard("CardType") = 0 Then
                            Response.Write "本站充值卡"
                        Else
                            Response.Write "<font color='blue'>其他公司卡</font>"
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
                    Response.Write "<tr class='tdbg' align='center'><td>" & rsOrderItem("Productname") & "</td><td colspan='10' height='50' align='center'>您购买的所有充值卡都已经使用。</td></tr>"
                End If
            End If
            Set rsCard = Nothing
            rsOrderItem.MoveNext
        Loop
        Response.Write "</table><br><br>"
    End If
    Set rsOrderItem = Nothing
    Response.Write "</td></tr>"
    Response.Write "<tr class='tdbg'><td><font color='red'>注意：</font><br>这里只显示了还未使用的充值卡的卡号及密码。为了安全起见，请您尽快使用！<br><br>如果您购买的是本站的充值卡，可以直接点击“充值卡充值”链接进行充值。<br>如果您购买的是其他公司的卡，请尽快去相关公司或网站的充值入口进行充值。</td></tr>"
    Response.Write "</table>"
End Sub

Sub SaveExchange()
    If UserSetting(18) = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>不允许进行自助兑换" & PointName & "！</li>"
        Exit Sub
    End If

    Dim rsUser, sqlUser
    Dim ChangeType, ChangeMoney, ChangeExp, GetPoint
    ChangeType = Abs(PE_CLng(Trim(Request("ChangeType"))))
    ChangeMoney = Abs(PE_CDbl(Trim(Request("ChangeMoney"))))
    ChangeExp = Abs(PE_CLng(Trim(Request("ChangeExp"))))

    If ChangeType = 1 Then '使用货币
        If ChangeMoney = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入要兑换的资金数！</li>"
        Else
            If ChangeMoney > Balance Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>输入的资金数大于您的资金余额！</li>"
            Else
                If Fix(ChangeMoney / MoneyExchangePoint) < 1 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>输入的资金数不足以兑换 1 " & PointUnit & PointName & "！</li>"
                End If
            End If
        End If
    Else  '使用积分
        If ChangeExp = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入要减去的积分数！</li>"
        Else
            If ChangeExp > UserExp Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>输入的积分数大于您的可用积分！</li>"
            Else
                If Fix(ChangeExp / UserExpExchangePoint) < 1 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>输入的积分数不足以兑换 1 " & PointUnit & PointName & "！</li>"
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
        Call AddBankrollItem("System", UserName, ClientID, ChangeMoney, 4, "", 0, 2, 0, 0, "用于兑换 " & GetPoint & " " & PointUnit & PointName, Now())
        Call AddConsumeLog("System", 0, UserName, 0, GetPoint, 1, "将 " & ChangeMoney & " 元资金兑换成 " & GetPoint & " " & PointUnit & PointName)
        Call WriteSuccessMsg("成功将 " & ChangeMoney & " 元资金兑换成 " & GetPoint & " " & PointUnit & PointName & " ！", ComeUrl)
    Else
        GetPoint = Fix(ChangeExp / UserExpExchangePoint)
        rsUser("UserExp") = rsUser("UserExp") - ChangeExp
        rsUser("UserPoint") = rsUser("UserPoint") + GetPoint
        Call AddConsumeLog("System", 0, UserName, 0, GetPoint, 1, "将 " & ChangeExp & " 分积分兑换成 " & GetPoint & " " & PointUnit & PointName)
        Call WriteSuccessMsg("成功将 " & ChangeExp & " 分积分兑换成 " & GetPoint & " " & PointUnit & PointName & " ！", ComeUrl)
    End If

    rsUser.Update
    rsUser.Close
    Set rsUser = Nothing
End Sub

Sub SaveValid()
    If UserSetting(19) = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>不允许进行自助兑换有效期！</li>"
        Exit Sub
    End If

    Dim rsUser, sqlUser
    Dim ChangeType, ChangeMoney, ChangeExp, GetValidDay
    ChangeType = Abs(PE_CLng(Trim(Request("ChangeType"))))
    ChangeMoney = Abs(PE_CDbl(Trim(Request("ChangeMoney"))))
    ChangeExp = Abs(PE_CLng(Trim(Request("ChangeExp"))))

    If ChangeType = 1 Then '使用货币
        If ChangeMoney = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入要兑换的资金数！</li>"
        Else
            If ChangeMoney > Balance Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>输入的资金数大于您的资金余额！</li>"
            Else
                If Fix(ChangeMoney / MoneyExchangeValidDay) < 1 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>输入的资金数不足以兑换 1 天有效期！</li>"
                End If
            End If
        End If
    Else  '使用积分
        If ChangeExp = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入要减去的积分数！</li>"
        Else
            If ChangeExp > UserExp Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>输入的积分数大于您的可用积分！</li>"
            Else
                If Fix(ChangeExp / UserExpExchangeValidDay) < 1 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>输入的积分数不足以兑换 1 天有效期！</li>"
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
        ErrMsg = ErrMsg & "<li>您的有效期为“无限期”，无需兑换有效期。"
    Else
        If ChangeType = 1 Then
            GetValidDay = Fix(ChangeMoney / MoneyExchangeValidDay)
            rsUser("Balance") = rsUser("Balance") - ChangeMoney
            Call AddBankrollItem("System", UserName, ClientID, ChangeMoney, 4, "", 0, 2, 0, 0, "用于兑换 " & GetValidDay & " 天有效期", Now())
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
                Call AddRechargeLog("System", UserName, 0, 0, 0, "兑换有效期时更改有效期计费单位")
            End If
        Else
            rsUser("BeginTime") = Now()
            rsUser("ValidNum") = GetValidDay
            rsUser("ValidUnit") = 1
            rsUser.Update
            Call AddRechargeLog("System", UserName, 0, 0, 0, "兑换有效期时将原来过期的有效期重新计算")
        End If

        If ChangeType = 1 Then
            Call AddRechargeLog("System", UserName, GetValidDay, 1, 1, "将 " & ChangeMoney & " 元资金兑换成 " & GetValidDay & " 天有效期")
            Call WriteSuccessMsg("成功将 " & ChangeMoney & " 元资金兑换成 " & GetValidDay & " 天有效期！", ComeUrl)
        Else
            Call AddRechargeLog("System", UserName, GetValidDay, 1, 1, "将 " & ChangeExp & " 分积分兑换成 " & GetValidDay & " 天有效期")
            Call WriteSuccessMsg("成功将 " & ChangeExp & " 分积分兑换成 " & GetValidDay & " 天有效期！", ComeUrl)
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
        ErrMsg = ErrMsg & "<li>请输入充值卡卡号！</li>"
    End If
    If Password = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入充值卡密码！</li>"
    Else
        Password = Base64encode(Password)
    End If
    If FoundErr = True Then Exit Sub
    
    Set rsCard = Server.CreateObject("Adodb.Recordset")
    rsCard.Open "select * from PE_Card where CardNum='" & CardNum & "' and Password='" & Password & "'", Conn, 1, 3
    If rsCard.BOF And rsCard.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>卡号或密码错误！</li>"
    Else
        If rsCard("CardType") <> 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>你输入的充值卡是其他公司的卡，不能在本站进行充值。请尽快去有关公司或网站的充值入口进行充值。</li>"
        End If
        If rsCard("UserName") <> "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>你输入的充值卡已经使用过了！</li>"
        End If
        If rsCard("EndDate") < Date Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>你输入的充值卡已经失效！此卡的充值截止日期为：" & rsCard("EndDate")
        End If
    End If
    If FoundErr = True Then
        rsCard.Close
        Set rsCard = Nothing
        Exit Sub
    End If
    

    Dim strMsg
    strMsg = "充值成功！"
    If rsCard("ValidUnit") = 5 Then
        strMsg = strMsg & "&nbsp;&nbsp;&nbsp;<font color='red'>恭喜您已升级成 “" & GetValidNum(rsCard("ValidNum"), rsCard("ValidUnit")) & "”</font>"
    End If
    strMsg = strMsg & "<br><br>充值卡卡号：" & rsCard("CardNum") & "<br>"
    strMsg = strMsg & "充值卡面值：" & rsCard("Money") & "元" & "<br>"
    If rsCard("ValidUnit") = 5 Then
        strMsg = strMsg & "会员级别："
    Else
        strMsg = strMsg & "充值卡点数："
    End If
        strMsg = strMsg & GetValidNum(rsCard("ValidNum"), rsCard("ValidUnit")) & arrCardUnit(rsCard("ValidUnit")) & "<br>"
    strMsg = strMsg & "充值截止日期：" & rsCard("EndDate") & "<br><br>"
    
    Dim rsUser, sqlUser
    Set rsUser = Server.CreateObject("Adodb.RecordSet")
    sqlUser = "select * from PE_User where UserID=" & UserID
    rsUser.Open sqlUser, Conn, 1, 3
    Select Case rsCard("ValidUnit")
    Case 0    '点数
        strMsg = strMsg & "您充值前的" & PointName & "数：" & rsUser("UserPoint") & "<br>"
        rsUser("UserPoint") = rsUser("UserPoint") + rsCard("ValidNum")
        rsUser.Update
        strMsg = strMsg & "您充值后的" & PointName & "数：" & rsUser("UserPoint") & "<br>"
        Call AddConsumeLog("System", 0, UserName, 0, rsCard("ValidNum"), 1, "充值卡充值。卡号：" & rsCard("CardNum") & "")
    Case 4    '元
        strMsg = strMsg & "您充值前的资金余额为： " & rsUser("Balance") & " 元<br>"
        rsUser("Balance") = rsUser("Balance") + rsCard("ValidNum")
        rsUser.Update
        strMsg = strMsg & "您充值后的资金余额为： " & rsUser("Balance") & " 元<br>"
        
        Call AddBankrollItem("System", UserName, ClientID, rsCard("ValidNum"), 4, "", 0, 1, 0, 0, "充值卡充值。卡号：" & rsCard("CardNum") & "", Now())

    Case 5  '会员组
        Conn.Execute ("Update PE_User Set GroupID = " & rsCard("ValidNum") & " where UserName='" & UserName & "'")

    Case Else    '有效期
        If rsUser("ValidNum") = -1 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>您的有效期为“无限期”，无需充值。"
        Else
            If ValidDays > 0 Then
                strMsg = strMsg & "您充值前的有效期：" & rsUser("ValidNum") & arrCardUnit(rsUser("ValidUnit")) & "<br>"
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

                    Call AddRechargeLog("System", UserName, 0, 0, 0, "充值卡充值时更改有效期计费单位。卡号：" & rsCard("CardNum") & "")
                End If
                strMsg = strMsg & "您充值后的有效期：" & rsUser("ValidNum") & arrCardUnit(rsUser("ValidUnit")) & "<br>"
            Else
                strMsg = strMsg & "您充值前有效期已经过期 " & Abs(ValidDays) & " 天<br>"
                rsUser("BeginTime") = Now()
                rsUser("ValidNum") = rsCard("ValidNum")
                rsUser("ValidUnit") = rsCard("ValidUnit")
                rsUser.Update
                strMsg = strMsg & "您充值后的有效期：" & rsUser("ValidNum") & arrCardUnit(rsUser("ValidUnit")) & "，开始计算日期：" & Date & "<br>"
                Call AddRechargeLog("System", UserName, 0, 0, 0, "充值卡充值时将原来过期的有效期重新计算。卡号：" & rsCard("CardNum") & "")
            End If
            Call AddRechargeLog("System", UserName, rsCard("ValidNum"), rsCard("ValidUnit"), 1, "充值卡充值。卡号：" & rsCard("CardNum") & "")
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
        ErrMsg = ErrMsg & "<li>不允许将" & PointName & "赠送给他人！</li>"
        Exit Sub
    End If
    Dim SendObject, SendPoint, i, j
    Dim arrSendObject
    Dim rsUser, rsObject
    
    SendObject = Trim(Request("SendObject"))
    SendPoint = PE_CLng(Trim(Request("SendPoint")))
    If SendObject = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入对方的用户名！</li>"
    Else
        If CheckBadChar(SendObject) = False Then
            ErrMsg = ErrMsg + "<li>用户名中含有非法字符</li>"
            FoundErr = True
        End If
    End If
    If SendPoint <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您还未输入" & PointName & "数或输入的" & PointName & "数中存在非法字符！</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    j = 0
    arrSendObject = Split(SendObject, ",")
    Set rsUser = Conn.Execute("select * from PE_User where UserID=" & UserID & "")
    If rsUser("UserPoint") - SendPoint * (UBound(arrSendObject) + 1) < 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您的" & PointName & "不够！</li>"
        Exit Sub
    Else
        For i = 0 To UBound(arrSendObject)
            Set rsObject = Conn.Execute("select UserID from PE_User where UserName='" & arrSendObject(i) & "'")
            If Not rsObject.EOF Then
                Conn.Execute "Update PE_User set UserPoint=UserPoint + " & SendPoint & "  where UserName='" & arrSendObject(i) & "'"
                Conn.Execute "Update PE_User set UserPoint=UserPoint - " & SendPoint & " where UserID=" & UserID
                Call AddConsumeLog("System", 0, UserName, 0, SendPoint, 2, "向" & arrSendObject(i) & "用户赠送" & PointName & "")
                Call AddConsumeLog("System", 0, arrSendObject(i), 0, SendPoint, 1, "获得" & UserName & "用户赠送的" & PointName & "")
                Conn.Execute "Insert into PE_Message (Incept,Sender,Title,IsSend,Content,Flag) values('" & arrSendObject(i) & "','" & UserName & "','获赠" & PointName & "',1,'" & UserName & "赠给您" & PointName & "" & SendPoint & PointUnit & "',0)"
            Else
                j = j + 1
            End If
            Set rsObject = Nothing
        Next
        If j = 0 Then
           Call WriteSuccessMsg(PointName & "赠送成功！", ComeUrl)
        Else
           Call WriteSuccessMsg("对" & UBound(arrSendObject) - j + 1 & "位用户赠送成功！其中有" & j & "位用户不存在！", ComeUrl)
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
'函数名：ValidNumToValidDays
'作  用：转换有效期为有效天数
'参  数：iValidNum ----有效期
'        iValidUnit ----有效期单位
'        iBeginTime ---- 开始计算日期
'返回值：有效天数
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
