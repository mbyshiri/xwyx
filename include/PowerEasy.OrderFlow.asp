<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************


Dim Phone, InvoiceContent, Remark, InvoiceInfo, BeginDate

Dim ProductList, ProductID
Dim PresentList, PresentID
Dim ShowTips_Login, ShowTips_CartIsEmpty
Dim rsCart, CartID, IsWholesale
Dim dblPrice, dblTruePrice, dblTempPrice, dblAmount, dblSubtotal, dblTotal, dblTotal2
Dim strProductType, strSaleType, strDiscount
Dim TotalExp, TotalMoney, TotalPoint
Dim strJS
Dim PaymentType, PaymentTypeName, Discount_Payment, rsPaymentType
Dim DeliverType, DeliverTypeName, Charge_Deliver, rsDeliverType
Dim ReleaseType, MinMoney1, ReleaseCharge, MinMoney2, MaxCharge, MinMoney3
Dim Charge_Min, Weight_Min, ChargePerUnit, WeightPerUnit, Charge_Max
Dim Charge_Percent
Dim TotalWeight
Dim Present2, Present3, Present4, Cash
Dim NeedInvoice

Dim ContacterName, Company, Department, OfficePhone, HomePhone, Fax, Mobile, Address, ZipCode, AgentName

ChannelShortName = "产品"
MaxPerPage = 20
HitsOfHot = 1000
XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))

UserLogined = CheckUserLogined()

If UserLogined = False Then
    ShowTips_Login = "<tr><td colspan='10'>" & XmlText_Class("ShowTips_Login", "<font color='#0000FF'>温馨提示：</font><font color='#006600'>您还没有注册或登录，请先<a href='../Reg/User_Reg.asp'>注册</a>或<a href='../User/User_Login.asp'>登录</a>，以获得更多优惠！</font>") & "</td></tr>"
Else
    Call GetUser(UserName)
    ShowTips_Login = ""
End If

strPageTitle = SiteTitle

Call GetChannel(ChannelID)

If Trim(ChannelName) <> "" And ShowChannelName <> False Then
    If UseCreateHTML > 0 Then
        strNavPath = strNavPath & "&nbsp;&gt;&gt;&nbsp;<a href='" & ChannelUrl & "/Index" & FileExt_Index & "'>" & ChannelName & "</a>"
    Else
        strNavPath = strNavPath & "&nbsp;&gt;&gt;&nbsp;<a href='" & ChannelUrl & "/Index.asp'>" & ChannelName & "</a>"
    End If
    strPageTitle = strPageTitle & " >> " & ChannelName
End If


Sub ReplaceCommon()
    Call ReplaceCommonLabel
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)
    strHtml = PE_Replace(strHtml, "{$MenuJS}", GetMenuJS(ChannelDir, ShowClassTreeGuide))
    strHtml = PE_Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))
End Sub


Sub ReplaceUserInfo()
    strHtml = PE_Replace(strHtml, "{$UserName}", UserName & "")
    strHtml = PE_Replace(strHtml, "{$GroupName}", GroupName & "")
    strHtml = PE_Replace(strHtml, "{$Discount_Member}", Discount_Member & "")
    strHtml = PE_Replace(strHtml, "{$IsOffer}", IsOffer & "")
    strHtml = PE_Replace(strHtml, "{$Balance}", Balance & "")
    strHtml = PE_Replace(strHtml, "{$UserPoint}", UserPoint & "")
    strHtml = PE_Replace(strHtml, "{$UserExp}", UserExp & "")
    strHtml = PE_Replace(strHtml, "{$Email}", Email & "")
End Sub

Function ShowCart()
    If IsNull(ProductList) Or ProductList = "" Then
        ShowCart = ""
        Exit Function
    End If
    If IsValidID(ProductList) = False Then
        ShowCart = "<li>ProductList数据非法！</li>"
        Exit Function
    End If
    
    Dim strCart
    strCart = strCart & "              <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border3'>" & vbCrLf
    strCart = strCart & "                <tr align='center' class='tdbg2' height='25'>" & vbCrLf
    strCart = strCart & "                  <td width='32'><b>购买</b></td>" & vbCrLf
    strCart = strCart & "                  <td><b>商 品 名 称</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='44'><b>单 位</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='54'><b>数 量</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='64'><b>商品类别</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='64'><b>销售类型</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='64'><b>原价</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='44'><b>折扣</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='64'><b>实价</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='84'><b>金 额</b></td>" & vbCrLf
    strCart = strCart & "                </tr>" & vbCrLf
    
    dblSubtotal = 0
    dblTotal = 0
    TotalExp = 0
    TotalMoney = 0
    TotalPoint = 0
    
    Set rsCart = Conn.Execute("select P.*,C.Quantity,C.CartItemID from PE_Product P inner join PE_ShoppingCarts C on C.ProductID=P.ProductID where C.CartID='" & CartID & "' and C.PresentID=0 and P.ProductType<>4 order by C.CartItemID desc")
    Do While Not rsCart.EOF
        dblAmount = rsCart("Quantity")
        If dblAmount <= 0 Then dblAmount = 1
        If rsCart("ProductType") = 3 And rsCart("LimitNum") > 0 And dblAmount > rsCart("LimitNum") Then
            strJS = strJS & "你订购了" & dblAmount & rsCart("Unit") & rsCart("ProductName") & "，而此商品每人最多限购" & rsCart("LimitNum") & rsCart("Unit") & "，所以系统已自动将您的订购数量更改为" & rsCart("LimitNum") & rsCart("Unit") & "！\n"
            dblAmount = rsCart("LimitNum")
        End If
        If rsCart("Stocks") - rsCart("OrderNum") < dblAmount Then
            strJS = strJS & "你订购了" & dblAmount & rsCart("Unit") & rsCart("ProductName") & "，而此商品目前只有" & rsCart("Stocks") - rsCart("OrderNum") & rsCart("Unit") & "库存！\n您可以修改订购数或者继续按原数量订购，我们将尽快配齐此商品。\n"
        End If
        If dblAmount <> rsCart("Quantity") Then
            Call UpdateAmount(rsCart("CartItemID"), dblAmount)
        End If
        
        dblPrice = rsCart("Price_Original")
    
        If PE_CLng(UserSetting(30)) = 1 And rsCart("EnableWholesale") = True And dblAmount >= rsCart("Number_Wholesale1") Then
            strProductType = "批发"
            strSaleType = "批发"
            strDiscount = "─"
            If dblAmount < rsCart("Number_Wholesale2") Then
                dblTruePrice = rsCart("Price_Wholesale1")
            Else
                If dblAmount < rsCart("Number_Wholesale3") Then
                    dblTruePrice = rsCart("Price_Wholesale2")
                Else
                    dblTruePrice = rsCart("Price_Wholesale3")
                End If
            End If
        Else
            Select Case GroupType
            Case 0, 1 '未登录
                Select Case rsCart("ProductType")
                Case 1
                    dblTruePrice = rsCart("Price")
                    strProductType = "正常销售"
                    strDiscount = "─"
                Case 2
                    strProductType = "涨价商品"
                    dblTruePrice = rsCart("Price")
                    strDiscount = "─"
                Case 3
                    If Date < rsCart("BeginDate") Or Date > rsCart("EndDate") Then
                        dblTruePrice = dblPrice
                        strProductType = "正常销售"
                        strDiscount = "─"
                    Else
                        dblTruePrice = rsCart("Price")
                        strProductType = "特价商品"
                        strDiscount = rsCart("Discount") & "折"
                    End If
                Case 5
                    strProductType = "降价商品"
                    dblTruePrice = rsCart("Price")
                    strDiscount = rsCart("Discount") & "折"
                End Select
            Case 2, 3   '注册会员
                Select Case rsCart("ProductType")
                Case 1
                    strProductType = "正常销售"
                    strDiscount = "─"
                    If rsCart("Price_Member") > 0 Then '如果指定了会员价
                        dblTruePrice = rsCart("Price_Member")
                    Else
                        dblTruePrice = rsCart("Price") * Discount_Member / 100
                    End If
                Case 2
                    strProductType = "涨价商品"
                    strDiscount = "─"
                    If rsCart("Price_Member") > 0 Then '如果指定了会员价
                        dblTruePrice = rsCart("Price_Member")
                    Else
                        dblTruePrice = rsCart("Price") * Discount_Member / 100
                    End If
                Case 3, 5
                    If rsCart("ProductType") = 3 Then
                        If Date < rsCart("BeginDate") Or Date > rsCart("EndDate") Then
                            strProductType = "正常销售"
                            strDiscount = "─"
                            dblTempPrice = dblPrice
                        Else
                            strProductType = "特价商品"
                            strDiscount = rsCart("Discount") & "折"
                            dblTempPrice = rsCart("Price")
                        End If
                    Else
                        strProductType = "降价商品"
                        strDiscount = rsCart("Discount") & "折"
                        dblTempPrice = rsCart("Price")
                    End If
                    If rsCart("Price_Member") > 0 Then '如果指定了会员价
                        If rsCart("Price_Member") <= dblTempPrice Then
                            dblTruePrice = rsCart("Price_Member")
                        Else
                            dblTruePrice = dblTempPrice
                        End If
                    Else
                        If PE_CLng(UserSetting(12)) = 1 Then '如可以享受折上折优惠
                            dblTruePrice = dblTempPrice * Discount_Member / 100
                        Else
                            If dblPrice * Discount_Member / 100 >= dblTempPrice Then
                                dblTruePrice = dblTempPrice
                            Else
                                dblTruePrice = dblPrice * Discount_Member / 100
                            End If
                        End If
                    End If
                End Select
            Case 4  '代理商
                Select Case rsCart("ProductType")
                Case 1
                    strProductType = "正常销售"
                    strDiscount = "─"
                Case 2
                    strProductType = "涨价商品"
                    strDiscount = "─"
                Case 3
                    If Date < rsCart("BeginDate") Or Date > rsCart("EndDate") Then
                        strProductType = "正常销售"
                    Else
                        strProductType = "特价商品"
                        strDiscount = rsCart("Discount") & "折"
                    End If
                Case 5
                    strProductType = "降价商品"
                    strDiscount = rsCart("Discount") & "折"
                End Select
                dblTempPrice = rsCart("Price")
                If rsCart("Price_Agent") > 0 Then '如果指定了代理价
                    dblTruePrice = rsCart("Price_Agent")
                    strDiscount = "─"
                Else
                    If Discount_Member = 100 Then
                        dblTruePrice = dblTempPrice
                    Else
                        If PE_CLng(UserSetting(12)) = 1 Then '如可以享受折上折优惠
                            dblTruePrice = dblTempPrice * Discount_Member / 100
                        Else
                            If rsCart("Price_Original") * Discount_Member / 100 <= dblTempPrice Then
                                dblTruePrice = rsCart("Price_Original") * Discount_Member / 100
                                strDiscount = Round(Discount_Member / 10, 1) & "折"
                            Else
                                dblTruePrice = dblTempPrice
                            End If
                        End If
                    End If
                End If
            End Select
            strSaleType = "零售"
        End If
        dblSubtotal = dblTruePrice * dblAmount
        dblTotal = dblTotal + dblSubtotal
        TotalExp = TotalExp + rsCart("PresentExp") * dblAmount
        TotalMoney = TotalMoney + rsCart("PresentMoney") * dblAmount
        TotalPoint = TotalPoint + rsCart("PresentPoint") * dblAmount

        strCart = strCart & "                <tr valign='middle' class='tdbg3' height='20'>" & vbCrLf
        strCart = strCart & "                  <td width='32' align='center'><input type='CheckBox' name='ProductID' value='" & rsCart("ProductID") & "' Checked></td>" & vbCrLf
        strCart = strCart & "                  <td align='left'>" & rsCart("ProductName") & "</td>" & vbCrLf
        strCart = strCart & "                  <td width='44' align='center'>" & rsCart("Unit") & "</td>" & vbCrLf
        strCart = strCart & "                  <td width='54' align='center'><input name='Amount_" & rsCart("ProductID") & "' type='Text' value='" & dblAmount & "' size='5' maxlength='10' style='text-align: center;'></td>" & vbCrLf
        strCart = strCart & "                  <td width='64' align='center'>" & strProductType & "</td>" & vbCrLf
        strCart = strCart & "                  <td width='64' align='center'>" & strSaleType & "</td>" & vbCrLf
        strCart = strCart & "                  <td width='64' align='right'>" & FormatNumber(dblPrice, 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
        strCart = strCart & "                  <td width='44' align=center>" & strDiscount & "</td>" & vbCrLf
        strCart = strCart & "                  <td width='64' align='right'>" & FormatNumber(dblTruePrice, 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
        strCart = strCart & "                  <td width='84' align='right'>" & FormatNumber(dblSubtotal, 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
        strCart = strCart & "                </tr>" & vbCrLf

        If rsCart("SalePromotionType") > 0 And dblAmount >= rsCart("MinNumber") Then
            Dim PresentNumber
            If rsCart("SalePromotionType") = 3 Or rsCart("SalePromotionType") = 4 Then
                PresentNumber = rsCart("PresentNumber")
            Else
                PresentNumber = Fix(dblAmount / rsCart("MinNumber")) * rsCart("PresentNumber")
            End If
            If rsCart("SalePromotionType") = 1 Or rsCart("SalePromotionType") = 3 Then
                strCart = strCart & "                <tr valign='middle' class='tdbg3' height='20'>" & vbCrLf
                strCart = strCart & "                  <td width='32' align='center'><input type='CheckBox' name='PresentID' value='" & rsCart("ProductID") & "'"
                If FoundInArr(PresentList, rsCart("ProductID"), ",") = True Then strCart = strCart & " checked"
                strCart = strCart & "></td>" & vbCrLf
                strCart = strCart & "                  <td align='left'>" & rsCart("ProductName") & " <font color='red'>（赠送）</font></td>" & vbCrLf
                strCart = strCart & "                  <td width='44' align='center'>" & rsCart("Unit") & "</td>" & vbCrLf
                strCart = strCart & "                  <td width='54' align='center'>" & PresentNumber & "</td>" & vbCrLf
                strCart = strCart & "                  <td width='64' align='center'><font color='red'>赠送礼品</font></td>" & vbCrLf
                strCart = strCart & "                  <td width='64' align='center'><font color='red'>赠送</font></td>" & vbCrLf
                strCart = strCart & "                  <td width='64' align='right'>" & FormatNumber(dblPrice, 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
                strCart = strCart & "                  <td width='44' align='center'>─</td>" & vbCrLf
                strCart = strCart & "                  <td width='64' align='right'>0.00</td>" & vbCrLf
                strCart = strCart & "                  <td width='84' align='right'>0.00</td>" & vbCrLf
                strCart = strCart & "                </tr>" & vbCrLf
            Else
                Dim rsPresent, strPresentType
                Set rsPresent = Conn.Execute("select * from PE_Product where ProductNum='" & rsCart("PresentID") & "' and ProductType=4")
                If Not (rsPresent.BOF And rsPresent.EOF) Then
                    If rsPresent("Price") > 0 Then
                        strPresentType = "换购"
                    Else
                        strPresentType = "赠送"
                    End If
                    strCart = strCart & "                <tr valign='middle' class='tdbg3' height='20'>" & vbCrLf
                    strCart = strCart & "                  <td width='32' align='center'><input type='CheckBox' name='PresentID' value='" & rsPresent("ProductID") & "'"
                    If FoundInArr(PresentList, rsPresent("ProductID"), ",") = True Then strCart = strCart & " checked"
                    strCart = strCart & "></td>" & vbCrLf
                    strCart = strCart & "                  <td align='left'>" & rsPresent("ProductName") & " <font color='red'>（" & strPresentType & "）</font></td>" & vbCrLf
                    strCart = strCart & "                  <td width='44' align='center'>" & rsPresent("Unit") & "</td>" & vbCrLf
                    strCart = strCart & "                  <td width='54' align='center'>" & PresentNumber & "</td>" & vbCrLf
                    strCart = strCart & "                  <td width='64' align='center'><font color='red'>促销礼品</font></td>" & vbCrLf
                    strCart = strCart & "                  <td width='64' align='center'><font color='red'>" & strPresentType & "</font></td>" & vbCrLf
                    strCart = strCart & "                  <td width='64' align='right'>" & FormatNumber(rsPresent("Price_Original"), 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
                    strCart = strCart & "                  <td width='44' align=center>─</td>" & vbCrLf
                    strCart = strCart & "                  <td width='64' align='right'>" & FormatNumber(rsPresent("Price"), 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
                    If FoundInArr(PresentList, rsPresent("ProductID"), ",") = False Then
                        PresentNumber = 0
                    End If
                    strCart = strCart & "                  <td width='84' align='right'>" & FormatNumber(rsPresent("Price") * PresentNumber, 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
                    strCart = strCart & "                </tr>" & vbCrLf
                    dblTotal = dblTotal + rsPresent("Price") * PresentNumber
                End If
                Set rsPresent = Nothing
            End If
        End If
        rsCart.MoveNext
    Loop
    rsCart.Close
    Set rsCart = Nothing

    strCart = strCart & "                <tr class='tdbg3'>" & vbCrLf
    strCart = strCart & "                  <td colspan='9' align='right'><b>合计：</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='80' align=right><b> ￥" & FormatNumber(dblTotal, 2, vbTrue, vbFalse, vbTrue) & "</b></td>" & vbCrLf
    strCart = strCart & "                </tr>" & vbCrLf
    strCart = strCart & "              </table>" & vbCrLf
    If strJS <> "" Then
        strCart = strCart & "<script language='javascript'>alert('" & strJS & "');</script>"
    End If
    ShowCart = strCart
End Function

Function ShowCart2(IsPreview)
    If IsNull(ProductList) Or ProductList = "" Then
        ShowCart2 = ""
        Exit Function
    End If
    Dim strCart
    strCart = strCart & "              <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border3'>" & vbCrLf
    strCart = strCart & "                <tr align='center' class='tdbg2' height='25'>" & vbCrLf
    strCart = strCart & "                  <td><b>商 品 名 称</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='44'><b>单 位</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='54'><b>数 量</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='64'><b>商品类别</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='64'><b>销售类型</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='64'><b>原价</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='44'><b>折扣</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='64'><b>实价</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='84'><b>金 额</b></td>" & vbCrLf
    strCart = strCart & "                </tr>" & vbCrLf
    
    dblSubtotal = 0
    dblTotal = 0
    TotalExp = 0
    TotalWeight = 0
    
    Set rsCart = Conn.Execute("select P.*,C.Quantity,C.CartItemID from PE_Product P inner join PE_ShoppingCarts C on C.ProductID=P.ProductID where C.CartID='" & CartID & "' and C.PresentID=0 and P.ProductType<>4 order by C.CartItemID desc")
    Do While Not rsCart.EOF
        dblAmount = rsCart("Quantity")
        If dblAmount <= 0 Then dblAmount = 1
        If rsCart("ProductType") = 3 And rsCart("LimitNum") > 0 And dblAmount > rsCart("LimitNum") Then
            dblAmount = rsCart("LimitNum")
        End If

        If dblAmount <> rsCart("Quantity") Then
            Call UpdateAmount(rsCart("CartItemID"), dblAmount)
        End If
        
        dblPrice = rsCart("Price_Original")
        TotalWeight = TotalWeight + PE_CDbl(rsCart("Weight")) * dblAmount
    
        If PE_CLng(UserSetting(30)) = 1 And rsCart("EnableWholesale") = True And dblAmount >= rsCart("Number_Wholesale1") Then
            strProductType = "批发"
            strSaleType = "批发"
            strDiscount = "─"
            If dblAmount < rsCart("Number_Wholesale2") Then
                dblTruePrice = rsCart("Price_Wholesale1")
            Else
                If dblAmount < rsCart("Number_Wholesale3") Then
                    dblTruePrice = rsCart("Price_Wholesale2")
                Else
                    dblTruePrice = rsCart("Price_Wholesale3")
                End If
            End If
        Else
            Select Case GroupType
            Case 0, 1 '未登录
                Select Case rsCart("ProductType")
                Case 1
                    dblTruePrice = rsCart("Price")
                    strProductType = "正常销售"
                    strDiscount = "─"
                Case 2
                    strProductType = "涨价商品"
                    dblTruePrice = rsCart("Price")
                    strDiscount = "─"
                Case 3
                    If Date < rsCart("BeginDate") Or Date > rsCart("EndDate") Then
                        dblTruePrice = dblPrice
                        strProductType = "正常销售"
                        strDiscount = "─"
                    Else
                        dblTruePrice = rsCart("Price")
                        strProductType = "特价商品"
                        strDiscount = rsCart("Discount") & "折"
                    End If
                Case 5
                    strProductType = "降价商品"
                    dblTruePrice = rsCart("Price")
                    strDiscount = rsCart("Discount") & "折"
                End Select
            Case 2, 3   '注册会员
                Select Case rsCart("ProductType")
                Case 1
                    strProductType = "正常销售"
                    strDiscount = "─"
                    If rsCart("Price_Member") > 0 Then '如果指定了会员价
                        dblTruePrice = rsCart("Price_Member")
                    Else
                        dblTruePrice = rsCart("Price") * Discount_Member / 100
                    End If
                Case 2
                    strProductType = "涨价商品"
                    strDiscount = "─"
                    If rsCart("Price_Member") > 0 Then '如果指定了会员价
                        dblTruePrice = rsCart("Price_Member")
                    Else
                        dblTruePrice = rsCart("Price") * Discount_Member / 100
                    End If
                Case 3, 5
                    If rsCart("ProductType") = 3 Then
                        If Date < rsCart("BeginDate") Or Date > rsCart("EndDate") Then
                            strProductType = "正常销售"
                            strDiscount = "─"
                            dblTempPrice = dblPrice
                        Else
                            strProductType = "特价商品"
                            strDiscount = rsCart("Discount") & "折"
                            dblTempPrice = rsCart("Price")
                        End If
                    Else
                        strProductType = "降价商品"
                        strDiscount = rsCart("Discount") & "折"
                        dblTempPrice = rsCart("Price")
                    End If
                    If rsCart("Price_Member") > 0 Then '如果指定了会员价
                        If rsCart("Price_Member") <= dblTempPrice Then
                            dblTruePrice = rsCart("Price_Member")
                        Else
                            dblTruePrice = dblTempPrice
                        End If
                    Else
                        If PE_CLng(UserSetting(12)) = 1 Then '如可以享受折上折优惠
                            dblTruePrice = dblTempPrice * Discount_Member / 100
                        Else
                            If dblPrice * Discount_Member / 100 >= dblTempPrice Then
                                dblTruePrice = dblTempPrice
                            Else
                                dblTruePrice = dblPrice * Discount_Member / 100
                            End If
                        End If
                    End If
                End Select
            Case 4  '代理商
                Select Case rsCart("ProductType")
                Case 1
                    strProductType = "正常销售"
                    strDiscount = "─"
                Case 2
                    strProductType = "涨价商品"
                    strDiscount = "─"
                Case 3
                    If Date < rsCart("BeginDate") Or Date > rsCart("EndDate") Then
                        strProductType = "正常销售"
                    Else
                        strProductType = "特价商品"
                        strDiscount = rsCart("Discount") & "折"
                    End If
                Case 5
                    strProductType = "降价商品"
                    strDiscount = rsCart("Discount") & "折"
                End Select
                dblTempPrice = rsCart("Price")
                If rsCart("Price_Agent") > 0 Then '如果指定了代理价
                    dblTruePrice = rsCart("Price_Agent")
                    strDiscount = "─"
                Else
                    If Discount_Member = 100 Then
                        dblTruePrice = dblTempPrice
                    Else
                        If PE_CLng(UserSetting(12)) = 1 Then '如可以享受折上折优惠
                            dblTruePrice = dblTempPrice * Discount_Member / 100
                        Else
                            If rsCart("Price_Original") * Discount_Member / 100 <= dblTempPrice Then
                                dblTruePrice = rsCart("Price_Original") * Discount_Member / 100
                                strDiscount = Round(Discount_Member / 10, 1) & "折"
                            Else
                                dblTruePrice = dblTempPrice
                            End If
                        End If
                    End If
                End If
            End Select
            strSaleType = "零售"
        End If
        dblSubtotal = dblTruePrice * dblAmount
        dblTotal = dblTotal + dblSubtotal
        TotalExp = TotalExp + rsCart("PresentExp") * dblAmount
        TotalPoint = TotalPoint + rsCart("PresentPoint") * dblAmount
        TotalMoney = TotalMoney + rsCart("PresentMoney") * dblAmount

        strCart = strCart & "                <tr valign='middle' class='tdbg3' height='20'>" & vbCrLf
        strCart = strCart & "                  <td align='left'>" & rsCart("ProductName") & "</td>" & vbCrLf
        strCart = strCart & "                  <td width='44' align='center'>" & rsCart("Unit") & "</td>" & vbCrLf
        strCart = strCart & "                  <td width='54' align='center'>" & dblAmount & "</td>" & vbCrLf
        strCart = strCart & "                  <td width='64' align='center'>" & strProductType & "</td>" & vbCrLf
        strCart = strCart & "                  <td width='64' align='center'>" & strSaleType & "</td>" & vbCrLf
        strCart = strCart & "                  <td width='64' align='right'>" & FormatNumber(dblPrice, 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
        strCart = strCart & "                  <td width='44' align=center>" & strDiscount & "</td>" & vbCrLf
        strCart = strCart & "                  <td width='64' align='right'>" & FormatNumber(dblTruePrice, 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
        strCart = strCart & "                  <td width='84' align='right'>" & FormatNumber(dblSubtotal, 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
        strCart = strCart & "                </tr>" & vbCrLf


        If rsCart("SalePromotionType") > 0 And dblAmount >= rsCart("MinNumber") Then
            Dim PresentNumber
            If rsCart("SalePromotionType") = 3 Or rsCart("SalePromotionType") = 4 Then
                PresentNumber = rsCart("PresentNumber")
            Else
                PresentNumber = Fix(dblAmount / rsCart("MinNumber")) * rsCart("PresentNumber")
            End If
            If rsCart("SalePromotionType") = 1 Or rsCart("SalePromotionType") = 3 Then
                If FoundInArr(PresentList, rsCart("ProductID"), ",") = True Then
                    strCart = strCart & "                <tr valign='middle' class='tdbg3' height='20'>" & vbCrLf
                    strCart = strCart & "                  <td align='left'>" & rsCart("ProductName") & " <font color='red'>（赠送）</font></td>" & vbCrLf
                    strCart = strCart & "                  <td width='44' align='center'>" & rsCart("Unit") & "</td>" & vbCrLf
                    strCart = strCart & "                  <td width='54' align='center'>" & PresentNumber & "</td>" & vbCrLf
                    strCart = strCart & "                  <td width='64' align='center'><font color='red'>赠送礼品</font></td>" & vbCrLf
                    strCart = strCart & "                  <td width='64' align='center'><font color='red'>赠送</font></td>" & vbCrLf
                    strCart = strCart & "                  <td width='64' align='right'>" & FormatNumber(dblPrice, 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
                    strCart = strCart & "                  <td width='44' align='center'>─</td>" & vbCrLf
                    strCart = strCart & "                  <td width='64' align='right'>0.00</td>" & vbCrLf
                    strCart = strCart & "                  <td width='84' align='right'>0.00</td>" & vbCrLf
                    strCart = strCart & "                </tr>" & vbCrLf
                    TotalWeight = TotalWeight + PE_CDbl(rsCart("Weight")) * PresentNumber
                End If
            Else
                Dim rsPresent, strPresentType
                Set rsPresent = Conn.Execute("select * from PE_Product where ProductNum='" & rsCart("PresentID") & "' and ProductType=4")
                If Not (rsPresent.BOF And rsPresent.EOF) Then
                    If FoundInArr(PresentList, rsPresent("ProductID"), ",") = True Then
                        If rsPresent("Price") > 0 Then
                            strPresentType = "换购"
                        Else
                            strPresentType = "赠送"
                        End If
                        strCart = strCart & "                <tr valign='middle' class='tdbg3' height='20'>" & vbCrLf
                        strCart = strCart & "                  <td align='left'>" & rsPresent("ProductName") & " <font color='red'>（" & strPresentType & "）</font></td>" & vbCrLf
                        strCart = strCart & "                  <td width='44' align='center'>" & rsPresent("Unit") & "</td>" & vbCrLf
                        strCart = strCart & "                  <td width='54' align='center'>" & PresentNumber & "</td>" & vbCrLf
                        strCart = strCart & "                  <td width='64' align='center'><font color='red'>促销礼品</font></td>" & vbCrLf
                        strCart = strCart & "                  <td width='64' align='center'><font color='red'>" & strPresentType & "</font></td>" & vbCrLf
                        strCart = strCart & "                  <td width='64' align='right'>" & FormatNumber(rsPresent("Price_Original"), 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
                        strCart = strCart & "                  <td width='44' align=center>─</td>" & vbCrLf
                        strCart = strCart & "                  <td width='64' align='right'>" & FormatNumber(rsPresent("Price"), 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
                        strCart = strCart & "                  <td width='84' align='right'>" & FormatNumber(rsPresent("Price") * PresentNumber, 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
                        strCart = strCart & "                </tr>" & vbCrLf
                        dblTotal = dblTotal + rsPresent("Price") * PresentNumber
                        TotalWeight = TotalWeight + PE_CDbl(rsPresent("Weight")) * PresentNumber
                    End If
                End If
                Set rsPresent = Nothing
            End If
        End If
        rsCart.MoveNext
    Loop
    rsCart.Close
    Set rsCart = Nothing
    
    If IsPreview = True Then
        strCart = strCart & ShowPresent2()
    End If
    
    strCart = strCart & "                <tr class='tdbg3'>" & vbCrLf
    strCart = strCart & "                  <td colspan='8' align='right'><b>合计：</b></td>" & vbCrLf
    strCart = strCart & "                  <td width='80' align=right><b> ￥" & FormatNumber(dblTotal, 2, vbTrue, vbFalse, vbTrue) & "</b></td>" & vbCrLf
    strCart = strCart & "                </tr>" & vbCrLf
    
    If IsPreview = True Then
        strCart = strCart & ShowPresent3()
    Else
        strCart = strCart & ShowPresent()
    End If
    
    strCart = strCart & "              </table>" & vbCrLf
    ShowCart2 = strCart
End Function

Function ShowPresent()
    Dim strPresent, rsPresent
    Set rsPresent = Conn.Execute("select * from PE_PresentProject where MinMoney<=" & dblTotal & " and MaxMoney>" & dblTotal & " and BeginDate<=" & PE_Now & " and EndDate>=" & PE_Now & "")
    If Not (rsPresent.BOF And rsPresent.EOF) Then
        If FoundInArr(rsPresent("PresentContent"), "1", ",") Then
            strPresent = strPresent & "              <tr class='tdbg3' height='25'>" & vbCrLf
            strPresent = strPresent & "                <td colspan='20'><b>你可以用 <font color='red'>" & rsPresent("Price") & "</font> 元超值换购以下商品中的任一款：</b></td>" & vbCrLf
            strPresent = strPresent & "              </tr>" & vbCrLf
            Set rsCart = Conn.Execute("select * from PE_Product where ProductID in (" & rsPresent("PresentID") & ") and ProductType=4")
            Do While Not rsCart.EOF
                strPresent = strPresent & "              <tr valign='middle' class='tdbg3'>"
                strPresent = strPresent & "                <td height='20'><input type='radio' name='PresentID2' value='" & rsCart("ProductID") & "'"
                If FoundInArr(PresentList, rsCart("ProductID"), ",") = True Or rsPresent("Price") <= 0 Then strPresent = strPresent & " checked"
                strPresent = strPresent & ">" & rsCart("ProductName") & " <font color='red'>（超值换购）</font></td>"
                strPresent = strPresent & "                <td width='44' align='center'>" & rsCart("Unit") & "</td>"
                strPresent = strPresent & "                <td width='54' align='center'>1</td>"
                strPresent = strPresent & "                <td width='64' align='center'><font color='red'>促销礼品</font></td>"
                strPresent = strPresent & "                <td width='64' align='center'><font color='red'>超值换购</font></td>"
                strPresent = strPresent & "                <td width='64' align='right'>" & FormatNumber(rsCart("Price_Original"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
                strPresent = strPresent & "                <td width='44' align=center>─</td>"
                strPresent = strPresent & "                <td width='64' align='right'>" & FormatNumber(rsPresent("Price"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
                strPresent = strPresent & "                <td width='84' align='right'>" & FormatNumber(rsPresent("Price"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
                strPresent = strPresent & "              </tr>"
                rsCart.MoveNext
            Loop
            Set rsCart = Nothing
        End If
        Cash = rsPresent("Cash")
        PresentPoint = rsPresent("PresentPoint")
        PresentExp = rsPresent("PresentExp")
        
        Dim Present2, Present3, Present4
        Present2 = FoundInArr(rsPresent("PresentContent"), "2", ",")
        Present3 = FoundInArr(rsPresent("PresentContent"), "3", ",")
        Present4 = FoundInArr(rsPresent("PresentContent"), "4", ",")
    End If

    If (Present2 = True Or Present3 = True Or Present4 = True Or TotalExp > 0 Or TotalMoney > 0 Or TotalPoint > 0) And CheckUserLogined = True Then
        strPresent = strPresent & "<tr class='tdbg3'><td colspan='20'><b>另外，你还可以得到 "
        If Present2 = True Or TotalMoney > 0 Then
            strPresent = strPresent & "<font color='red'>" & Cash + TotalMoney & "</font> 元现金券"
        End If
        If Present3 = True Or TotalExp > 0 Then
            If Present2 = True Or TotalMoney > 0 Then
                strPresent = strPresent & " 和 "
            End If
            strPresent = strPresent & "<font color='red'>" & PresentExp + TotalExp & "</font> 点积分"&""
        End If
        If Present4 = True Or TotalPoint > 0 Then
            If Present2 = True Or Present3 = True Or TotalMoney > 0 Or TotalExp > 0 Then
                strPresent = strPresent & " 和 "
            End If
            strPresent = strPresent & "<font color='red'>" & PresentPoint + TotalPoint & "</font> " & PointUnit & PointName
        End If
        strPresent = strPresent & "</b></td></tr>"
    End If
    Set rsPresent = Nothing
    ShowPresent = strPresent
End Function

Function ShowPresent2()
    Dim PresentID2
    Dim strPresent, rsPresent
    PresentID2 = PE_CLng(Trim(Request("PresentID2")))
    
    Set rsPresent = Conn.Execute("select * from PE_PresentProject where MinMoney<=" & dblTotal & " and MaxMoney>" & dblTotal & " and MaxMoney>" & dblTotal & " and BeginDate<=" & PE_Now & " and EndDate>=" & PE_Now & "")
    If Not (rsPresent.BOF And rsPresent.EOF) Then
        If FoundInArr(rsPresent("PresentContent"), "1", ",") And PresentID2 > 0 Then
            Set rsCart = Conn.Execute("select * from PE_Product where ProductID=" & PresentID2 & " and ProductType=4")
            If Not (rsCart.BOF And rsCart.EOF) Then
                strPresent = strPresent & "              <tr valign='middle' class='tdbg3'>"
                strPresent = strPresent & "                <td height='20' align='left'><input type='hidden' name='PresentID2' value='" & rsCart("ProductID") & "'>" & rsCart("ProductName") & " <font color='red'>（超值换购）</font></td>"
                strPresent = strPresent & "                <td width='44' align='center'>" & rsCart("Unit") & "</td>"
                strPresent = strPresent & "                <td width='54' align='center'>1</td>"
                strPresent = strPresent & "                <td width='64' align='center'><font color='red'>促销礼品</font></td>"
                strPresent = strPresent & "                <td width='64' align='center'><font color='red'>超值换购</font></td>"
                strPresent = strPresent & "                <td width='64' align='right'>" & FormatNumber(rsCart("Price_Original"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
                strPresent = strPresent & "                <td width='44' align=center>─</td>"
                strPresent = strPresent & "                <td width='64' align='right'>" & FormatNumber(rsPresent("Price"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
                strPresent = strPresent & "                <td width='84' align='right'>" & FormatNumber(rsPresent("Price"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
                strPresent = strPresent & "              </tr>"
                dblTotal = dblTotal + rsPresent("Price")
                TotalWeight = TotalWeight + PE_CDbl(rsCart("Weight"))
            End If
            Set rsCart = Nothing
        End If
        Present2 = FoundInArr(rsPresent("PresentContent"), "2", ",")
        Present3 = FoundInArr(rsPresent("PresentContent"), "3", ",")
        Present4 = FoundInArr(rsPresent("PresentContent"), "4", ",")
        Cash = rsPresent("Cash")
        PresentExp = rsPresent("PresentExp")
        PresentPoint = rsPresent("PresentPoint")
    End If
    Set rsPresent = Nothing
    ShowPresent2 = strPresent
End Function

Function ShowPresent3()
    
    Dim strPresent, strTotalMoney, dblTotal_Original, Charge_Deliver_Original
    dblTotal_Original = dblTotal
    strPresent = strPresent & "            <tr class='tdbg3'>" & vbCrLf
    strPresent = strPresent & "              <td colspan='6'>付款方式折扣率：" & Discount_Payment & "%"
    strTotalMoney = "实际金额：(" & dblTotal & "×" & Discount_Payment & "%"
    If Discount_Payment > 0 And Discount_Payment < 100 Then
        dblTotal = dblTotal * Discount_Payment / 100
    End If
    
    Select Case ChargeType
    Case 0
        Charge_Deliver = 0
    Case 1
        If TotalWeight > 0 Then
            If TotalWeight > Weight_Min Then
                Dim iTemp
                iTemp = (TotalWeight - Weight_Min) / WeightPerUnit
                If iTemp > Fix(iTemp) Then
                    iTemp = Fix(iTemp) + 1
                End If
                Charge_Deliver = Charge_Min + iTemp * ChargePerUnit
                If Charge_Deliver > Charge_Max Then
                    Charge_Deliver = Charge_Max
                End If
            Else
                Charge_Deliver = Charge_Min
            End If
        Else
            Charge_Deliver = 0
        End If
    Case 2
        Charge_Deliver = dblTotal_Original * Charge_Percent / 100 + Charge_Min
        If Charge_Deliver > Charge_Max Then
            Charge_Deliver = Charge_Max
        End If
    End Select

    If Charge_Deliver > 0 And ReleaseType > 0 And dblTotal_Original >= MinMoney1 Then
        Charge_Deliver_Original = Charge_Deliver
        If Charge_Deliver <= ReleaseCharge Then
            Charge_Deliver = 0
        Else
            Charge_Deliver = Charge_Deliver - ReleaseCharge
        End If
        If dblTotal_Original >= MinMoney2 Then
            If dblTotal_Original >= MinMoney3 Then
                Charge_Deliver = 0
            Else
                If Charge_Deliver_Original <= MaxCharge Then
                    Charge_Deliver = 0
                End If
            End If
        End If
    End If

    strPresent = strPresent & "&nbsp;&nbsp;&nbsp;&nbsp;运费：" & Charge_Deliver & " 元"
    strTotalMoney = strTotalMoney & "＋" & Charge_Deliver & ")"
    dblTotal = dblTotal + Charge_Deliver
    
    strPresent = strPresent & "&nbsp;&nbsp;&nbsp;&nbsp;税率：" & TaxRate & "%&nbsp;&nbsp;&nbsp;&nbsp;价格含税："
    If IncludeTax = True Then
        strPresent = strPresent & "是"
        If NeedInvoice <> "Yes" Then
            strTotalMoney = strTotalMoney & "×(1-" & TaxRate & "%)"
            dblTotal = dblTotal * (100 - TaxRate) / 100
        Else
            strTotalMoney = strTotalMoney & "×100%"
        End If
    Else
        strPresent = strPresent & "否"
        If NeedInvoice = "Yes" Then
            strTotalMoney = strTotalMoney & "×(1+" & TaxRate & "%)"
            dblTotal = dblTotal * (100 + TaxRate) / 100
        Else
            strTotalMoney = strTotalMoney & "×100%"
        End If
    End If
    strTotalMoney = strTotalMoney & "＝" & FormatNumber(dblTotal, 2, vbTrue, vbFalse, vbTrue) & " 元"
    strPresent = strPresent & "<br>" & strTotalMoney
    
    strPresent = strPresent & "              </td>" & vbCrLf
    strPresent = strPresent & "              <td colspan='2' align='right'><b>实际金额：</b></td>" & vbCrLf
    strPresent = strPresent & "              <td width='80' align=right><b> ￥" & FormatNumber(dblTotal, 2, vbTrue, vbFalse, vbTrue) & "</b></td>" & vbCrLf
        
    If (Present2 = True Or Present3 = True Or Present4 = True Or TotalExp > 0 Or TotalMoney > 0 Or TotalPoint > 0) And CheckUserLogined = True Then
        strPresent = strPresent & "<tr class='tdbg3'><td colspan='20'><b>另外，你还可以得到 "
        If Present2 = True Or TotalMoney > 0 Then
            strPresent = strPresent & "<font color='red'>" & Cash + TotalMoney & "</font> 元现金券"
        End If
        If Present3 = True Or TotalExp > 0 Then
            If Present2 = True Or TotalMoney > 0 Then
                strPresent = strPresent & " 和 "
            End If
            strPresent = strPresent & "<font color='red'>" & PresentExp + TotalExp & "</font> 点积分"
        End If
        If Present4 = True Or TotalPoint > 0 Then
            If Present2 = True Or Present3 = True Or TotalMoney > 0 Or TotalExp > 0 Then
                strPresent = strPresent & " 和 "
            End If
            strPresent = strPresent & "<font color='red'>" &PresentPoint  + TotalPoint & "</font> " & PointUnit & PointName
        End If
        strPresent = strPresent & "</b></td></tr>"
    End If
    ShowPresent3 = strPresent
End Function

Sub AddToCart(CartID, iProductID, Quantity, IsPresent)
    Dim CartItemID, IsExistential
    If IsPresent = 0 Then
        Conn.Execute ("update PE_ShoppingCarts set Quantity=Quantity+0,UserName='" & UserName & "', UpdateTime=" & PE_Now & " where CartID='" & CartID & "' and ProductID=" & iProductID & ""), IsExistential
        If IsExistential = 0 Then
            Conn.Execute ("insert into PE_ShoppingCarts (CartID,ProductID,Quantity,PresentID,UserName,UpdateTime) values ('" & CartID & "'," & iProductID & "," & Quantity & "," & IsPresent & ",'" & UserName & "'," & PE_Now & ")")
        End If
    Else
        Conn.Execute ("insert into PE_ShoppingCarts (CartID,ProductID,Quantity,PresentID,UserName,UpdateTime) values ('" & CartID & "'," & iProductID & "," & Quantity & "," & IsPresent & ",'" & UserName & "'," & PE_Now & ")")
    End If
End Sub

Function SelectCart(CartID, IsPresent)
    Dim rsCartItem, ProductList
    ProductList = ""
    If IsPresent = 1 Then
        Set rsCartItem = Conn.Execute("select ProductID from PE_ShoppingCarts where CartID='" & CartID & "' and PresentID=1 order by CartItemID desc")
    Else
        Set rsCartItem = Conn.Execute("select ProductID from PE_ShoppingCarts where CartID='" & CartID & "' and PresentID=0 order by CartItemID desc")
    End If
    Do While Not rsCartItem.EOF
        If ProductList = "" Then
            ProductList = rsCartItem("ProductID")
        Else
            ProductList = ProductList & "," & rsCartItem("ProductID")
        End If
        rsCartItem.MoveNext
    Loop
    SelectCart = ProductList
    Set rsCartItem = Nothing
End Function

Function GetAmount(CartID, ProductID)
    Dim rsAmount
    Set rsAmount = Conn.Execute("select Quantity from PE_ShoppingCarts where CartID='" & CartID & "' and ProductID=" & ProductID & "")
    If Not (rsAmount.BOF And rsAmount.EOF) Then
        GetAmount = rsAmount(0)
    End If
End Function

Sub UpdateAmount(CartItemID, Amount)
    Conn.Execute ("Update PE_ShoppingCarts set Quantity=" & Amount & " where CartItemID=" & CartItemID & "")
End Sub

Sub DelCart(CartID)
    Conn.Execute ("delete from PE_ShoppingCarts where CartID='" & CartID & "'")
End Sub

Sub ReplaceCommon()
    Call ReplaceCommonLabel
    
    strHtml = PE_Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = PE_Replace(strHtml, "{$ShowPath}", ShowPath())
    strHtml = PE_Replace(strHtml, "{$MenuJS}", GetMenuJS(ChannelDir, ShowClassTreeGuide))
    strHtml = PE_Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(DefaultSkinID))
End Sub

Function XmlText_Class(ByVal iSmallNode, ByVal DefChar)
    XmlText_Class = XmlText("Product", iSmallNode, DefChar)
End Function

Function R_XmlText_Class(ByVal iSmallNode, ByVal DefChar)
    R_XmlText_Class = Replace(XmlText("Product", iSmallNode, DefChar), "{$ChannelShortName}", ChannelShortName)
End Function

%>
