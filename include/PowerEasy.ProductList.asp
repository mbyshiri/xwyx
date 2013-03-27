<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim PriClassID, ClassField(5)
Function GetProductList(ByVal str1, UseType)
    Dim strtmp, rsProduct, sqlProduct, i
    Dim strTemp, arrTemp, TempUrl
    Dim iType, iDate, iLink, iNum, iorder, iHeight, iWidth
    If str1 = "" Then
        GetProductList = ""
        Exit Function
    End If

    arrTemp = Split(str1, ",")
    If UBound(arrTemp) < 8 Then
        GetProductList = "函数式标签：{$ShowProductList()}的参数个数不对。请检查模板中的此标签。"
        Exit Function
    Else
        If UBound(arrTemp) = 10 Then
            iWidth = PE_CLng(arrTemp(9))
            iHeight = PE_CLng(arrTemp(10))
        Else
            iWidth = 130
            iHeight = 90
        End If
    End If
    iType = PE_CLng(arrTemp(4))
    iNum = PE_CLng(arrTemp(7))
    iorder = PE_CLng(arrTemp(8))
    
    strtmp = "<Table width='100%'><tr><td>"
    
    If UseType = 1 Then
        sqlProduct = "select ProductID,ClassID,ProductName,ProductIntro,Keyword,UpdateTime,ProductThumb,Price_Market,Price,Price_Member from PE_Product where ProducerName like '" & arrTemp(0) & "' and ChannelID=" & arrTemp(1) & " and EnableSale=" & PE_True & " and Deleted=" & PE_False & " and Stocks>0"
    Else
        sqlProduct = "select ProductID,ClassID,ProductName,ProductIntro,Keyword,UpdateTime,ProductThumb,Price_Market,Price,Price_Member from PE_Product where TrademarkName like '" & arrTemp(0) & "' and ChannelID=" & arrTemp(1) & " and EnableSale=" & PE_True & " and Deleted=" & PE_False & " and Stocks>0"
    End If
    If TimeData <> "0" Then
        sqlProduct = sqlProduct & " and DateDiff(" & PE_DatePart_D & ",UpdateTime,'" & TimeData & "')=0"
    End If
    If iorder = 1 Then
        sqlProduct = sqlProduct & " order by UpdateTime"
    Else
        sqlProduct = sqlProduct & " order by UpdateTime desc"
    End If
    
    Set rsProduct = Server.CreateObject("ADODB.Recordset")
    rsProduct.Open sqlProduct, Conn, 1, 1
    If rsProduct.BOF And rsProduct.EOF Then
        totalPut = 0
        If TimeData <> "0" Then
            strtmp = strtmp & Replace(XmlText("ShowSource", "ShowProducer/NoDay", "&nbsp;&nbsp;<h3>{$ProducerName}本日没有发布商品</h3>"), "{$ProducerName}", arrTemp(0))
        Else
            strtmp = strtmp & Replace(XmlText("ShowSource", "ShowProducer/NoFound", "&nbsp;&nbsp;<h3>{$ProducerName}目前没有发布商品</h3>"), "{$ProducerName}", arrTemp(0))
        End If
        
    Else
        totalPut = rsProduct.RecordCount
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                    rsProduct.Move (CurrentPage - 1) * MaxPerPage
                Else
                    CurrentPage = 1
                End If
        End If
        i = 0
        Dim strShowDetal
        strShowDetal = XmlText("ShowSource", "ShowProducer/ShowDetal", "点击这里浏览具体内容>>>")

        Do While Not rsProduct.EOF
        TempUrl = GetProductUrl(GetClassFild(rsProduct("ClassID"), 4), GetClassFild(rsProduct("ClassID"), 3), rsProduct("UpdateTime"), rsProduct("ProductID"))
            Select Case iType
            Case 1
                If PE_CBool(arrTemp(5)) Then strtmp = strtmp & ("<p>" & Year(rsProduct("UpdateTime")) & strYear & Month(rsProduct("UpdateTime")) & strMonth & Day(rsProduct("UpdateTime")) & strDay & "</p>")
                strtmp = strtmp & ("<h4>" & ReplaceText(rsProduct("ProductName"), 2) & "</h4>&nbsp;&nbsp;")
                    If rsProduct("ProductIntro") = "" Then
                        strtmp = strtmp & rsProduct("Keyword")
                    Else
                        strtmp = strtmp & GetSubStr(ReplaceText(nohtml(rsProduct("ProductIntro")), 1), iNum, True)
                    End If
                    If arrTemp(6) = "True" Then
                        strtmp = strtmp & ("<div align='right'><a href='" & TempUrl)
                        strtmp = strtmp & ("'>" & strShowDetal & "</a></div><hr>")
                    End If
            Case 2
                If PE_CBool(arrTemp(6)) Then
                    strtmp = strtmp & ("<li><a href='" & TempUrl)
                    strtmp = strtmp & ("' Target=""_blank"">" & GetSubStr(ReplaceText(rsProduct("ProductName"), 2), iNum, False) & "</a>")
                Else
                    strtmp = strtmp & ("<li>" & GetSubStr(ReplaceText(rsProduct("ProductName"), 2), iNum, False))
                End If
                If PE_CBool(arrTemp(5)) Then
                    strtmp = strtmp & ("[" & Year(rsProduct("UpdateTime")) & strYear & Month(rsProduct("UpdateTime")) & strMonth & Day(rsProduct("UpdateTime")) & strDay & "]")
                End If
                strtmp = strtmp & "</li>"
            Case 3
                strtmp = strtmp & "<table width='100%' cellspacing='2' border='0'>"
                strtmp = strtmp & "<tr><td align='center' class='productpic'>"
                strtmp = strtmp & ("<a href='" & TempUrl & "' title='" & ReplaceText(rsProduct("ProductName"), 2) & "' target='_blank'>")
                If rsProduct("ProductThumb") = "" Or IsNull(rsProduct("ProductThumb")) Then
                    strtmp = strtmp & "<img src='" & strInstallDir & "images/nopic.gif'  width='" & iWidth & "' height='" & iHeight & "' border='0'></a></td>"
                Else
                    If Left(rsProduct("ProductThumb"), 4) = "http" Or Left(rsProduct("ProductThumb"), 1) = "/" Then
                        strtmp = strtmp & "<img src='" & rsProduct("ProductThumb") & "'  width='" & iWidth & "' height='" & iHeight & "' border='0'></a></td>"
                    Else
                        strtmp = strtmp & "<img src='" & strInstallDir & arrTemp(2) & "/" & arrTemp(3) & "/" & rsProduct("ProductThumb") & "'  width='" & iWidth & "' height='" & iHeight & "' border='0'></a></td>"
                    End If
                End If
                strtmp = strtmp & "<td align='left'><a href='"
                strtmp = strtmp & TempUrl
                strtmp = strtmp & ("' title='" & ReplaceText(rsProduct("ProductName"), 2) & "' target='_blank'>" & GetSubStr(ReplaceText(rsProduct("ProductName"), 2), iNum, False) & "</a><br>市场价：<font class='price'><STRIKE>￥" & rsProduct("Price_Market") & "</STRIKE></font><br>商城价：<font class='price'>￥" & rsProduct("Price") & "</font><br>会员价：<font class='price'>￥" & rsProduct("Price_Member") & "</font></td>")
                strtmp = strtmp & "</tr><tr>"
                strtmp = strtmp & "<td align='center'><a href='"
                strtmp = strtmp & TempUrl
                strtmp = strtmp & ("' title='" & rsProduct("ProductName") & "' target='_blank'>" & GetSubStr(ReplaceText(rsProduct("ProductName"), 2), iNum, False) & "</a></td>")
                strtmp = strtmp & "<td align='left'><a href='" & strInstallDir & arrTemp(2) & "/ShoppingCart.asp?Action=Add&ProductID=" & rsProduct("ProductID") & "' target='ShoppingCart'><img src='" & strInstallDir & arrTemp(2) & "/images/ProductBuy.gif' border='0'></a>&nbsp;&nbsp;<a href='"
                strtmp = strtmp & TempUrl
                strtmp = strtmp & ("' title='" & rsProduct("ProductName") & "' target='_blank'><img src='" & strInstallDir & arrTemp(2) & "/images/ProductContent.gif' border='0'></a></td>")
                strtmp = strtmp & "</tr></table>"
            End Select
            rsProduct.MoveNext
            i = i + 1
            If i >= MaxPerPage Then Exit Do
        Loop
    End If
    rsProduct.Close
    Set rsProduct = Nothing
    strtmp = strtmp & "</td></tr></table>"
    GetProductList = strtmp
End Function
%>
