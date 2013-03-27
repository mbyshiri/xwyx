<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim HouseID, InfoType
Class House
Private MoreCorrelative

Public Sub Init()
    ClassID = PE_CLng(Trim(Request("ClassID")))
    SpecialID = PE_CLng(Trim(Request("SpecialID")))
    HouseID = PE_CLng(Request("HouseID"))
    ChannelShortName = "房产"
    
    
    '*****************************
    '读取语言包中的字符设置
    strListStr_Font = XmlText_Class("HouseList/UpdateTimeColor_New", "color=""red""")
    '*****************************
    
    strPageTitle = SiteTitle

    Call GetChannel(ChannelID)
    HtmlDir = InstallDir & ChannelDir
    If Trim(ChannelName) <> "" And ShowNameOnPath <> False Then
        strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;<a class='LinkPath' href='" & ChannelUrl & "/Index.asp'>" & ChannelName & "</a>"
        strPageTitle = strPageTitle & " >> " & ChannelName
    End If
    Call GetClass
End Sub


'=================================================
'函数名：ShowChannelCount
'作  用：显示频道统计信息
'参  数：无
'=================================================
Private Function GetChannelCount()
    Dim rs, Count_All, Count_CZ, Count_CS, Count_QG, Count_QZ, Count_HZ
    Set rs = Conn.Execute("select Count(0) from PE_HouseCZ")
    Count_CZ = rs(0)
    rs.Close
    Set rs = Nothing
    Set rs = Conn.Execute("select Count(0) from PE_HouseCS")
    Count_CS = rs(0)
    rs.Close
    Set rs = Nothing
    Set rs = Conn.Execute("select Count(0) from PE_HouseQG")
    Count_QG = rs(0)
    rs.Close
    Set rs = Nothing
    Set rs = Conn.Execute("select Count(0) from PE_HouseQZ")
    Count_QZ = rs(0)
    rs.Close
    Set rs = Nothing
    Set rs = Conn.Execute("select Count(0) from PE_HouseHZ")
    Count_HZ = rs(0)
    rs.Close
    Set rs = Nothing
    Count_All = Count_CZ + Count_CS + Count_QG + Count_QZ + Count_HZ

    GetChannelCount = Replace(Replace(Replace(Replace(Replace(Replace(R_XmlText_Class("ChannelCount", "信息总数： {$Count_All} 条<br>出租信息： {$Count_CZ} 条<br>出售信息： {$Count_CS} 条<br>求租信息： {$Count_QZ} 条<br>求购信息： {$Count_QG} 条<br>合租信息： {$Count_HZ} 条"), "{$Count_All}", Count_All), "{$Count_CZ}", Count_CZ), "{$Count_CS}", Count_CS), "{$Count_QZ}", Count_QZ), "{$Count_QG}", Count_QG), "{$Count_HZ}", Count_HZ)
End Function


Private Function GetAreaList()
    Dim sql, rsArealist, strAreaList, i
    Set rsArealist = Server.CreateObject("adodb.recordset")
    sql = "select * from PE_HouseArea"
    rsArealist.Open sql, Conn, 1, 3
    If rsArealist.EOF Then
        strAreaList = "<li>没有任何区域</li>"
    Else
        strAreaList = "<table width=100% ><tr>"
        Do While Not rsArealist.EOF
            strAreaList = strAreaList & "<td><a href='Search.asp?InfoType=" & ClassID & "&HouseQuYu=" & rsArealist("AreaName") & "'>" & rsArealist("AreaName") & "</td>"
            i = i + 1
            If (i Mod 10) = 0 Then
                strAreaList = strAreaList & "</tr><tr>"
            End If
            rsArealist.MoveNext
        Loop
        strAreaList = strAreaList & "</tr></table>"
    End If
    GetAreaList = strAreaList
End Function



Public Function GetHouseList(ClassID, HouseSource, Num, isHot, isElite, isContactName, isDiZhi, isHuXing, isLeiXing, isPrice, isMianJi, isUpdateTime, isSource)
    Dim sqlstr, rsHouseList, strHouseList, UserPage, strThisClass, OpenType, iCount, iYouXiaoQi, CssName
    Call GetClass
    strThisClass = ""
    iCount = 1
    CurrentPage = PE_CLng(Trim(Request("page")))
    If Num <> 0 Then
        sqlstr = "select Top " & Num & " "
    Else
        sqlstr = "select "
    End If
    Select Case PE_CLng(ClassID)
    Case 1
        sqlstr = sqlstr & "HouseID,HouseDiZhi,HouseSource,HouseHuXing,HouseLeiXing,CommendClassDays,YouXiaoQi,TotalPrice,HousePriceType,HouseMianJi,Elite,UpdateTime,Deleted from PE_HouseCS "
    Case 2
        sqlstr = sqlstr & "HouseID,HouseDiZhi,Elite,HouseSource,HouseHuXing,HouseLeiXing,CommendClassDays,YouXiaoQi,HouseMianJI,HouseZuJin,HouseZuJinType,HouseZhuangXiu,UpdateTime,Deleted from PE_HouseCZ "
    Case 3
        sqlstr = sqlstr & "HouseID,HouseDiZhi,Elite,HouseLeiXing,HouseHuXing,HouseSource,CommendClassDays,YouXiaoQi,HouseMianJi1,HouseMianJi2,HousePrice1,HousePrice2,HousePriceType,UpdateTime,ContactName,Deleted from PE_HouseQG "
    Case 4
        sqlstr = sqlstr & "HouseID,HouseDiZhi,Elite,HouseSource,HouseHuXing,HouseLeiXing,CommendClassDays,YouXiaoQi,HouseMianJi1,HouseMianJi2,UpdateTime,Elite,HouseZuJin1,HouseZuJin2,ContactName,HouseZuJinType,Deleted from PE_HouseQZ "
    Case 5
        sqlstr = sqlstr & "HouseID,HeZhuType,HouseDiZhi,HouseLeiXing,Elite,HouseHuXing,HouseSource,CommendClassDays,YouXiaoQi,HouseZuJin,HouseMianJi1,UpdateTime,HouseZuJinType,ContactName,Deleted from PE_HouseHZ "
    Case Else
        FoundErr = True
        GetHouseList = "<li>参数丢失！</li>"
        Exit Function
    End Select
    sqlstr = sqlstr & " Where Passed=" & PE_True & " and  Deleted=" & PE_False & ""
    Select Case PE_CLng(HouseSource)
    Case 1
        sqlstr = sqlstr & " and HouseSource='中介'"
    Case 2
        sqlstr = sqlstr & " and  HouseSource='个人'"
    End Select
    If isHot = True Then
        sqlstr = sqlstr & " and Hot=" & PE_True & ""
    End If
    If isElite = True Then
        sqlstr = sqlstr & " and Elite=" & PE_True & ""
    End If
    If ItemOpenType = 0 Then
        OpenType = "_self"
    Else
        OpenType = "_blank"
    End If
    Set rsHouseList = Server.CreateObject("ADODB.Recordset")
    sqlstr = sqlstr & " order by OnTop " & PE_OrderType & ",UpdateTime Desc"
    rsHouseList.Open sqlstr, Conn, 1, 3
    If rsHouseList.BOF And rsHouseList.EOF Then
        If UserPage = True Then totalPut = 0
        If isHot = False And isElite = False And MoreCorrelative = False Then
            strHouseList = "<li>" & strThisClass & XmlText_Class("HouseList/t1", "没有") & ChannelShortName & "</li>"
        ElseIf isHot = True And isElite = False Then
            strHouseList = "<li>" & strThisClass & XmlText_Class("HouseList/t1", "没有") & XmlText_Class("HouseList/t2", "热门") & ChannelShortName & "</li>"
        ElseIf isHot = False And isElite = True Then
            strHouseList = "<li>" & strThisClass & XmlText_Class("HouseList/t1", "没有") & XmlText_Class("HouseList/t3", "推荐") & ChannelShortName & "</li>"
        Else
            strHouseList = "<li>" & strThisClass & XmlText_Class("HouseList/t1", "没有") & XmlText_Class("HouseList/t2", "热门") & XmlText_Class("HouseList/t3", "推荐") & ChannelShortName & "</li>"
        End If
        rsHouseList.Close
        Set rsHouseList = Nothing
        GetHouseList = strHouseList
        Exit Function
    Else
        totalPut = rsHouseList.RecordCount
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > totalPut Then
            If (totalPut Mod MaxPerPage) = 0 Then
                CurrentPage = totalPut \ MaxPerPage
            Else
                CurrentPage = totalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                rsHouseList.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
    End If

    CssName = "houselistbg"
    strHouseList = "<table class='housetable' width=100% ><tr>"
    If isContactName = True Then
        strHouseList = strHouseList & "<td align='center' class='houseth'>联系人</td>"
    End If
    If ClassID = 5 Then
        strHouseList = strHouseList & "<td align='center' class='houseth'>合租类型</td>"
    End If
    If isDiZhi = True Then
        strHouseList = strHouseList & "<td align='center' class='houseth'>地址</td>"
    End If
    If isHuXing = True Then
        strHouseList = strHouseList & "<td align='center' class='houseth'>户型</td>"
    End If
    If isLeiXing = True Then
        strHouseList = strHouseList & "<td align='center' class='houseth'>类型</td>"
    End If
    If isPrice = True Then
        If ClassID = 1 Or ClassID = 3 Then
            strHouseList = strHouseList & "<td align='center' class='houseth'>价格</td>"
        Else
            strHouseList = strHouseList & "<td align='center' class='houseth'>租金</td>"
        End If
    End If
    If isMianJi = True Then
        strHouseList = strHouseList & "<td align='center' class='houseth'>面积</td>"
    End If
    If isUpdateTime = True Then
        strHouseList = strHouseList & "<td align='center' class='houseth'>发布时间</td>"
    End If
    If isSource = True Then
        strHouseList = strHouseList & "<td align='center' class='houseth'>来源</td>"
    End If
    strHouseList = strHouseList & "</tr>"

    Do While Not rsHouseList.EOF
        Select Case rsHouseList("YouXiaoQi")
        Case "一周"
            iYouXiaoQi = 7
        Case "半个月"
            iYouXiaoQi = 14
        Case "一个月"
            iYouXiaoQi = 30
        Case "两个月"
            iYouXiaoQi = 61
        Case "半年"
            iYouXiaoQi = 183
        Case "一年"
            iYouXiaoQi = 365
        End Select
        If DateDiff("D", rsHouseList("UpdateTime"), Now()) > iYouXiaoQi Then
            rsHouseList("Deleted") = PE_True
            rsHouseList.Update
        Else
            If isElite = True And DateDiff("D", rsHouseList("UpdateTime"), Now()) > rsHouseList("CommendClassDays") Then
                rsHouseList("CommendClassDays") = 0
                rsHouseList("Elite") = PE_False
                rsHouseList.Update
            Else
                strHouseList = strHouseList & "<tr>"
                If isContactName = True Then
                    strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>" & rsHouseList("ContactName") & "</td>"
                End If
                If ClassID = 5 Then
                    strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>" & rsHouseList("HeZhuType") & "</td>"
                End If
                If isDiZhi = True Then
                    strHouseList = strHouseList & "<td align='center' class='" & CssName & "'><a href='ShowHouse.asp?ClassID=" & ClassID & "&HouseID=" & rsHouseList("HouseID") & "' target='" & OpenType & "'>" & rsHouseList("HouseDiZhi") & "</a></td>"
                End If
                If isHuXing = True Then
                    strHouseList = strHouseList & "<td align='center' class='" & CssName & "'><a href='ShowHouse.asp?ClassID=" & ClassID & "&HouseID=" & rsHouseList("HouseID") & "' target='" & OpenType & "'>" & Replace(PE_HTMLEncode(rsHouseList("HouseHuXing")), ",", "") & "</a></td>"
                End If
                If isLeiXing = True Then
                    If rsHouseList("HouseLeiXing") = "" Then
                        strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>不详</td>"
                    Else
                        strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>" & rsHouseList("HouseLeiXing") & "</td>"
                    End If
                End If
                If isPrice = True Then
                    Select Case PE_CLng(ClassID)
                    Case 1
                        strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>" & rsHouseList("TotalPrice") & rsHouseList("HousePriceType") & "</td>"
                    Case 2
                        strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>"
                        If rsHouseList("HouseZuJinType") = "价格面议" Then
                            strHouseList = strHouseList & rsHouseList("HouseZuJinType") & "</td>"
                        Else
                            strHouseList = strHouseList & rsHouseList("HouseZuJin") & "" & rsHouseList("HouseZuJinType") & "</td>"
                        End If
                    Case 3
                        strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>" & rsHouseList("HousePrice1") & "-" & rsHouseList("HousePrice2") & "" & rsHouseList("HousePriceType") & "</td>"
                    Case 4
                        strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>"
                        If rsHouseList("HouseZuJinType") = "价格面议" Then
                            strHouseList = strHouseList & rsHouseList("HouseZuJinType") & "</td>"
                        Else
                            strHouseList = strHouseList & rsHouseList("HouseZuJin1") & "-" & rsHouseList("HouseZuJin2") & "" & rsHouseList("HouseZuJinType") & "</td>"
                        End If
                    Case 5
                        strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>"
                        If rsHouseList("HouseZuJinType") = "价格面议" Then
                            strHouseList = strHouseList & rsHouseList("HouseZuJinType") & "</td>"
                        Else
                            strHouseList = strHouseList & rsHouseList("HouseZuJin") & "" & rsHouseList("HouseZuJinType") & "</td>"
                        End If
                    End Select
                End If
                If isMianJi = True Then
                    Select Case PE_CLng(ClassID)
                    Case 1
                        strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>" & rsHouseList("HouseMianJi") & "㎡</td>"
                    Case 2
                        strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>" & rsHouseList("HouseMianJi") & "㎡</td>"
                    Case 3
                        strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>" & rsHouseList("HouseMianJi1") & "-" & rsHouseList("HouseMianJi2") & "㎡</td>"
                    Case 4
                        strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>" & rsHouseList("HouseMianJi1") & "-" & rsHouseList("HouseMianJi2") & "㎡</td>"
                    Case 5
                        strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>" & rsHouseList("HouseMianJi1") & "㎡</td>"
                    End Select
                End If
                If isUpdateTime = True Then
                    strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>" & FormatDateTime(rsHouseList("UpdateTime"), 2) & "</td>"
                End If
                If isSource = True Then
                    strHouseList = strHouseList & "<td align='center' class='" & CssName & "'>" & rsHouseList("HouseSource") & "</td>"
                End If
                strHouseList = strHouseList & "</tr>"
                iCount = iCount + 1
                If iCount Mod 2 = 0 Then
                    CssName = "houselistbg2"
                Else
                    CssName = "houselistbg"
                End If
            End If
        End If
        If iCount >= MaxPerPage Then Exit Do
        rsHouseList.MoveNext
    Loop
    If Num <> 0 And CurrentPage < 2 Then
        If isElite = False And isHot = False Then strHouseList = strHouseList & "<tr><td align='right' colSpan=8 class='" & CssName & "'><a href='ShowClass.asp?ClassID=" & ClassID & "' target='_blank'>更多>>&nbsp; &nbsp;  </a></td></tr>"
        If isElite = True Then strHouseList = strHouseList & "<tr><td align='right' colSpan=8 class='" & CssName & "'><a href='ShowElite.asp?ClassID=" & ClassID & "' target='_blank'>更多>>&nbsp; &nbsp;  </a></td></tr>"
        If isHot = True Then strHouseList = strHouseList & "<tr><td align='right' colSpan=8 class='" & CssName & "'><a href='ShowHot.asp?ClassID=" & ClassID & "' target='_blank'>更多>>&nbsp; &nbsp;  </a></td></tr>"
    End If
    strHouseList = strHouseList & "</table>"
    rsHouseList.Close
    Set rsHouseList = Nothing
    GetHouseList = strHouseList
End Function

Public Function GetListFromTemplate(ByVal strValue)
    Dim strList
    strList = strValue
    regEx.Pattern = "\{\$GetHouseList\((.*?)\)\}"
    Set Matches = regEx.Execute(strList)
    For Each Match In Matches
        strList = PE_Replace(strList, Match.value, GetListFromLabel(Match.SubMatches(0)))
    Next
    GetListFromTemplate = strList
End Function

Private Function GetListFromLabel(ByVal str1)
    Dim strTemp, arrTemp
    Dim tChannelID, HouseNum, arrClassID, tSpecialID, AuthorName, OrderType, OpenType, ClassID
    If str1 = "" Then
        GetListFromLabel = ""
        Exit Function
    End If
    
    strTemp = Replace(str1, Chr(34), "")
    arrTemp = Split(strTemp, ",")
    If UBound(arrTemp) <> 12 Then
        GetListFromLabel = "函数式标签：{$GetHouseList(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Exit Function
    End If
    If Trim(arrTemp(0)) = "ClassID" Then
        ClassID = Request("ClassID")
    Else
        ClassID = arrTemp(0)
    End If
    GetListFromLabel = GetHouseList(PE_CLng(ClassID), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CBool(ReplaceBadChar(arrTemp(3))), PE_CBool(ReplaceBadChar(arrTemp(4))), PE_CBool(ReplaceBadChar(arrTemp(5))), PE_CBool(ReplaceBadChar(arrTemp(6))), PE_CBool(ReplaceBadChar(arrTemp(7))), PE_CBool(ReplaceBadChar(arrTemp(8))), PE_CBool(ReplaceBadChar(arrTemp(9))), PE_CBool(ReplaceBadChar(arrTemp(10))), PE_CBool(ReplaceBadChar(arrTemp(11))), PE_CBool(ReplaceBadChar(arrTemp(12))))
End Function

'=================================================
'函数名：GetSearchResult
'作  用：分页显示搜索结果
'参  数：无
'=================================================
Function GetSearchResult()
    Dim sqlSearch, ItemOpenType, Intro, rsSearch, iCount, HouseNum, HouseSource
    Dim arrHouseID, OpenType, strSearchResult, PriceUnit, Content, TableName, TimeBound, Price1, Price2, HouseHuXing1, HouseHuXing2, HouseQuYu, HouseMianJi, Address
    strSearchResult = ""
    HouseSource = ReplaceBadChar(Trim(Request("HouseSource")))
    InfoType = PE_CLng(Trim(Request("InfoType")))
    HouseHuXing1 = ReplaceBadChar(Trim(Request("HouseHuxing1")))
    HouseHuXing2 = ReplaceBadChar(Trim(Request("HouseHuxing2")))
    HouseQuYu = ReplaceBadChar(Trim(Request("HouseQuYu")))
    HouseMianJi = ReplaceBadChar(Trim(Request("HouseMianJi")))
    TimeBound = PE_CLng(Trim(Request("TimeBound")))
    Price1 = PE_CLng(Trim(Request("Price1")))
    Price2 = PE_CLng(Trim(Request("Price2")))
    Address = ReplaceBadChar(Trim(Request("Address")))
    Select Case Trim(Request("PriceUnit"))
    Case "元/月"
        PriceUnit = "元/月"
    Case "元/周"
        PriceUnit = "元/周"
    Case "元/年"
        PriceUnit = "元/年"
    Case "元/天"
        PriceUnit = "元/天"
    Case "元/季"
        PriceUnit = "元/季"
    Case "元/㎡"
        PriceUnit = "元/㎡"
    Case "万元"
        PriceUnit = "万元"
    End Select
    Select Case InfoType
    Case 1
        TableName = "PE_HouseCS"
    Case 2
        TableName = "PE_HouseCZ"
    Case 3
        TableName = "PE_HouseQG"
    Case 4
        TableName = "PE_HouseQZ"
    Case 5
        TableName = "PE_HouseHZ"
    Case Else
        TableName = "PE_HouseCS"
    End Select
    If PE_CLng(SearchResultNum) > 0 Then
        sqlSearch = "select top " & PE_CLng(SearchResultNum) & " HouseID"
    Else
        sqlSearch = "select HouseID"
    End If
    sqlSearch = sqlSearch & " from " & TableName & " where 1=1"
    If HouseQuYu <> "" Then
        sqlSearch = sqlSearch & " and HouseQuYu like '%" & HouseQuYu & "%'"
    End If
    If HouseSource = "1" Then
        sqlSearch = sqlSearch & " and HouseSource='个人'"
    Else
        If HouseSource = "2" Then
            sqlSearch = sqlSearch & " and HouseSource='中介'"
        End If
    End If
    If Address <> "" Then sqlSearch = sqlSearch & " and HouseDiZhi like '%" & Address & "%'"
    If HouseHuXing1 <> "" And HouseHuXing2 = "" Then
        sqlSearch = sqlSearch & " and HouseHuXing like '%" & HouseHuXing1 & "房%'"
    End If
    If HouseHuXing1 <> "" And HouseHuXing2 <> "" Then
        sqlSearch = sqlSearch & " and HouseHuXing like '%" & HouseHuXing1 & "房," & HouseHuXing2 & "厅%'"
    End If
    If HouseMianJi <> "" Then
        Select Case HouseMianJi
        Case 1
            Select Case InfoType
            Case 1, 2
                sqlSearch = sqlSearch & " and HouseMianJi<20"
            Case 5, 3, 4
                sqlSearch = sqlSearch & " and HouseMianJi1<20"
            End Select
        Case 2
            Select Case InfoType
            Case 1, 2
                sqlSearch = sqlSearch & " and HouseMianJi between 20 and 40"
            Case 3, 4
                sqlSearch = sqlSearch & " and HouseMianJi1>=20 and HouseMianJi1<=40 or HouseMianJi2>=20 and HouseMianJi2<=40"
            Case 5
                sqlSearch = sqlSearch & " and HouseMianJi1 between 20 and 40"
            End Select
        Case 3
            Select Case InfoType
            Case 1, 2
                sqlSearch = sqlSearch & " and HouseMianJi between 40 and 60"
            Case 3, 4
                sqlSearch = sqlSearch & " and HouseMianJi1>=40 and HouseMianJi1<=60 or HouseMianJi2>=40 and HouseMianJi2<=60"
            Case 5
                sqlSearch = sqlSearch & " and HouseMianJi1 between 40 and 60"
            End Select
        Case 4
            Select Case InfoType
            Case 1, 2
                sqlSearch = sqlSearch & " and HouseMianJi between 60 and 100"
            Case 3, 4
                sqlSearch = sqlSearch & " and HouseMianJi1>=60 and HouseMianJi1<=100 or HouseMianJi2>=60 and HouseMianJi2<=100"
            Case 5
                sqlSearch = sqlSearch & " and HouseMianJi1 between 60 and 100"
            End Select
        Case 5
            Select Case InfoType
            Case 1, 2
                sqlSearch = sqlSearch & " and HouseMianJi between 100 and 200"
            Case 3, 4
                sqlSearch = sqlSearch & " and HouseMianJi1>=100 and HouseMianJi1<=200 or HouseMianJi2>=100 and HouseMianJi2<=200"
            Case 5
                sqlSearch = sqlSearch & " and HouseMianJi1 between 100 and 200"
            End Select
        Case 6
            Select Case InfoType
            Case 1, 2
                sqlSearch = sqlSearch & " and HouseMianJi between 200 and 500"
            Case 3, 4
                sqlSearch = sqlSearch & " and HouseMianJi1>=200 and HouseMianJi1<=500 or HouseMianJi2>=200 and HouseMianJi2<=500"
            Case 5
                sqlSearch = sqlSearch & " and HouseMianJi1 between 200 and 500"
            End Select
        Case 7
            Select Case InfoType
            Case 1, 2
                sqlSearch = sqlSearch & " and HouseMianJi between 500 and 1000"
            Case 3, 4
                sqlSearch = sqlSearch & " and HouseMianJi1>=500 and HouseMianJi1<=1000 or HouseMianJi2>=500 and HouseMianJi2<=1000"
            Case 5
                sqlSearch = sqlSearch & " and HouseMianJi1 between 500 and 1000"
            End Select
        Case 8
            Select Case InfoType
            Case 1, 2
                sqlSearch = sqlSearch & " and HouseMianJi>1000"
            Case 3, 4
                sqlSearch = sqlSearch & " and HouseMianJi1>1000 or HouseMianJi2>1000"
            Case 5, 3, 4
                sqlSearch = sqlSearch & " and HouseMianJi1>1000"
            End Select
        End Select
    End If
    If Price1 <> 0 And Price2 = 0 Then
        Select Case InfoType
        Case 1
            sqlSearch = sqlSearch & " and TotalPrice=" & Price1 & " and HousePriceType='" & PriceUnit & "'"
        Case 2, 5
            sqlSearch = sqlSearch & " and HouseZuJin=" & Price1 & " and HouseZuJinType='" & PriceUnit & "'"
        Case 3
            sqlSearch = sqlSearch & " and HousePriceType='" & PriceUnit & "' and HousePrice1=" & Price1 & " or HousePrice2=" & Price1
        Case 4
            sqlSearch = sqlSearch & " and HouseZuJinType='" & PriceUnit & "' and HouseZuJin1=" & Price1 & " or HouseZuJin2=" & Price1 & ""
        End Select
    End If
    If Price1 = 0 And Price2 <> 0 Then
        Select Case InfoType
        Case 1
            sqlSearch = sqlSearch & " and TotalPrice=" & Price2 & " and HousePriceType='" & PriceUnit & "'"
        Case 2, 5
            sqlSearch = sqlSearch & " and HouseZuJin=" & Price2 & " and HouseZuJinType='" & PriceUnit & "'"
        Case 3
            sqlSearch = sqlSearch & " and HousePriceType='" & PriceUnit & "' and HousePrice1=" & Price2 & " or HousePrice2=" & Price2
        Case 4
            sqlSearch = sqlSearch & " and HouseZuJinType='" & PriceUnit & "' and HouseZuJin1=" & Price2 & " or HouseZuJin2=" & Price2
        End Select
    End If
    If Price1 <> 0 And Price2 <> 0 Then
        Select Case InfoType
        Case 1
            sqlSearch = sqlSearch & " and TotalPrice between " & Price1 & " and " & Price2 & " and HousePriceType='" & PriceUnit & "'"
        Case 2, 5
            sqlSearch = sqlSearch & " and HouseZuJin between " & Price1 & " and " & Price2 & " and HouseZuJinType='" & PriceUnit & "'"
        Case 3
            sqlSearch = sqlSearch & " and HousePriceType='" & PriceUnit & "' and HousePrice1>=" & Price1 & " and  HousePrice1<=" & Price2 & " or  HousePrice2>=" & Price1 & " and  HousePrice2<=" & Price2 & ""
        Case 4
            sqlSearch = sqlSearch & " and HouseZuJinType='" & PriceUnit & "' and HouseZuJin1>=" & Price1 & " and  HouseZuJin1<=" & Price2 & " or  HouseZuJin2>=" & Price1 & " and  HouseZuJin2<=" & Price2 & ""
        End Select
    End If
    Select Case TimeBound
    Case 1
        sqlSearch = sqlSearch & " and DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ")<1"
    Case 2
        sqlSearch = sqlSearch & " and DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ")<=3"
    Case 3
        sqlSearch = sqlSearch & " and DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ")<=7"
    Case 4
        sqlSearch = sqlSearch & " and DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ")<=15"
    Case 5
        sqlSearch = sqlSearch & " and DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ")<=30"
    Case 6
        sqlSearch = sqlSearch & " and DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ")<=60"
    Case 7
        sqlSearch = sqlSearch & " and DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ")<=90"
    End Select
    sqlSearch = sqlSearch & " and Deleted=" & PE_False & " and Passed=" & PE_True & " order by HouseID desc"
    Set rsSearch = Server.CreateObject("ADODB.Recordset")
    rsSearch.Open sqlSearch, Conn, 1, 1

    If rsSearch.BOF And rsSearch.EOF Then
        totalPut = 0
        strSearchResult = "<p align='center'><br><br>没有或没有找到任何" & ChannelShortName & "<br><br></p>"
    Else
        totalPut = rsSearch.RecordCount
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > totalPut Then
            If (totalPut Mod MaxPerPage) = 0 Then
                CurrentPage = totalPut \ MaxPerPage
            Else
                CurrentPage = totalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                rsSearch.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        HouseNum = 1
        arrHouseID = ""
        Do While Not rsSearch.EOF
            If arrHouseID = "" Then
                arrHouseID = rsSearch(0)
            Else
                arrHouseID = arrHouseID & "," & rsSearch(0)
            End If
            HouseNum = HouseNum + 1
            If HouseNum >= MaxPerPage Then Exit Do
            rsSearch.MoveNext
        Loop
    End If
    rsSearch.Close
    
    If arrHouseID = "" Then
        GetSearchResult = "<p align='center'><br><br>没有或没有找到任何" & ChannelShortName & "<br><br></p>"
        Set rsSearch = Nothing
        Exit Function
    End If
    Dim CssName
    sqlSearch = "select H.HouseID,H.HouseDiZhi,H.HouseHuXing,H.UpdateTime"
    Select Case InfoType
    Case 1
        sqlSearch = sqlSearch & ",H.HouseMianJi,H.TotalPrice,H.HousePriceType"
    Case 2
        sqlSearch = sqlSearch & ",H.HouseMianJi,H.HouseZuJin,H.HouseZuJinType"
    Case 3
        sqlSearch = sqlSearch & ",H.HouseMianJi1,H.HouseMianJi2,H.HousePriceType,H.HousePrice1,H.HousePrice2"
    Case 4
        sqlSearch = sqlSearch & ",H.HouseMianJi1,H.HouseMianJi2,H.HouseZuJinType,H.HouseZuJin1,H.HouseZuJin2,H.Deleted"
    Case 5
        sqlSearch = sqlSearch & ",H.HouseZuJin,H.HouseMianJi,H.HouseZuJinType,H.HouseMianJi1"
    Case Else
        sqlSearch = sqlSearch & ",H.HouseMianJi,H.TotalPrice,H.HousePriceType"
    End Select
    sqlSearch = sqlSearch & ",C.ClassID,C.ClassName,C.OpenType,C.ClassDir,C.ItemOpenType from " & TableName & " H left join PE_HouseConfig C on H.ClassID=C.ClassID where HouseID in (" & arrHouseID & ") order by HouseID desc"
    rsSearch.Open sqlSearch, Conn, 1, 1
    HouseNum = 1
    If rsSearch("ItemOpenType") = 0 Then
        OpenType = "_Self"
    Else
        OpenType = "_Blank"
    End If
    CssName = "houselistbg"
    strSearchResult = "<table width=100%  class='housetable'><tr ><td align='center' class='houseth'>地址</td><td align='center' class='houseth'>户型</td><td align='center' class='houseth'>"
    If rsSearch("ClassID") = 1 Or rsSearch("ClassID") = 3 Then
        strSearchResult = strSearchResult & "价格"
    Else
        strSearchResult = strSearchResult & "租金"
    End If
    strSearchResult = strSearchResult & "</td><td align='center' class='houseth'>面积</td><td align='center' class='houseth'>发布日期</td></tr>"
    Do While Not rsSearch.EOF
        strSearchResult = strSearchResult & "<tr>"
        strSearchResult = strSearchResult & "<td align='center' class='" & CssName & "'><a href='ShowHouse.asp?ClassID=" & rsSearch("ClassID") & "&HouseID=" & rsSearch("HouseID") & "' target=" & OpenType & ">" & rsSearch("HouseDiZhi") & "</a></td>"
        strSearchResult = strSearchResult & "<td align='center' class='" & CssName & "'><a href='ShowHouse.asp?ClassID=" & rsSearch("ClassID") & "&HouseID=" & rsSearch("HouseID") & "' target=" & OpenType & ">" & Replace(rsSearch("HouseHuXing"), ",", "") & "</a></td>"
        Select Case rsSearch("ClassID")
        Case 1
            strSearchResult = strSearchResult & "<td align='center' class='" & CssName & "'>" & rsSearch("TotalPrice") & "" & rsSearch("HousePriceType") & "</td>"
        Case 2
            strSearchResult = strSearchResult & "<td align='center' class='" & CssName & "'>"
            If rsSearch("HouseZuJinType") = "价格面议" Then
                strSearchResult = strSearchResult & rsSearch("HouseZuJinType") & "</td>"
            Else
                strSearchResult = strSearchResult & rsSearch("HouseZuJin") & "" & rsSearch("HouseZuJinType") & "</td>"
            End If
        Case 3
            strSearchResult = strSearchResult & "<td align='center' class='" & CssName & "'>" & rsSearch("HousePrice1") & "-" & rsSearch("HousePrice2") & "" & rsSearch("HousePriceType") & "</td>"
        Case 4
            strSearchResult = strSearchResult & "<td align='center' class='" & CssName & "'>"
            If rsSearch("HouseZuJinType") = "价格面议" Then
                strSearchResult = strSearchResult & rsSearch("HouseZuJinType") & "</td>"
            Else
                strSearchResult = strSearchResult & rsSearch("HouseZuJin1") & "-" & rsSearch("HouseZuJin2") & "" & rsSearch("HouseZuJinType") & "</td>"
            End If
        Case 5
            strSearchResult = strSearchResult & "<td align='center' class='" & CssName & "'>"
            If rsSearch("HouseZuJinType") = "价格面议" Then
                strSearchResult = strSearchResult & rsSearch("HouseZuJinType") & "</td>"
            Else
                strSearchResult = strSearchResult & rsSearch("HouseZuJin") & "" & rsSearch("HouseZuJinType") & "</td>"
            End If
        End Select
        Select Case rsSearch("ClassID")
        Case 1
            strSearchResult = strSearchResult & "<td align='center' class='" & CssName & "'>" & rsSearch("HouseMianJi") & "㎡</td>"
        Case 2
            strSearchResult = strSearchResult & "<td align='center' class='" & CssName & "'>" & rsSearch("HouseMianJi") & "㎡</td>"
        Case 3
            strSearchResult = strSearchResult & "<td align='center' class='" & CssName & "'>" & rsSearch("HouseMianJi1") & "-" & rsSearch("HouseMianJi2") & "㎡</td>"
        Case 4
            strSearchResult = strSearchResult & "<td align='center' class='" & CssName & "'>" & rsSearch("HouseMianJi1") & "-" & rsSearch("HouseMianJi2") & "㎡</td>"
        Case 5
            strSearchResult = strSearchResult & "<td align='center' class='" & CssName & "'>" & rsSearch("HouseMianJi1") & "㎡</td>"
        End Select
        strSearchResult = strSearchResult & "<td align='center' class='" & CssName & "'>" & FormatDateTime(rsSearch("UpdateTime"), 2) & "</td>"
        strSearchResult = strSearchResult & "</tr>"
        HouseNum = HouseNum + 1
        If HouseNum Mod 2 = 0 Then
            CssName = "houselistbg2"
        Else
            CssName = "houselistbg"
        End If
        rsSearch.MoveNext
    Loop
    strSearchResult = strSearchResult & "</table>"
    rsSearch.Close
    Set rsSearch = Nothing
    strFileName = "Search.asp?InfoType=" & InfoType & "&HouseQuYu=" & HouseQuYu & "&HouseSource=" & HouseSource & "&HouseHuXing1=" & HouseHuXing1 & "&HouseHuXing2=" & HouseHuXing2 & "&HouseMianJi=" & HouseMianJi & "&Price1=" & Price1 & "Price2= " & Price2 & "& Address=" & Address & "&TimeBound=" & TimeBound
    GetSearchResult = strSearchResult
End Function

Function GetResultTitle()
    Dim strTitle
    If Keyword = "" Then
        strTitle = "所有" & ChannelShortName
    Else
        Select Case strField
        Case "Title"
            strTitle = "标题含有 <font color=red>" & Keyword & "</font> 的" & ChannelShortName & ""
        Case "Content"
            strTitle = "内容含有 <font color=red>" & Keyword & "</font> 的" & ChannelShortName & ""
        Case "Author"
            strTitle = "作者姓名中含有 <font color=red>" & Keyword & "</font> 的" & ChannelShortName & ""
        Case "Inputer"
            strTitle = "<font color=red>" & Keyword & "</font> 录入的" & ChannelShortName & ""
        Case Else
            strTitle = "标题含有 <font color=red>" & Keyword & "</font> 的" & ChannelShortName & ""
        End Select
    End If
    GetResultTitle = strTitle
End Function

Private Sub ReplaceCommon()
    '以下这段代码放在Call ReplaceCommonLabel的前面，是用于在自定义动态函数标签中可以解析个别标签
    '{$InstallDir}{$ChannelDir}的替换一定要放在单个{$ChannelDir}的前面
    strHtml = PE_Replace(strHtml, "{$InstallDir}{$ChannelDir}", ChannelUrl)
    strHtml = PE_Replace(strHtml, "{$ChannelID}", ChannelID)
    strHtml = PE_Replace(strHtml, "{$ChannelDir}", ChannelDir)
    strHtml = PE_Replace(strHtml, "{$ChannelUrl}", ChannelUrl)

    Call ReplaceCommonLabel
    
    strHtml = PE_Replace(strHtml, "{$InstallDir}{$ChannelDir}", ChannelUrl)
    strHtml = PE_Replace(strHtml, "{$ChannelID}", ChannelID)
    strHtml = PE_Replace(strHtml, "{$ChannelDir}", ChannelDir)
    strHtml = PE_Replace(strHtml, "{$ChannelName}", ChannelName)
    strHtml = PE_Replace(strHtml, "{$ChannelShortName}", ChannelShortName)
    strHtml = PE_Replace(strHtml, "{$UploadDir}", UploadDir)
    strHtml = PE_Replace(strHtml, "{$Meta_Keywords_Channel}", Meta_Keywords_Channel)
    strHtml = PE_Replace(strHtml, "{$Meta_Description_Channel}", Meta_Description_Channel)
    strHtml = PE_Replace(strHtml, "{$MenuJS}", GetMenuJS(ChannelDir, ShowClassTreeGuide))
    strHtml = PE_Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(SkinID))
    '替换底部栏目导航标签
    
End Sub

Function GetHousePic(HousePicWidth, HousePicHeight, PicUrl)
    Dim strHousePic
    If PicUrl = "" Then
        strHousePic = "暂无图片"
    Else
        strHousePic = "<a href='" & PicUrl & "' title='" & SiteName & "' target='_blank'>"
        strHousePic = strHousePic & "<img src='" & PicUrl & "'"
        If HousePicWidth > 0 Then strHousePic = strHousePic & " width='" & HousePicWidth & "'"
        If HousePicHeight > 0 Then strHousePic = strHousePic & " height='" & HousePicHeight & "'"
        strHousePic = strHousePic & " border='0'>"
        strHousePic = strHousePic & "</a>"
    End If
    GetHousePic = strHousePic
End Function

Sub GetHTML_Index()
    Dim strTemp, strTopUser, strFriendSite, arrTemp, strAnnounce, strPopAnnouce, iCols, iClassID
    Dim HouseList_ChildClass, HouseList_ChildClass2
    ClassID = 0

    strHtml = GetTemplate(ChannelID, 1, Template_Index)
    Call ReplaceCommonLabel

    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)
    
    strHtml = Replace(strHtml, "{$ShowChannelCount}", GetChannelCount())
    
    strHtml = GetListFromTemplate(strHtml)
    
    If UseCreateHTML = 0 Then
        If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
        If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    Else
        If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage_Html(ChannelUrl & "/", 0, FileExt_Index, strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
        If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en_Html(ChannelUrl & "/", 0, FileExt_Index, strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName))
    End If
    
End Sub

Sub GetHtml_Class()
    Dim strTemp, iCols, iClassID

    strHtml = arrTemplate(1)
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    Call ReplaceCommonLabel
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", ShowPath())
    
    strHtml = PE_Replace(strHtml, "{$ClassName}", ClassName)
    strHtml = PE_Replace(strHtml, "{$ParentDir}", ParentDir)
    strHtml = PE_Replace(strHtml, "{$ClassDir}", ClassDir)
    strHtml = PE_Replace(strHtml, "{$Readme}", ReadMe)
    strHtml = PE_Replace(strHtml, "{$Meta_Keywords_Class}", Meta_Keywords_Class)
    strHtml = PE_Replace(strHtml, "{$Meta_Description_Class}", Meta_Description_Class)
    strHtml = Replace(strHtml, "{$ClassUrl}", GetClassUrl(ParentDir, ClassDir, ClassID, ClassPurview))
    strHtml = Replace(strHtml, "{$ClassListUrl}", GetClass_1Url(ParentDir, ClassDir, ClassID, ClassPurview))
    strHtml = PE_Replace(strHtml, "{$GetAreaList}", GetAreaList())

    Dim HouseList_CurrentClass, HouseList_CurrentClass2, HouseList_ChildClass, HouseList_ChildClass2

    strHtml = GetListFromTemplate(strHtml)
    Dim strPath
    strPath = ChannelUrl & GetListPath(StructureType, ListFileType, ParentDir, ClassDir)
    
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
End Sub

Public Sub GetHtml_List()
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    Call ReplaceCommonLabel
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", ShowPath())
    strHtml = PE_Replace(strHtml, "{$ClassName}", ClassName)
    strHtml = PE_Replace(strHtml, "{$ParentDir}", ParentDir)
    strHtml = PE_Replace(strHtml, "{$ClassDir}", ClassDir)
    strHtml = PE_Replace(strHtml, "{$Readme}", ReadMe)
    'strHTML = PE_Replace(strHTML, "{$SpecialName}", SpecialName)
    strHtml = GetListFromTemplate(strHtml)
    
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    
End Sub

Public Sub GetHtml_Search()
    Dim SearchChannelID
    SearchChannelID = ChannelID
    If ChannelID > 0 Then
        strHtml = GetTemplate(ChannelID, 5, 0)
    Else
        strHtml = GetTemplate(ChannelID, 5, 0)
        ChannelID = PE_CLng(Conn.Execute("select min(ChannelID) from PE_Channel where ModuleType=7 and Disabled=" & PE_False & "")(0))
        CurrentChannelID = ChannelID
        Call GetChannel(ChannelID)
    End If

    Call ReplaceCommonLabel
    strHtml = Replace(strHtml, "{$Keyword}", Keyword)
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", ShowPath())
    Call GetClass
    strHtml = PE_Replace(strHtml, "{$ClassID}", ClassID)
    strHtml = PE_Replace(strHtml, "{$ClassName}", ClassName)
    strHtml = PE_Replace(strHtml, "{$ParentDir}", ParentDir)
    strHtml = PE_Replace(strHtml, "{$ClassDir}", ClassDir)

    MaxPerPage = 10
    strHtml = Replace(strHtml, "{$ResultTitle}", GetResultTitle())
    strHtml = Replace(strHtml, "{$SearchResult}", GetSearchResult())

    'strFileName = "Search.asp?InfoType="&InfoType&"&Field=" & strField & "&Keyword=" & Keyword & "&ClassID=" & ClassID
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
    strHtml = GetListFromTemplate(strHtml)
End Sub

Function XmlText_Class(ByVal iSmallNode, ByVal DefChar)
    XmlText_Class = XmlText("House", iSmallNode, DefChar)
End Function

Function R_XmlText_Class(ByVal iSmallNode, ByVal DefChar)
    R_XmlText_Class = Replace(XmlText("House", iSmallNode, DefChar), "{$ChannelShortName}", ChannelShortName)
End Function

End Class
%>
