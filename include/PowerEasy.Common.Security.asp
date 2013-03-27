<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'**************************************************
'函数名：PE_CBool
'作  用：将字符转为布尔弄变量
'参  数：strBool---- 字符
'返回值：True/False
'**************************************************
Function PE_CBool(strBool)
    If strBool = True Or LCase(Trim(strBool)) = "true" Or LCase(Trim(strBool)) = "yes" Or Trim(strBool) = "1" Then
        PE_CBool = True
    Else
        PE_CBool = False
    End If
End Function

'**************************************************
'函数名：PE_CLng
'作  用：将字符转为整型数值
'参  数：str1 ---- 字符
'返回值：如果传入的参数不是数值，返回0，其他情况返回对应的数值
'**************************************************
Function PE_CLng(ByVal str1)
    If IsNumeric(str1) Then
        PE_CLng = Fix(CDbl(str1))
    Else
        PE_CLng = 0
    End If
End Function

'**************************************************
'函数名：PE_CLng1
'作  用：将字符转为整型数值
'参  数：str1 ---- 字符
'返回值：如果传入的参数不是数值，返回1，其他情况返回对应的数值
'**************************************************
Function PE_CLng1(ByVal str1)
    If IsNumeric(str1) Then
        PE_CLng1 = CLng(str1)
        If PE_CLng1 <= 0 Then PE_CLng1 = 1
    Else
        PE_CLng1 = 1
    End If
End Function

'**************************************************
'函数名：PE_CDbl
'作  用：将字符转为双精度数值
'参  数：str1 ---- 字符
'返回值：如果传入的参数不是数值，返回0，其他情况返回对应的数值
'**************************************************
Function PE_CDbl(ByVal str1)
    If IsNumeric(str1) Then
        PE_CDbl = CDbl(str1)
    Else
        PE_CDbl = 0
    End If
End Function

'**************************************************
'函数名：PE_CDate
'作  用：将字符转为日期
'参  数：str1 ---- 字符
'返回值：如果参数不是日期型字符，则返回当前时间，否则返回对应的日期型数据
'**************************************************
Function PE_CDate(ByVal str1)
    If IsDate(str1) Then
        PE_CDate = CDate(str1)
    Else
        PE_CDate = Now
    End If
End Function

'**************************************************
'函数名：EncodeIP
'作  用：将IP地址转为数字
'参  数：Sip ---- IP地址
'返回值：数字
'**************************************************
Function EncodeIP(sip)
    Dim strIP
    strIP = Split(sip, ".")
    If UBound(strIP) < 3 Then
        EncodeIP = 0
        Exit Function
    End If
    If IsNumeric(strIP(0)) = False Or IsNumeric(strIP(1)) = False Or IsNumeric(strIP(2)) = False Or IsNumeric(strIP(3)) = False Then
        sip = 0
    Else
        sip = CSng(strIP(0)) * 256 * 256 * 256 + CLng(strIP(1)) * 256 * 256 + CLng(strIP(2)) * 256 + CLng(strIP(3)) - 1
    End If
    EncodeIP = sip
End Function

'**************************************************
'函数名：
'作  用：
'参  数：
'返回值：
'**************************************************
'白名单的端点可以访问和黑名单的端点将不允许访问。
Function ChecKIPlock(ByVal sLockType, ByVal sLockList, ByVal sUserIP)
    Dim IPlock, rsLockIP
    Dim arrLockIPW, arrLockIPB, arrLockIPWCut, arrLockIPBCut
    IPlock = False
    ChecKIPlock = IPlock
    Dim i, sKillIP
    If sLockType = "" Or IsNull(sLockType) Then Exit Function
    If sLockList = "" Or IsNull(sLockList) Then Exit Function
    If sUserIP = "" Or IsNull(sUserIP) Then Exit Function
    sUserIP = CDbl(EncodeIP(sUserIP))
    rsLockIP = Split(sLockList, "|||")
    If sLockType = 4 Then
        arrLockIPB = Split(Trim(rsLockIP(1)), "$$$")
        For i = 0 To UBound(arrLockIPB)
            If arrLockIPB(i) <> "" Then
                arrLockIPBCut = Split(Trim(arrLockIPB(i)), "----")
                IPlock = True
                If CDbl(arrLockIPBCut(0)) > sUserIP Or sUserIP > CDbl(arrLockIPBCut(1)) Then IPlock = False
                If IPlock Then Exit For
            End If
        Next
        If IPlock = True Then
            arrLockIPW = Split(Trim(rsLockIP(0)), "$$$")
            For i = 0 To UBound(arrLockIPW)
                If arrLockIPW(i) <> "" Then
                    arrLockIPWCut = Split(Trim(arrLockIPW(i)), "----")
                    IPlock = True
                    If CDbl(arrLockIPWCut(0)) <= sUserIP And sUserIP <= CDbl(arrLockIPWCut(1)) Then IPlock = False
                    If IPlock Then Exit For
                End If
            Next
        End If
    Else
        If sLockType = 1 Or sLockType = 3 Then
            arrLockIPW = Split(Trim(rsLockIP(0)), "$$$")
            For i = 0 To UBound(arrLockIPW)
                If arrLockIPW(i) <> "" Then
                    arrLockIPWCut = Split(Trim(arrLockIPW(i)), "----")
                    IPlock = True
                    If CDbl(arrLockIPWCut(0)) <= sUserIP And sUserIP <= CDbl(arrLockIPWCut(1)) Then IPlock = False
                    If IPlock = False Then Exit For
                End If
            Next
        End If
        If IPlock = False And (sLockType = 2 Or sLockType = 3) Then
            arrLockIPB = Split(Trim(rsLockIP(1)), "$$$")
            For i = 0 To UBound(arrLockIPB)
                If arrLockIPB(i) <> "" Then
                    arrLockIPBCut = Split(Trim(arrLockIPB(i)), "----")
                    IPlock = True
                    If CDbl(arrLockIPBCut(0)) > sUserIP Or sUserIP > CDbl(arrLockIPBCut(1)) Then IPlock = False
                    If IPlock Then Exit For
                End If
            Next
        End If
    End If
    ChecKIPlock = IPlock
End Function


'**************************************************
'函数名：IsValidEmail
'作  用：检查Email地址合法性
'参  数：email ----要检查的Email地址
'返回值：True  ----Email地址合法
'        False ----Email地址不合法
'**************************************************
Function IsValidEmail(Email)
    regEx.Pattern = "^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$"
    IsValidEmail = regEx.Test(Email)
End Function


'**************************************************
'函数名：IsValidStr
'作  用：检查字符是否在有效范围内
'参  数：str ----要检查的字符
'返回值：True  ----字符合法
'        False ----字符不合法
'**************************************************
Function IsValidStr(ByVal str)
    Dim i, c
    For i = 1 To Len(str)
        c = LCase(Mid(str, i, 1))
        If InStr("abcdefghijklmnopqrstuvwxyz1234567890", c) <= 0 Then
            IsValidStr = False
            Exit Function
        End If
    Next
    If IsNumeric(Left(str, 1)) Then
        IsValidStr = False
    Else
        IsValidStr = True
    End If
End Function

'**************************************************
'函数名：IsValidJsFileName
'作  用：检查是否是有效的JS文件名
'参  数：str ----要检查的字符
'返回值：True  ----文件名合法
'        False ----文件名不合法
'**************************************************
Function IsValidJsFileName(ByVal str, ByVal ContentType)
    Dim i, c
    For i = 1 To Len(str)
        c = LCase(Mid(str, i, 1))
        If InStr("abcdefghijklmnopqrstuvwxyz_1234567890.", c) <= 0 Then
            IsValidJsFileName = False
            Exit Function
        End If
    Next
    If ContentType = 0 Then
        If LCase(Right(str, 3)) <> ".js" Then
            IsValidJsFileName = False
        Else
            IsValidJsFileName = True
        End If
    Else
        If LCase(Right(str, 5)) <> ".html" Then
            IsValidJsFileName = False
        Else
            IsValidJsFileName = True
        End If
    End If
End Function

'**************************************************
'函数名：ReplaceBadChar
'作  用：过滤非法的SQL字符
'参  数：strChar-----要过滤的字符
'返回值：过滤后的字符
'**************************************************
Function ReplaceBadChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceBadChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ",--"
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceBadChar = tempChar
End Function

'**************************************************
'函数名：ReplaceLabelBadChar
'作  用：函数标签过滤非法的SQL字符
'参  数：strChar-----要过滤的字符
'返回值：过滤后的字符
'**************************************************
Function ReplaceLabelBadChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceLabelBadChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0)
	arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    Dim oldString
    oldString = ""
    Do While oldString <> tempChar
        oldString = tempChar
        regEx.Pattern = "(select|union|update|insert|delete|exec|from|pe_admin|--)?"
        tempChar = regEx.Replace(tempChar, "")
    Loop
    ReplaceLabelBadChar = tempChar
End Function

'**************************************************
'函数名：ReplaceUrlBadChar
'作  用：过滤Url中非法的SQL字符
'参  数：strChar-----要过滤的字符
'返回值：过滤后的字符
'**************************************************
Function ReplaceUrlBadChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceUrlBadChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',(,),<,>,[,],{,},\,;," & Chr(34) & "," & Chr(0) & ",--"
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceUrlBadChar = tempChar
End Function
'=================================================
'函数名：ReplaceBadUrl
'作  用：过滤非法Url地址函数
'=================================================
Function ReplaceBadUrl(ByVal strContent)
    regEx.Pattern = "(a|%61|%41)(d|%64|%44)(m|%6D|4D)(i|%69|%49)(n|%6E|%4E)(\_|%5F)(.*?)(.|%2E)(a|%61|%41)(s|%73|%53)(p|%70|%50)"
    Set Matches = regEx.Execute(strContent)
    For Each Match In Matches
        strContent = Replace(strContent, Match.value, "")
    Next
    regEx.Pattern = "(u|%75|%55)(s|%73|%53)(e|%65|%45)(r|%72|%52)(\_|%5F)(.*?)(.|%2E)(a|%61|%41)(s|%73|%53)(p|%70|%50)"
    Set Matches = regEx.Execute(strContent)
    For Each Match In Matches
        strContent = Replace(strContent, Match.value, "")
    Next
    ReplaceBadUrl = strContent
End Function

'**************************************************
'函数名：CheckBadChar
'作  用：检查是否包含非法的SQL字符
'参  数：strChar-----要检查的字符
'返回值：True  ----字符合法
'        False ----字符不合法
'**************************************************
Function CheckBadChar(strChar)
    Dim strBadChar, arrBadChar, i
    strBadChar = "@@,+,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & ",--,union,select,insert,delete,from,pe_admin"
    arrBadChar = Split(strBadChar, ",")
    If strChar = "" Then
        CheckBadChar = False
    Else
        Dim tempChar
        tempChar = LCase(strChar)
        For i = 0 To UBound(arrBadChar)
            If InStr(tempChar, arrBadChar(i)) > 0 Then
                CheckBadChar = False
                Exit Function
            End If
        Next
    End If
    CheckBadChar = True
End Function


Function CheckUserBadChar(strChar)
    Dim strBadChar, arrBadChar, i
    strBadChar = "',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & ",*,|,"",.,#,union,select,insert,delete,from,pe_admin"
    arrBadChar = Split(strBadChar, ",")
    If strChar = "" Then
        CheckUserBadChar = False
    Else
        Dim tempChar
        tempChar = LCase(strChar)
        For i = 0 To UBound(arrBadChar)
            If InStr(tempChar, arrBadChar(i)) > 0 Then
                CheckUserBadChar = False
                Exit Function
            End If
        Next
    End If
    CheckUserBadChar = True
    
End Function

'**************************************************
'函数名：CheckValidStr
'作  用：检查数组中有无相同的字符
'参  数：arrInvalidStr ----要查询的数组
'        str1 ---- 要比较的字符
'返回值：True  ----是否存在
'**************************************************
Function CheckValidStr(arrInvalidStr, str1)
    Dim arrStr, i
    If InStr(arrInvalidStr, ",") > 0 Then
        arrStr = Split(arrInvalidStr, ",")
        For i = 0 To UBound(arrStr)
            If LCase(Trim(arrStr(i))) = LCase(Trim(str1)) Then
                CheckValidStr = False
                Exit Function
            End If
        Next
    Else
        If LCase(Trim(arrInvalidStr)) = LCase(Trim(str1)) Then
            CheckValidStr = False
            Exit Function
        End If
    End If
    CheckValidStr = True
End Function

'**************************************************
'函数名：IsValidID
'作  用：检查传过来的ＩＤ是否是合法ＩＤ或者ＩＤ串
'参  数：Check_ID ---- ID 字符串
'返回值：True  ---- 合法ID
'**************************************************
Function IsValidID(Check_ID)
    Dim FixID, i
    If IsNull(Check_ID) Or Check_ID = "" Then
        IsValidID = False
        Exit Function
    End If
    FixID = Replace(Check_ID, "|", "")
    FixID = Replace(FixID, ",", "")
    FixID = Replace(FixID, "-", "")
    FixID = Trim(Replace(FixID, " ", ""))
    If FixID = "" Or IsNull(FixID) Then
        IsValidID = False
    Else
        For i = 1 To Len(FixID) Step 100
            If Not IsNumeric(Mid(FixID, i, 100)) Then
                IsValidID = False
                Exit Function
            End If
        Next
        IsValidID = True
    End If
End Function

'**************************************************
'函数名：PE_ConvertBR
'作  用：将文本区域内的<BR>替换换行
'参  数：fString ---- 要处理的字符串
'返回值：处理后的字符串
'**************************************************
Function PE_ConvertBR(ByVal fString)
    If IsNull(fString) Or Trim(fString) = "" Then
        PE_ConvertBR = ""
        Exit Function
    End If
    fString = Replace(fString, "</P><P>", Chr(10) & Chr(10))
    fString = Replace(fString, "<BR>", Chr(10))
    fString = Replace(fString, "<br>", Chr(10))
    PE_ConvertBR = fString
End Function

'**************************************************
'函数名：PE_HTMLEncode
'作  用：将html 标记替换成 能在IE显示的HTML
'参  数：fString ---- 要处理的字符串
'返回值：处理后的字符串
'**************************************************
Function PE_HTMLEncode(ByVal fString)
    If IsNull(fString) Or Trim(fString) = "" Then
        PE_HTMLEncode = ""
        Exit Function
    End If
    fString = Replace(fString, ">", "&gt;")
    fString = Replace(fString, "<", "&lt;")

    fString = Replace(fString, Chr(32), "&nbsp;")
    fString = Replace(fString, Chr(9), "&nbsp;")
    fString = Replace(fString, Chr(34), "&quot;")
    fString = Replace(fString, Chr(39), "&#39;")
    fString = Replace(fString, Chr(13), "")
    fString = Replace(fString, Chr(10) & Chr(10), "</P><P>")
    fString = Replace(fString, Chr(10), "<BR>")

    PE_HTMLEncode = fString
End Function


'**************************************************
'函数名：PE_HtmlDecode
'作  用：还原Html标记,配合PE_HTMLEncode 使用
'参  数：fString ---- 要处理的字符串
'返回值：处理后的字符串
'**************************************************
Function PE_HtmlDecode(ByVal fString)
    If IsNull(fString) Or Trim(fString) = "" Then
        PE_HtmlDecode = ""
        Exit Function
    End If
    fString = Replace(fString, "&gt;", ">")
    fString = Replace(fString, "&lt;", "<")

    fString = Replace(fString, "&nbsp;", " ")
    fString = Replace(fString, "&quot;", Chr(34))
    fString = Replace(fString, "&#39;", Chr(39))
    fString = Replace(fString, "</P><P> ", Chr(10) & Chr(10))
    fString = Replace(fString, "<BR> ", Chr(10))

    PE_HtmlDecode = fString
End Function


'**************************************************
'函数名：nohtml
'作  用：过滤html 元素
'参  数：str ---- 要过滤字符
'返回值：没有html 的字符
'**************************************************
Function nohtml(ByVal str)
    If IsNull(str) Or Trim(str) = "" Then
        nohtml = ""
        Exit Function
    End If
    regEx.Pattern = "(\<.[^\<]*\>)"
    str = regEx.Replace(str, "")
    regEx.Pattern = "(\<\/[^\<]*\>)"
    str = regEx.Replace(str, "")
    regEx.Pattern = "\[NextPage(.*?)\]"   '解决“当在文章模块的频道中发布的是图片并使用分页标签[NextPage]或内容开始的前几行就使用分页标签时，一旦使用搜索来搜索该文时，搜索页就会显示分页标签”的问题
    str = regEx.Replace(str, "")
    
    str = Replace(str, "'", "")
    str = Replace(str, Chr(34), "")
    str = Replace(str, vbCrLf, "")
    str = Trim(str)
    nohtml = str
End Function

'**************************************************
'函数名：xml_nohtml
'作  用：过滤xml 和 html 元素
'参  数：str ---- 要过滤字符
'返回值：没有 xml 和 html 的字符串
'**************************************************
Function xml_nohtml(ByVal fString)
    If IsNull(fString) Or Trim(fString) = "" Then
        xml_nohtml = ""
        Exit Function
    End If
    Dim str, k
    str = Replace(fString, "&gt;", ">")
    str = Replace(str, "&lt;", "<")
    str = Replace(str, "&nbsp;", "")
    str = Replace(str, "&quot;", "")
    str = Replace(str, "&#39;", "")

    str = nohtml(str)
    str = Replace(Replace(str, "<![CDATA[", ""), "]]>", "")
    xml_nohtml = str
End Function

'**************************************************
'函数名：unicode
'作  用：转换为 UTF8 编码
'参  数：str ---- 要转换的字符
'返回值：转换后的字符
'**************************************************
Function unicode(ByVal str)
    Dim i, j, c, i1, i2, u, fs, f, p
    unicode = ""
    p = ""
    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        j = AscW(c)
        If j < 0 Then
            j = j + 65536
        End If
        If j >= 0 And j <= 128 Then
            If p = "c" Then
                unicode = " " & unicode
                p = "e"
            End If
            unicode = unicode & c
        Else
            If p = "e" Then
                unicode = unicode & " "
                p = "c"
            End If
            unicode = unicode & ("&#" & j & ";")
        End If
    Next
End Function

'**************************************************
'函数名：Jencode
'作  用：替换那26个片假名字符(效率很差目前没有用到)
'参  数：str ---- 要替换的字符
'        DatabaseType ---- 数据库类型
'返回值：替换后的字符
'**************************************************
Function Jencode(ByVal iStr, DatabaseType)
    If IsNull(iStr) Or IsEmpty(iStr) Or iStr = "" Or DatabaseType = "SQL" Then
        Jencode = ""
        Exit Function
    End If
    Dim E, f, i
    E = Array("Jn0;", "Jn1;", "Jn2;", "Jn3;", "Jn4;", "Jn5;", "Jn6;", "Jn7;", "Jn8;", "Jn9;", "Jn10;", "Jn11;", "Jn12;", "Jn13;", "Jn14;", "Jn15;", "Jn16;", "Jn17;", "Jn18;", "Jn19;", "Jn20;", "Jn21;", "Jn22;", "Jn23;", "Jn24;", "Jn25;")
    f = Array(Chr(-23116), Chr(-23124), Chr(-23122), Chr(-23120), Chr(-23118), Chr(-23114), Chr(-23112), Chr(-23110), Chr(-23099), Chr(-23097), Chr(-23095), Chr(-23075), Chr(-23079), Chr(-23081), Chr(-23085), Chr(-23087), Chr(-23052), Chr(-23076), Chr(-23078), Chr(-23082), Chr(-23084), Chr(-23088), Chr(-23102), Chr(-23104), Chr(-23106), Chr(-23108))
    Jencode = iStr
    For i = 0 To 25
        Jencode = Replace(Jencode, f(i), E(i))
    Next
End Function

Function Juncode(ByVal iStr, DatabaseType)
    If IsNull(iStr) Or IsEmpty(iStr) Or iStr = "" Or DatabaseType = "SQL" Then
        Juncode = ""
        Exit Function
    End If
    Dim E, f, i
    E = Array("Jn0;", "Jn1;", "Jn2;", "Jn3;", "Jn4;", "Jn5;", "Jn6;", "Jn7;", "Jn8;", "Jn9;", "Jn10;", "Jn11;", "Jn12;", "Jn13;", "Jn14;", "Jn15;", "Jn16;", "Jn17;", "Jn18;", "Jn19;", "Jn20;", "Jn21;", "Jn22;", "Jn23;", "Jn24;", "Jn25;")
    f = Array(Chr(-23116), Chr(-23124), Chr(-23122), Chr(-23120), Chr(-23118), Chr(-23114), Chr(-23112), Chr(-23110), Chr(-23099), Chr(-23097), Chr(-23095), Chr(-23075), Chr(-23079), Chr(-23081), Chr(-23085), Chr(-23087), Chr(-23052), Chr(-23076), Chr(-23078), Chr(-23082), Chr(-23084), Chr(-23088), Chr(-23102), Chr(-23104), Chr(-23106), Chr(-23108))
    Juncode = iStr
    For i = 0 To 25
        Juncode = Replace(Juncode, E(i), f(i))
    Next
End Function


Function IsValidPhone(Phone)
    Dim i, c
    IsValidPhone = True
    For i = 1 To Len(Phone)
        c = LCase(Mid(Phone, i, 1))
        If InStr("-()", c) <= 0 And Not IsNumeric(c) Then
            IsValidPhone = False
            Exit Function
        End If
    Next
End Function


'**************************************************
'函数名：DelRightComma
'作  用：删除字符串（如："1,3,5,8"）右侧多余的逗号以消除SQL查询时出错的问题，Comma：逗号。
'参  数：str ---- 待处理的字符串
'**************************************************
Function DelRightComma(ByVal str)
    str = Trim(str)
    If Right(str, 1) = "," Then
        str = Left(str, Len(str) - 1)
    End If
    DelRightComma = str
End Function

'**************************************************
'函数名：FilterArrNull
'作  用：过滤数组空字符
'**************************************************
Function FilterArrNull(ByVal ArrString, ByVal CompartString)
    Dim arrContent, arrTemp, i

    If CompartString = "" Or ArrString = "" Then
        FilterArrNull = ArrString
        Exit Function
    End If
    If InStr(ArrString, CompartString) = 0 Then
        FilterArrNull = ArrString
        Exit Function
    Else
        arrContent = Split(ArrString, CompartString)
        For i = 0 To UBound(arrContent)
            If Trim(arrContent(i)) <> "" Then
                If arrTemp = "" Then
                    arrTemp = Trim(arrContent(i))
                Else
                    arrTemp = arrTemp & CompartString & Trim(arrContent(i))
                End If
            End If
        Next
    End If
    FilterArrNull = arrTemp
End Function
'=================================================
'函数名：FilterJS()
'作  用：过滤非法JS字符
'参  数：strInput 需要过滤的内容
'=================================================
Function FilterJS(ByVal strInput)
    If IsNull(strInput) Or Trim(strInput) = "" Then
        FilterJS = ""
        Exit Function
    End If
    Dim reContent

    ' 替换掉HTML字符实体(Character Entities)名字和分号之间的空白字符，比如：&auml    ;替换成&auml;
    regEx.Pattern = "(&#*\w+)[\x00-\x20]+;"
    strInput = regEx.Replace(strInput, "$1;")

    ' 将无分号结束符的数字编码实体规范成带分号的标准形式
    regEx.Pattern = "(&#x*[0-9A-F]+);*"
    strInput = regEx.Replace(strInput, "$1;")

    ' 将&nbsp; &lt; &gt; &amp; &quot;字符实体中的 & 替换成 &amp; 以便在进行HtmlDecode时保留这些字符实体
    'RegEx.Pattern = "&(amp|lt|gt|nbsp|quot);"
    'strInput = RegEx.Replace(strInput, "&amp;$1;")

    ' 将HTML字符实体进行解码，以消除编码字符对后续过滤的影响
    'strInput = HtmlDecode(strInput);

    ' 将ASCII码表中前32个字符中的非打印字符替换成空字符串，保留 9、10、13、32，它们分别代表 制表符、换行符、回车符和空格。
    regEx.Pattern = "[\x00-\x08\x0b-\x0c\x0e-\x19]"
    strInput = regEx.Replace(strInput, "")  
       
    oldhtmlString = ""
    Do While oldhtmlString <> strInput
        oldhtmlString = strInput
        regEx.Pattern = "(<[^>]+src[\x00-\x20]*=[\x00-\x20]*[^>]*?)&#([^>]*>)"  '过虑掉 src 里的 &#
        strInput = regEx.Replace(strInput, "$1&amp;#$2")
        regEx.Pattern = "(<[^>]+style[\x00-\x20]*=[\x00-\x20]*[^>]*?)&#([^>]*>)"  '过虑掉style 里的 &#
        strInput = regEx.Replace(strInput, "$1&amp;#$2")
        regEx.Pattern = "(<[^>]+style[\x00-\x20]*=[\x00-\x20]*[^>]*?)\\([^>]*>)"   '替换掉style中的 "\" 
        strInput = regEx.Replace(strInput, "$1/$2")  
    Loop
    ' 替换以on和xmlns开头的属性，动易系统的几个JS需要保留
    regEx.Pattern = "on(load\s*=\s*""*'*resizepic\(this\)'*""*)"
    strInput = regEx.Replace(strInput, "off$1")
    regEx.Pattern = "on(mousewheel\s*=\s*""*'*return\s*bbimg\(this\)'*""*)"
    strInput = regEx.Replace(strInput, "off$1")

    regEx.Pattern = "(<[^>]+[\x00-\x20""'/])(on|xmlns)([^>]*)>"
    strInput = regEx.Replace(strInput, "$1pe$3>")

    regEx.Pattern = "off(load\s*=\s*""*'*resizepic\(this\)'*""*)"
    strInput = regEx.Replace(strInput, "on$1")
    regEx.Pattern = "off(mousewheel\s*=\s*""*'*return\s*bbimg\(this\)'*""*)"
    strInput = regEx.Replace(strInput, "on$1")

    
    ' 替换javascript
    regEx.Pattern = "([a-z]*)[\x00-\x20]*=[\x00-\x20]*([`'""]*)[\x00-\x20]*j[\x00-\x20]*a[\x00-\x20]*v[\x00-\x20]*a[\x00-\x20]*s[\x00-\x20]*c[\x00-\x20]*r[\x00-\x20]*i[\x00-\x20]*p[\x00-\x20]*t[\x00-\x20]*:"
    strInput = regEx.Replace(strInput, "$1=$2nojavascript...")

    ' 替换vbscript
    regEx.Pattern = "([a-z]*)[\x00-\x20]*=[\x00-\x20]*([`'""]*)[\x00-\x20]*v[\x00-\x20]*b[\x00-\x20]*s[\x00-\x20]*c[\x00-\x20]*r[\x00-\x20]*i[\x00-\x20]*p[\x00-\x20]*t[\x00-\x20]*:"
    strInput = regEx.Replace(strInput, "$1=$2novbscript...")

    '替换style中的注释部分，比如：<div style="xss:expres/*comment*/sion(alert(x))">
    regEx.Pattern = "(<[^>]+style[\x00-\x20]*=[\x00-\x20]*[^>]*?)/\*[^>]*\*/([^>]*>)"
    strInput = regEx.Replace(strInput, "$1$2")
    ' 替换expression
    regEx.Pattern = "(<[^>]+)style[\x00-\x20]*=[\x00-\x20]*([`'""]*).*[eｅＥ][xｘＸ][pｐＰ][rｒＲ][eｅＥ][sｓＳ][sｓＳ][iｉＩ][oｏＯ][nｎＮ][\x00-\x20]*[\(\（][^>]*>"
    strInput = regEx.Replace(strInput, "$1>")

    ' 替换behaviour
    regEx.Pattern = "(<[^>]+)style[\x00-\x20]*=[\x00-\x20]*([`'""]*).*behaviour[^>]*>>"
    strInput = regEx.Replace(strInput, "$1>")
    ' 替换behavior
    regEx.Pattern = "(<[^>]+)style[\x00-\x20]*=[\x00-\x20]*([`'""]*).*behavior[^>]*>>"
    strInput = regEx.Replace(strInput, "$1>")

    ' 替换script
    regEx.Pattern = "(<[^>]+)style[\x00-\x20]*=[\x00-\x20]*([`'""]*).*s[\x00-\x20]*c[\x00-\x20]*r[\x00-\x20]*i[\x00-\x20]*p[\x00-\x20]*t[\x00-\x20]*:*[^>]*>"
    strInput = regEx.Replace(strInput, "$1>")

    ' 替换namespaced elements 不需要
    regEx.Pattern = "</*\w+:\w[^>]*>"
    strInput = regEx.Replace(strInput, "　")

    Dim oldhtmlString
    oldhtmlString = ""
    Do While oldhtmlString <> strInput
        oldhtmlString = strInput
        '实行严格过滤
        regEx.Pattern = "</*(applet|meta|xml|blink|link|style|script|embed|object|iframe|frame|frameset|ilayer|layer|bgsound|title|base)[^>]*>?"
        strInput = regEx.Replace(strInput, "　")
        '过滤掉SHTML的Include包含文件漏洞
        regEx.Pattern = "<!--\s*#include[^>]*>"
        strInput = regEx.Replace(strInput, "noshtml")
        'If FilterLevel > 0 Then
        '   '实行严格过滤
        '   RegEx.Pattern = "</*(embed|object)[^>]*>"
        '   strInput = RegEx.Replace(strInput, "")
        'End If
    Loop
    FilterJS = strInput
End Function

Private Function RemoveStr(str1, str2, strSplit)
    If IsNull(str1) Or str1 = "" Then
        RemoveStr = ""
        Exit Function
    End If
    If IsNull(str2) Or str2 = "" Then
        RemoveStr = str1
        Exit Function
    End If
    If InStr(str1, strSplit) > 0 Then
        Dim arrStr, tempStr, i
        arrStr = Split(str1, strSplit)
        For i = 0 To UBound(arrStr)
            If arrStr(i) <> str2 Then
                If tempStr = "" Then
                    tempStr = arrStr(i)
                Else
                    tempStr = tempStr & strSplit & arrStr(i)
                End If
            End If
        Next
        RemoveStr = tempStr
    Else
        If str1 = str2 Then
            RemoveStr = ""
        Else
            RemoveStr = str1
        End If
    End If
End Function

Private Function AppendStr(str1, str2, strSplit)
    If IsNull(str2) Or str2 = "" Then
        AppendStr = str1
        Exit Function
    End If
    If IsNull(str1) Or str1 = "" Then
        AppendStr = str2
        Exit Function
    End If
    Dim Foundstr, arrStr, i
    Foundstr = False
    If InStr(str1, strSplit) > 0 Then
        arrStr = Split(str1, strSplit)
        For i = 0 To UBound(arrStr)
            If arrStr(i) = str2 Then
                Foundstr = True
                Exit For
            End If
        Next
    Else
        If str1 = str2 Then
            Foundstr = True
        End If
    End If
    If Foundstr = False Then
        AppendStr = str1 & strSplit & str2
    Else
        AppendStr = str1
    End If
End Function

Private Function StyleDisplay(Compare1, Compare2)
    If Compare1 = Compare2 Then
        StyleDisplay = ""
    Else
        StyleDisplay = "none"
    End If
End Function

Private Function IsRadioChecked(Compare1, Compare2)
    If Compare1 = Compare2 Then
        IsRadioChecked = " checked"
    Else
        IsRadioChecked = ""
    End If
End Function

Private Function IsOptionSelected(Compare1, Compare2)
    If Compare1 = Compare2 Then
        IsOptionSelected = " selected"
    Else
        IsOptionSelected = ""
    End If
End Function

Private Function FixJs(str)
    If str <> "" Then
        str = Replace(str, "&#39;", "'")
        str = Replace(str, "\", "\\")
        str = Replace(str, Chr(34), "\""")
        str = Replace(str, Chr(39), "\'")
        str = Replace(str, Chr(13), "\n")
        str = Replace(str, Chr(10), "\r")
        str = Replace(str, "'", "&#39;")
        str = Replace(str, """", "&quot;")
    End If
    FixJs = str
End Function

Private Function Html2Js(str)
    If str <> "" Then
        str = Replace(str, Chr(34), "\""")
        str = Replace(str, Chr(39), "\'")
        str = Replace(str, Chr(13), "\n")
        str = Replace(str, Chr(10), "\r")
    End If
    Html2Js = str
End Function

'==================================================
'函数名：ScriptHtml
'作  用：过滤html标记
'参  数：iConStr  ------ 要过滤的字符串
'参  数：TagName ------ 字符串种型
'参  数：FType   ------ 过滤的类型
'==================================================
Function ScriptHtml(ByVal iConStr, TagName, FType)
    Dim ConStr
    ConStr = iConStr
    Select Case FType
    Case 1
        regEx.Pattern = "<" & TagName & "([^>])*>"
        ConStr = regEx.Replace(ConStr, "")
    Case 2
        regEx.Pattern = "<" & TagName & "([^>])*>[\s\S]*?</" & TagName & "([^>])*>"
        ConStr = regEx.Replace(ConStr, "")
    Case 3
        regEx.Pattern = "<" & TagName & "([^>])*>"
        ConStr = regEx.Replace(ConStr, "")
        regEx.Pattern = "</" & TagName & "([^>])*>"
        ConStr = regEx.Replace(ConStr, "")
    End Select
    ScriptHtml = ConStr
End Function

'==================================================
'过程名：FilterScript
'作  用：脚本过滤
'==================================================
Function FilterScript(ByVal iContent, iScript)
    If IsNull(iContent) = True Then Exit Function
    If IsNull(iScript) = True Then
        iScript = "0|0|0|0|0|0|0|0|0|0|0|0|0"
    End If
    Dim Script_Property, Content
    Script_Property = Split(iScript, "|")
    Content = iContent
    If PE_CBool(Script_Property(0)) = True Then
        Content = ScriptHtml(Content, "Iframe", 2)
    End If
    If PE_CBool(Script_Property(1)) = True Then
        Content = ScriptHtml(Content, "Object", 2)
    End If
    If PE_CBool(Script_Property(2)) = True Then
        Content = ScriptHtml(Content, "Script", 2)
    End If
    If PE_CBool(Script_Property(3)) = True Then
        Content = ScriptHtml(Content, "Style", 2)
    End If
    If PE_CBool(Script_Property(4)) = True Then
        Content = ScriptHtml(Content, "Div", 3)
    End If
    If PE_CBool(Script_Property(5)) = True Then
        Content = ScriptHtml(Content, "Table", 3)
        Content = ScriptHtml(Content, "Tbody", 3)
    End If
    If PE_CBool(Script_Property(6)) = True Then
        Content = ScriptHtml(Content, "Tr", 3)
    End If
    If PE_CBool(Script_Property(7)) = True Then
        Content = ScriptHtml(Content, "Td", 3)
    End If
    If PE_CBool(Script_Property(8)) = True Then
        Content = ScriptHtml(Content, "Span", 3)
    End If
    If PE_CBool(Script_Property(9)) = True Then
        Content = ScriptHtml(Content, "Img", 1)
    End If
    If PE_CBool(Script_Property(10)) = True Then
        Content = ScriptHtml(Content, "Font", 3)
    End If
    If PE_CBool(Script_Property(11)) = True Then
        Content = ScriptHtml(Content, "A", 3)
    End If
    If PE_CBool(Script_Property(12)) = True Then
        Content = nohtml(Content)
    End If
    FilterScript = Content
End Function

'**************************************************
'函数名：ZeroToEmpty
'作  用：判断字符串是否等于"0"，如果是将字符串置为空，用于JS生成处理
'参  数：str ---- 待处理的字符串
'**************************************************
Function ZeroToEmpty(str)
    If str = "0" Then
        ZeroToEmpty = ""
    Else
        ZeroToEmpty = str
    End If
End Function

Function URLDecode(enStr)
    Dim deStr
    Dim c, i, v
    deStr = ""
    For i = 1 To Len(enStr)
        c = Mid(enStr, i, 1)
        If c = "%" Then
            v = eval("&h" + Mid(enStr, i + 1, 2))
            If v < 128 Then
                deStr=deStr&chr(v)
                i = i + 2
            Else
                If isvalidhex(Mid(enStr, i, 3)) Then
                    If isvalidhex(Mid(enStr, i + 3, 3)) Then
                        v = eval("&h" + Mid(enStr, i + 1, 2) + Mid(enStr, i + 4, 2))
                        deStr=deStr&chr(v)
                        i = i + 5
                    Else
                        v = eval("&h" + Mid(enStr, i + 1, 2) + CStr(Hex(Asc(Mid(enStr, i + 3, 1)))))
                        deStr=deStr&chr(v)
                        i = i + 3
                    End If
                Else
                    destr=destr&c
                End If
            End If
        Else
            If c = "+" Then
                deStr=deStr&" "
            Else
                deStr=deStr&c
            End If
        End If
    Next
    URLDecode = deStr
End Function

Function isIP(strng)
    regEx.Pattern = "^(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])$"
    isIP = regEx.Test(strng)
End Function

Function EncodeIP(sip)
    Dim strIP
    strIP = Split(sip, ".")
    If UBound(strIP) < 3 Then
        EncodeIP = 0
        Exit Function
    End If
    If IsNumeric(strIP(0)) = False Or IsNumeric(strIP(1)) = False Or IsNumeric(strIP(2)) = False Or IsNumeric(strIP(3)) = False Then
        EncodeIP = 0
    Else
        EncodeIP = CDbl(strIP(0)) * 256 * 256 * 256 + CLng(strIP(1)) * 256 * 256 + CLng(strIP(2)) * 256 + CLng(strIP(3)) - 1
    End If
End Function

Function DecodeIP(sip)
    Dim s1, s21, s2, s31, s3, s4
    sip = sip + 1
    s1 = Int(sip / 256 / 256 / 256)
    s21 = s1 * 256 * 256 * 256
    s2 = Int((sip - s21) / 256 / 256)
    s31 = s2 * 256 * 256 + s21
    s3 = Int((sip - s31) / 256)
    s4 = sip - s3 * 256 - s31
    DecodeIP = CStr(s1) + "." + CStr(s2) + "." + CStr(s3) + "." + CStr(s4)
End Function


Function FilterBadTag(strContent, Inputer)
    Dim rsAdmin
    Set rsAdmin = Conn.Execute("select AdminName from PE_Admin where UserName='" & Inputer & "'")
    If rsAdmin.bof And rsAdmin.EOF Then
        FilterBadTag = FilterJS(strContent)
    Else
        FilterBadTag = strContent
    End If
    rsAdmin.Close
    Set rsAdmin = Nothing
End Function

%>
