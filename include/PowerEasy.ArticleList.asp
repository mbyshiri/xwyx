<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************
Function ShowArticleList(ByVal str1, UseType)
    Dim strtmp, rsArticle, sqlArticle, i, j
    Dim strTemp, arrTemp
    Dim iType, iDate, iLink, iNum, iorder, iCol, iHeight, iWidth
    If str1 = "" Then
        ShowArticleList = ""
        Exit Function
    End If

    arrTemp = Split(str1, ",")
    If UBound(arrTemp) < 10 Then
        ShowArticleList = "函数式标签：{$AuthorArticleList()}的参数个数不对。请检查模板中的此标签。"
        Exit Function
    Else
        If UBound(arrTemp) = 12 Then
            iWidth = PE_CLng(arrTemp(11))
            iHeight = PE_CLng(arrTemp(12))
        Else
            iWidth = 130
            iHeight = 90
        End If
    End If
    If arrTemp(1) = 0 Then
        Exit Function
    End If

    iType = PE_CLng(arrTemp(5))
    iNum = PE_CLng(arrTemp(8))
    iorder = PE_CLng(arrTemp(9))
    If arrTemp(10) = "" Or arrTemp(10) < 1 Then
        iCol = 1
    Else
        iCol = arrTemp(10)
    End If
    strtmp = strtmp & "<Table width='100%'><tr><td>"

    sqlArticle = "select * from PE_" & arrTemp(2) & " where"
    If UseType = 1 Then
         sqlArticle = sqlArticle & " Author='" & arrTemp(0) & "'"
    Else
         sqlArticle = sqlArticle & " CopyFrom='" & arrTemp(0) & "'"
    End If
    sqlArticle = sqlArticle & (" and ChannelID=" & arrTemp(1) & " and Status=3 and Deleted=" & PE_False)
    If TimeData <> "0" Then
        sqlArticle = sqlArticle & " and DateDiff(" & PE_DatePart_D & ",UpdateTime,'" & TimeData & "')=0"
    End If
    Select Case iorder
    Case 0
        sqlArticle = sqlArticle & " order by UpdateTime desc"
    Case 1
        sqlArticle = sqlArticle & " order by UpdateTime"
    Case 2
        sqlArticle = sqlArticle & " order by Hits desc,UpdateTime desc"
    End Select
    Set rsArticle = Server.CreateObject("ADODB.Recordset")
    rsArticle.Open sqlArticle, Conn, 1, 1
    If rsArticle.BOF And rsArticle.EOF Then
        totalPut = 0
        If TimeData <> "0" Then
            strtmp = strtmp & Replace(XmlText("ShowSource", "ShowAuthorList/NoDay", "&nbsp;&nbsp;<h3>{$AuthorName}本日未发表任何作品</h3>"), "{$AuthorName}", arrTemp(0))
        Else
            strtmp = strtmp & Replace(XmlText("ShowSource", "ShowAuthorList/NoArticle", "&nbsp;&nbsp;<h3>{$AuthorName}在本频道还未收录任何作品</h3>"), "{$AuthorName}", arrTemp(0))
        End If
    Else
        totalPut = rsArticle.RecordCount
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                    rsArticle.Move (CurrentPage - 1) * MaxPerPage
                Else
                    CurrentPage = 1
                End If
        End If
        i = 0
        Dim strShowDetal
        strShowDetal = XmlText("ShowSource", "ShowAuthorList/ShowDetal", "点击这里浏览具体内容>>>")

        Select Case iType
        Case 1
            Do While Not rsArticle.EOF
                If PE_CBool(arrTemp(6)) Then strtmp = strtmp & ("<p>" & Year(rsArticle("UpdateTime")) & strYear & Month(rsArticle("UpdateTime")) & strMonth & Day(rsArticle("UpdateTime")) & strDay & "</p>")
                Select Case arrTemp(2)
                Case "Article"
                    strtmp = strtmp & "<table width='100%'><tr>"
                    If rsArticle("DefaultPicUrl") <> "" And Not IsNull(rsArticle("DefaultPicUrl")) Then
                        If Left(rsArticle("DefaultPicUrl"), 4) = "http" Or Left(rsArticle("DefaultPicUrl"), 1) = "/" Then
                            strtmp = strtmp & "<td width=150 align=center><img src='" & rsArticle("DefaultPicUrl") & "'  width='" & iWidth & "' height='" & iHeight & "' border='0'></td>"
                        Else
                            strtmp = strtmp & "<td width=150 align=center><img src='" & strInstallDir & arrTemp(3) & "/" & arrTemp(4) & "/" & rsArticle("DefaultPicUrl") & "'  width='" & iWidth & "' height='" & iHeight & "' border='0'></td>"
                        End If
                    End If
                    strtmp = strtmp & "<td>"
                    strtmp = strtmp & ("<h4>" & ReplaceText(rsArticle("Title"), 2) & "</h4>&nbsp;&nbsp;")
                    strtmp = strtmp & GetSubStr(ReplaceText(nohtml(rsArticle("Content")), 1), iNum, True)
                    If PE_CBool(arrTemp(7)) Then
                       strtmp = strtmp & ("<div align='right'><a href='")
                       strtmp = strtmp & GetArticleUrl(GetClassFild(rsArticle("ClassID"), 4), GetClassFild(rsArticle("ClassID"), 3), rsArticle("UpdateTime"), rsArticle("ArticleID"), GetClassFild(rsArticle("ClassID"), 2), rsArticle("InfoPurview"), rsArticle("InfoPoint"))
                       strtmp = strtmp & ("'>" & strShowDetal & "</a></div>")
                    End If
                    strtmp = strtmp & "</td></tr></table>"
                    strtmp = strtmp & "<hr>"
                Case "Soft"
                    strtmp = strtmp & "<table width='100%'><tr>"
                    If rsArticle("SoftPicUrl") <> "" And Not IsNull(rsArticle("SoftPicUrl")) Then
                        If Left(rsArticle("SoftPicUrl"), 4) = "http" Or Left(rsArticle("SoftPicUrl"), 1) = "/" Then
                            strtmp = strtmp & "<td width=150 align=center><img src='" & rsArticle("SoftPicUrl") & "'  width='" & iWidth & "' height='" & iHeight & "' border='0' align='left'></td>"
                        Else
                            strtmp = strtmp & ("<td width=150 align=center><img src='")
                            strtmp = strtmp & GetSoftUrl(GetClassFild(rsArticle("ClassID"), 4), GetClassFild(rsArticle("ClassID"), 3), rsArticle("UpdateTime"), rsArticle("SoftID"))
                            strtmp = strtmp & ("'  width='" & iWidth & "' height='" & iHeight & "' border='0' align='left'></td>")
                        End If
                    End If
                    strtmp = strtmp & "<td>"
                    strtmp = strtmp & ("<h4>" & ReplaceText(rsArticle("SoftName"), 2) & "</h4>&nbsp;&nbsp;")
                    If rsArticle("SoftIntro") = "" Then
                        strtmp = strtmp & rsArticle("Keyword")
                    Else
                        strtmp = strtmp & GetSubStr(ReplaceText(nohtml(rsArticle("SoftIntro")), 1), iNum, True)
                    End If
                    If PE_CBool(arrTemp(7)) Then strtmp = strtmp & ("<div align='right'><a href='" & strInstallDir & arrTemp(3) & "/ShowSoft.asp?SoftID=" & rsArticle("SoftID") & "' Target=""_blank"">" & strShowDetal & "</a></div>")
                    strtmp = strtmp & "</td></tr></table>"
                    strtmp = strtmp & "<hr>"
                Case "Photo"
                    strtmp = strtmp & "<table width='100%'><tr>"
                    If rsArticle("PhotoThumb") <> "" And Not IsNull(rsArticle("PhotoThumb")) Then
                        If Left(rsArticle("PhotoThumb"), 4) = "http" Or Left(rsArticle("PhotoThumb"), 1) = "/" Then
                            strtmp = strtmp & "<td width=150 align=center><img src='" & rsArticle("PhotoThumb") & "'  width='" & iWidth & "' height='" & iHeight & "' border='0' align='left'></td>"
                        Else
                            strtmp = strtmp & "<td width=150 align=center><img src='" & strInstallDir & arrTemp(3) & "/" & arrTemp(4) & "/" & rsArticle("PhotoThumb") & "'  width='" & iWidth & "' height='" & iHeight & "' border='0' align='left'></td>"
                        End If
                    End If
                    strtmp = strtmp & "<td>"
                    strtmp = strtmp & ("<h4>" & ReplaceText(rsArticle("PhotoName"), 2) & "</h4>&nbsp;&nbsp;")
                    If rsArticle("PhotoIntro") = "" Then
                        strtmp = strtmp & rsArticle("Keyword")
                    Else
                        strtmp = strtmp & GetSubStr(ReplaceText(nohtml(rsArticle("PhotoIntro")), 1), iNum, True)
                    End If
                    If PE_CBool(arrTemp(7)) Then
                        strtmp = strtmp & ("<div align='right'><a href='")
                        strtmp = strtmp & GetPhotoUrl(GetClassFild(rsArticle("ClassID"), 4), GetClassFild(rsArticle("ClassID"), 3), rsArticle("UpdateTime"), rsArticle("PhotoID"), GetClassFild(rsArticle("ClassID"), 2), rsArticle("InfoPurview"), rsArticle("InfoPoint"))
                        strtmp = strtmp & ("' Target=""_blank"">" & strShowDetal & "</a></div>")
                    End If
                    strtmp = strtmp & "</td></tr></table>"
                    strtmp = strtmp & "<hr>"
                End Select
                rsArticle.MoveNext
                i = i + 1
                If i >= MaxPerPage Then Exit Do
            Loop
        Case 2
            Do While Not rsArticle.EOF
                Select Case arrTemp(2)
                Case "Article"
                    If PE_CBool(arrTemp(7)) Then
                        strtmp = strtmp & ("<li><a href='" & GetArticleUrl(GetClassFild(rsArticle("ClassID"), 4), GetClassFild(rsArticle("ClassID"), 3), rsArticle("UpdateTime"), rsArticle("ArticleID"), GetClassFild(rsArticle("ClassID"), 2), rsArticle("InfoPurview"), rsArticle("InfoPoint")) & "' Target=""_blank"">" & GetSubStr(rsArticle("Title"), iNum, False) & "</a>")
                    Else
                        strtmp = strtmp & ("<li>" & GetSubStr(ReplaceText(rsArticle("Title"), 2), iNum, False))
                    End If
                Case "Soft"
                    If PE_CBool(arrTemp(7)) Then
                        strtmp = strtmp & ("<li><a href='" & GetSoftUrl(GetClassFild(rsArticle("ClassID"), 4), GetClassFild(rsArticle("ClassID"), 3), rsArticle("UpdateTime"), rsArticle("SoftID")) & "' Target=""_blank"">" & GetSubStr(rsArticle("SoftName"), iNum, False) & "</a>")
                    Else
                        strtmp = strtmp & ("<li>" & GetSubStr(ReplaceText(rsArticle("SoftName"), 2), iNum, False))
                    End If
                Case "Photo"
                    If PE_CBool(arrTemp(7)) Then
                        strtmp = strtmp & ("<li><a href='" & GetPhotoUrl(GetClassFild(rsArticle("ClassID"), 4), GetClassFild(rsArticle("ClassID"), 3), rsArticle("UpdateTime"), rsArticle("PhotoID"), GetClassFild(rsArticle("ClassID"), 2), rsArticle("InfoPurview"), rsArticle("InfoPoint")) & "' Target=""_blank"">" & GetSubStr(rsArticle("PhotoName"), iNum, False) & "</a>")
                    Else
                        strtmp = strtmp & ("<li>" & GetSubStr(ReplaceText(rsArticle("PhotoName"), 2), iNum, False))
                    End If
                End Select
                If PE_CBool(arrTemp(6)) Then strtmp = strtmp & ("[" & Year(rsArticle("UpdateTime")) & strYear & Month(rsArticle("UpdateTime")) & strMonth & Day(rsArticle("UpdateTime")) & strDay & "]")
                strtmp = strtmp & "</li>"
                rsArticle.MoveNext
                i = i + 1
                If i >= MaxPerPage Then Exit Do
            Loop
        Case 3
            j = 1
            strtmp = strtmp & "<table class=ItemList><tr>"
            Do While Not rsArticle.EOF
                Select Case arrTemp(2)
                Case "Article"
                    strtmp = strtmp & "<td><table width='100%'><tr><td align='center' valign='top' class='ArticlePic'>"
                    strtmp = strtmp & "<a href='" & strInstallDir & arrTemp(3) & "/ShowArticle.asp?ArticleID=" & rsArticle("ArticleID") & "' Target=""_blank"">"
                    If rsArticle("DefaultPicUrl") = "" Or IsNull(rsArticle("DefaultPicUrl")) Then
                        strtmp = strtmp & "<img src='" & strInstallDir & "images/nopic.gif'  width='" & iWidth & "' height='" & iHeight & "' border='0'></a></td></tr>"
                    Else
                        If Left(rsArticle("DefaultPicUrl"), 4) = "http" Or Left(rsArticle("DefaultPicUrl"), 1) = "/" Then
                            strtmp = strtmp & "<img src='" & rsArticle("DefaultPicUrl") & "'  width='" & iWidth & "' height='" & iHeight & "' border='0'></a></td></tr>"
                        Else
                            strtmp = strtmp & "<img src='" & strInstallDir & arrTemp(3) & "/" & arrTemp(4) & "/" & rsArticle("DefaultPicUrl") & "'  width='" & iWidth & "' height='" & iHeight & "' border='0'></a></td></tr>"
                        End If
                    End If
                    strtmp = strtmp & "<tr><td align='center' class='ArticleName'><a href='" & GetArticleUrl(GetClassFild(rsArticle("ClassID"), 4), GetClassFild(rsArticle("ClassID"), 3), rsArticle("UpdateTime"), rsArticle("ArticleID"), GetClassFild(rsArticle("ClassID"), 2), rsArticle("InfoPurview"), rsArticle("InfoPoint")) & "' target='_blank'>" & GetSubStr(ReplaceText(rsArticle("Title"), 2), iNum, False) & "</a>"
                    If PE_CBool(arrTemp(6)) Then strtmp = strtmp & "<br>" & ("[" & Year(rsArticle("UpdateTime")) & strYear & Month(rsArticle("UpdateTime")) & strMonth & Day(rsArticle("UpdateTime")) & strDay & "]")
                    strtmp = strtmp & "</td></tr></table>"
                Case "Soft"
                    strtmp = strtmp & "<td><table width='100%'><tr><td align='center' valign='top' class='SoftPic'>"
                    strtmp = strtmp & "<a href='" & strInstallDir & arrTemp(3) & "/ShowSoft.asp?SoftID=" & rsArticle("SoftID") & "' Target=""_blank"">"
                    If rsArticle("SoftPicUrl") = "" Or IsNull(rsArticle("SoftPicUrl")) Then
                        strtmp = strtmp & "<img src='" & strInstallDir & "images/nopic.gif'  width='" & iWidth & "' height='" & iHeight & "' border='0'></a></td></tr>"
                    Else
                        If Left(rsArticle("SoftPicUrl"), 4) = "http" Or Left(rsArticle("SoftPicUrl"), 1) = "/" Then
                            strtmp = strtmp & "<img src='" & rsArticle("SoftPicUrl") & "'  width='" & iWidth & "' height='" & iHeight & "' border='0'></a></td></tr>"
                        Else
                            strtmp = strtmp & "<img src='" & strInstallDir & arrTemp(3) & "/" & rsArticle("SoftPicUrl") & "'  width='" & iWidth & "' height='" & iHeight & "' border='0'></a></td></tr>"
                        End If
                    End If
                    strtmp = strtmp & "<tr><td align='center' class='SoftName'><a href='" & GetSoftUrl(GetClassFild(rsArticle("ClassID"), 4), GetClassFild(rsArticle("ClassID"), 3), rsArticle("UpdateTime"), rsArticle("SoftID")) & "' target='_blank'>" & GetSubStr(ReplaceText(rsArticle("SoftName"), 2), iNum, False) & "</a>"
                    If PE_CBool(arrTemp(6)) Then strtmp = strtmp & "<br>" & ("[" & Year(rsArticle("UpdateTime")) & strYear & Month(rsArticle("UpdateTime")) & strMonth & Day(rsArticle("UpdateTime")) & strDay & "]")
                    strtmp = strtmp & "</td></tr></table>"
                Case "Photo"
                    strtmp = strtmp & "<td><table width='100%'><tr><td align='center' valign='top' class='PhotoPic'>"
                    strtmp = strtmp & "<a href='" & strInstallDir & arrTemp(3) & "/ShowPhoto.asp?PhotoID=" & rsArticle("PhotoID") & "' Target=""_blank"">"
                    If rsArticle("PhotoThumb") = "" Or IsNull(rsArticle("PhotoThumb")) Then
                        strtmp = strtmp & "<img src='" & strInstallDir & "images/nopic.gif'  width='" & iWidth & "' height='" & iHeight & "' border='0'></a></td></tr>"
                    Else
                        If Left(rsArticle("PhotoThumb"), 4) = "http" Or Left(rsArticle("PhotoThumb"), 1) = "/" Then
                            strtmp = strtmp & "<img src='" & rsArticle("PhotoThumb") & "'  width='" & iWidth & "' height='" & iHeight & "' border='0'></a></td></tr>"
                        Else
                            strtmp = strtmp & "<img src='" & strInstallDir & arrTemp(3) & "/" & arrTemp(4) & "/" & rsArticle("PhotoThumb") & "'  width='" & iWidth & "' height='" & iHeight & "' border='0'></a></td></tr>"
                        End If
                    End If
                    strtmp = strtmp & "<tr><td align='center' class='PhotoName'><a href='" & GetPhotoUrl(GetClassFild(rsArticle("ClassID"), 4), GetClassFild(rsArticle("ClassID"), 3), rsArticle("UpdateTime"), rsArticle("PhotoID"), GetClassFild(rsArticle("ClassID"), 2), rsArticle("InfoPurview"), rsArticle("InfoPoint")) & "' target='_blank'>" & GetSubStr(ReplaceText(rsArticle("PhotoName"), 2), iNum, False) & "</a>"
                    If PE_CBool(arrTemp(6)) Then strtmp = strtmp & "<br>" & ("[" & Year(rsArticle("UpdateTime")) & strYear & Month(rsArticle("UpdateTime")) & strMonth & Day(rsArticle("UpdateTime")) & strDay & "]")
                    strtmp = strtmp & "</td></tr></table>"
                End Select
                rsArticle.MoveNext
                i = i + 1
                j = j + 1
                If j > Int(iCol) Then
                    strtmp = strtmp & "</td></tr><tr>"
                    j = 1
                Else
                    strtmp = strtmp & "</td>"
                End If
                If i >= MaxPerPage Then Exit Do
            Loop
            strtmp = strtmp & "</tr></table>"
        End Select
    End If
    rsArticle.Close
    Set rsArticle = Nothing
    strtmp = strtmp & "</td></tr></table>"
    ShowArticleList = strtmp
    strtmp = ""
End Function
%>
