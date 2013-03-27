<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'定义专题设置相关的变量
Private SpecialID, SpecialName, SpecialDir, SpecialPicUrl, MaxPerPage_Special, Custom_Content_Special

Private Sub GetSpecial()
    Dim tSpecial
    Set tSpecial = Conn.Execute("select SpecialID,SpecialName,Readme,SkinID,TemplateID,SpecialDir,SpecialPicUrl,MaxPerPage,Custom_Content from PE_Special where SpecialID=" & SpecialID & "")
    If tSpecial.BOF And tSpecial.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的专题！</li>"
    Else
        SpecialName = tSpecial("SpecialName")
        SkinID = tSpecial("SkinID")
        ReadMe = tSpecial("Readme")
        TemplateID = tSpecial("TemplateID")
        SpecialDir = tSpecial("SpecialDir")
        SpecialPicUrl = tSpecial("SpecialPicUrl")
        Custom_Content_Special = tSpecial("Custom_Content")
        MaxPerPage_Special = tSpecial("MaxPerPage")
        If UseCreateHTML = 1 Or UseCreateHTML = 3 Then
            strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;<font color=blue>[专题]</font><a href='" & ChannelUrl & "/Special/" & SpecialDir & "/Index" & FileExt_List & "'>" & SpecialName & "</a>"
        Else
            strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;<font color=blue>[专题]</font><a href='" & ChannelUrl_ASPFile & "/ShowSpecial.asp?SpecialID=" & SpecialID & "'>" & SpecialName & "</a>"
        End If
        strPageTitle = strPageTitle & " >> " & "[专题]" & SpecialName
    End If
    tSpecial.Close
    Set tSpecial = Nothing
End Sub

Private Function GetAllSpecial()
    Dim sqlSpecial, rsSpecial, strSpecial, iCount
    iCount = 0
    sqlSpecial = "select SpecialID,SpecialName,SpecialDir from PE_Special where ChannelID=" & ChannelID & " order by OrderID"
    Set rsSpecial = Server.CreateObject("ADODB.Recordset")
    rsSpecial.Open sqlSpecial, Conn, 1, 1
    If rsSpecial.BOF And rsSpecial.EOF Then
        totalPut = 0
        strSpecial = "<li>没有任何专题！</li>"
        rsSpecial.Close
        Set rsSpecial = Nothing
        GetAllSpecial = strSpecial
        Exit Function
    End If
    totalPut = rsSpecial.RecordCount
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
            rsSpecial.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If
    strSpecial = "<table width='100%' cellpadding='0' cellspacing='0'>"
    Do While Not rsSpecial.EOF
        If iCount Mod 2 = 0 Then
            strSpecial = strSpecial & "<tr class='listbg'>"
        Else
            strSpecial = strSpecial & "<tr class='listbg2'>"
        End If
        strSpecial = strSpecial & "<td>·&nbsp;"
        If UseCreateHTML > 0 Then
            strSpecial = strSpecial & "<a href='" & ChannelUrl & "/Special/" & rsSpecial(2) & "/Index" & FileExt_List & "'>" & rsSpecial(1) & "</a>"
        Else
            strSpecial = strSpecial & "<a href='" & ChannelUrl_ASPFile & "/ShowSpecial.asp?SpecialID=" & rsSpecial(0) & "'>" & rsSpecial(1) & "</a>"
        End If
        strSpecial = strSpecial & "</td></tr>"
        rsSpecial.MoveNext
        iCount = iCount + 1
        If iCount >= MaxPerPage Then Exit Do
    Loop
    strSpecial = strSpecial & "</table>"

    GetAllSpecial = strSpecial
End Function

%>
