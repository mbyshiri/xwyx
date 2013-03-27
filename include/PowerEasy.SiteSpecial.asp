<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Private Sub GetHtml_Special()
    ChannelID = 0
    strHtml = PE_Replace(strHtml, "{$SpecialID}", SpecialID)
    Call ReplaceCommonLabel
    strHtml = PE_Replace(strHtml, "{$MenuJS}", "")
    strHtml = PE_Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(SkinID))

    strNavPath = XmlText("BaseText", "Nav", "您现在的位置：") & "&nbsp;<a class='LinkPath' href='" & SiteUrl & "'>" & SiteName & "</a>&nbsp;" & strNavLink & "&nbsp;" & "[专题]" & SpecialName

    strHtml = Replace(strHtml, "{$PageTitle}", SiteTitle & " >> " & "[专题]" & SpecialName)
    strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)

    strHtml = PE_Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = PE_Replace(strHtml, "{$ShowPath}", ShowPath())
    strHtml = PE_Replace(strHtml, "{$SpecialID}", SpecialID)
    strHtml = PE_Replace(strHtml, "{$SpecialName}", SpecialName)
    strHtml = PE_Replace(strHtml, "{$SpecialPicUrl}", SpecialPicUrl)
    strHtml = PE_Replace(strHtml, "{$Readme}", ReadMe)
    strHtml = CustomContent("Special", Custom_Content_Special, strHtml)
    totalPut = 0
    
    Dim PE_Article, PE_Soft, PE_Photo, PE_Shop
    Set PE_Article = New Article
    PE_Article.Init
    strHtml = PE_Article.GetCustomFromTemplate(strHtml)
    strHtml = PE_Article.GetPicFromTemplate(strHtml)
    strHtml = PE_Article.GetListFromTemplate(strHtml)
    strHtml = PE_Article.GetSlidePicFromTemplate(strHtml)
    Set PE_Article = Nothing

    Set PE_Soft = New Soft
    PE_Soft.Init
    strHtml = PE_Soft.GetCustomFromTemplate(strHtml)
    strHtml = PE_Soft.GetPicFromTemplate(strHtml)
    strHtml = PE_Soft.GetListFromTemplate(strHtml)
    strHtml = PE_Soft.GetSlidePicFromTemplate(strHtml)
    Set PE_Soft = Nothing

    Set PE_Photo = New Photo
    PE_Photo.Init
    strHtml = PE_Photo.GetCustomFromTemplate(strHtml)
    strHtml = PE_Photo.GetPicFromTemplate(strHtml)
    strHtml = PE_Photo.GetListFromTemplate(strHtml)
    strHtml = PE_Photo.GetSlidePicFromTemplate(strHtml)
    Set PE_Photo = Nothing
    
    Set PE_Shop = New Product
    PE_Shop.Init
    strHtml = PE_Shop.GetCustomFromTemplate(strHtml)
    strHtml = PE_Shop.GetPicFromTemplate(strHtml)
    strHtml = PE_Shop.GetListFromTemplate(strHtml)
    strHtml = PE_Shop.GetSlidePicFromTemplate(strHtml)
    Set PE_Shop = Nothing
    Dim strPath
    strPath = InstallDir & "Special/" & SpecialDir
    If FileExt_SiteSpecial = ".asp" Then
        If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个内容", False))
        If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个内容", False))
    Else
        If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage_Html(strPath, 0, FileExt_SiteSpecial, "", totalPut, MaxPerPage, CurrentPage, True, True, "个内容"))
        If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en_Html(strPath, 0, FileExt_SiteSpecial, "", totalPut, MaxPerPage, CurrentPage, True, True, "个内容"))
    End If
    strHtml = Replace(strHtml, "{$Rss}", "")
    
End Sub

Private Sub GetHtml_SpecialList()
    ChannelID = 0
    Dim PE_Article, PE_Soft, PE_Photo, PE_Shop
    Call ReplaceCommonLabel
    strHtml = PE_Replace(strHtml, "{$MenuJS}", "")
    strHtml = PE_Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(SkinID))
    strHtml = CustomContent("Special", Custom_Content_Special, strHtml)
    totalPut = 0

    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", ShowPath())
    If EnableRss = True Then
        strHtml = Replace(strHtml, "{$Rss}", "<a href='" & InstallDir & "Rss.asp?ChannelID=0&SpecialID=" & SpecialID & "' Target='_blank'><img src='" & strInstallDir & "images/rss.gif' border=0></a>")
    Else
        strHtml = Replace(strHtml, "{$Rss}", "")
    End If
    strHtml = PE_Replace(strHtml, "{$GetAllSpecial}", GetAllSpecial)
    
    
    Set PE_Article = New Article
    PE_Article.Init
    strHtml = PE_Article.GetCustomFromTemplate(strHtml)
    strHtml = PE_Article.GetPicFromTemplate(strHtml)
    strHtml = PE_Article.GetListFromTemplate(strHtml)
    strHtml = PE_Article.GetSlidePicFromTemplate(strHtml)
    Set PE_Article = Nothing

    Set PE_Soft = New Soft
    PE_Soft.Init
    strHtml = PE_Soft.GetCustomFromTemplate(strHtml)
    strHtml = PE_Soft.GetPicFromTemplate(strHtml)
    strHtml = PE_Soft.GetListFromTemplate(strHtml)
    strHtml = PE_Soft.GetSlidePicFromTemplate(strHtml)
    Set PE_Soft = Nothing

    Set PE_Photo = New Photo
    PE_Photo.Init
    strHtml = PE_Photo.GetCustomFromTemplate(strHtml)
    strHtml = PE_Photo.GetPicFromTemplate(strHtml)
    strHtml = PE_Photo.GetListFromTemplate(strHtml)
    strHtml = PE_Photo.GetSlidePicFromTemplate(strHtml)
    Set PE_Photo = Nothing
    
    Set PE_Shop = New Product
    PE_Shop.Init
    strHtml = PE_Shop.GetCustomFromTemplate(strHtml)
    strHtml = PE_Shop.GetPicFromTemplate(strHtml)
    strHtml = PE_Shop.GetListFromTemplate(strHtml)
    strHtml = PE_Shop.GetSlidePicFromTemplate(strHtml)
    Set PE_Shop = Nothing
   
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个专题", False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个专题", False))
    
End Sub

Private Sub GetSpecial()
    Dim tSpecial
    Set tSpecial = Conn.Execute("select SpecialID,SpecialName,Readme,SkinID,TemplateID,SpecialDir,SpecialPicUrl,Custom_Content,MaxPerPage from PE_Special where ChannelID=0 and SpecialID=" & SpecialID & "")
    If tSpecial.BOF And tSpecial.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的专题！</li>"
    Else
        SpecialName = tSpecial("SpecialName")
        SkinID = tSpecial("SkinID")
        TemplateID = tSpecial("TemplateID")
        SpecialDir = tSpecial("SpecialDir")
        SpecialPicUrl = tSpecial("SpecialPicUrl")
        Custom_Content_Special = tSpecial("Custom_Content")
        MaxPerPage_Special = tSpecial("MaxPerPage")
        ReadMe = tSpecial("Readme")
        If FileExt_SiteSpecial <> ".asp" Then
            strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;" & "<font color=blue>[专题]</font>" & "<a href='" & InstallDir & "Special/" & SpecialDir & "/Index" & FileExt_SiteSpecial & "'>" & SpecialName & "</a>"
        Else
            strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;" & "<font color=blue>[专题]</font>" & "<a href='" & InstallDir & "ShowSpecial.asp?SpecialID=" & SpecialID & "'>" & SpecialName & "</a>"
        End If
        strPageTitle = strPageTitle & " >> " & "[专题]" & SpecialName
    End If
    tSpecial.Close
    Set tSpecial = Nothing
End Sub


Private Function GetAllSpecial()
    Dim sqlSpecial, rsSpecial, strSpecial, iCount
    iCount = 0
    sqlSpecial = "select SpecialID,SpecialName,SpecialDir from PE_Special where ChannelID=0 order by OrderID"
    Set rsSpecial = Server.CreateObject("ADODB.Recordset")
    rsSpecial.Open sqlSpecial, Conn, 1, 1
    If rsSpecial.BOF And rsSpecial.EOF Then
        totalPut = 0
        strSpecial = "<li>没有任何专题</li>"
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
        If FileExt_SiteSpecial <> ".asp" Then
            strSpecial = strSpecial & "<a href='" & InstallDir & "Special/" & rsSpecial(2) & "/Index" & FileExt_SiteSpecial & "'>" & rsSpecial(1) & "</a>"
        Else
            strSpecial = strSpecial & "<a href='" & InstallDir & "ShowSpecial.asp?SpecialID=" & rsSpecial(0) & "'>" & rsSpecial(1) & "</a>"
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
