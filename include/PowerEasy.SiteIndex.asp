<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************


Public Sub GetHTML_SiteIndex()
    Dim PE_Content, strPath
    ChannelID = 0
    strHtml = GetTemplate(0, 1, 0)
    Call ReplaceCommonLabel
    
    strNavPath = XmlText("BaseText", "Nav", "您现在的位置：") & "&nbsp;<a class='LinkPath' href='" & SiteUrl & "'>" & SiteName & "</a>&nbsp;" & strNavLink & "&nbsp;" & "首页"

    strHtml = Replace(strHtml, "{$PageTitle}", SiteTitle & " >> " & "首页")
    strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)

    strHtml = Replace(strHtml, "{$MenuJS}", GetMenuJS("", False))
    strHtml = Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))
    If EnableRss = True Then
        If FileExt_SiteIndex = ".asp" Then
            strHtml = Replace(strHtml, "{$Rss}", "<a href='" & InstallDir & "Rss.asp' Target='_blank'><img src='" & InstallDir & "images/rss.gif' border=0 alt='" & SiteTitle & "Rss 2.0'></a>")
        Else
            strHtml = Replace(strHtml, "{$Rss}", "<a href='" & InstallDir & "xml/Rss.xml' Target='_blank'><img src='" & InstallDir & "images/rss.gif' border=0 alt='" & SiteTitle & "Rss 2.0'></a>")
        End If
    Else
        strHtml = Replace(strHtml, "{$Rss}", "")
    End If
    If EnableWap = True Then
        strHtml = Replace(strHtml, "{$Wap}", "<img src='" & InstallDir & "images/Wap.gif' border=0 alt='" & SiteTitle & XmlText("Wap", "Alt", "手机WAP浏览支持") & "' onClick=""window.open('" & XmlText("Wap", "Domain", strInstallDir & "Wap") & "?ReadMe=Yes', 'Wap', 'width=" & XmlText("Wap", "ImgWidth", "160") & ",height=" & XmlText("Wap", "ImgHight", "257") & ",resizable=0,scrollbars=no');"" style=""cursor:hand;"">")
    Else
        strHtml = Replace(strHtml, "{$Wap}", "")
    End If
    
    Set PE_Content = New Article
    PE_Content.Init
    strHtml = PE_Content.GetCustomFromTemplate(strHtml)
    strHtml = PE_Content.GetPicFromTemplate(strHtml)
    strHtml = PE_Content.GetListFromTemplate(strHtml)
    strHtml = PE_Content.GetSlidePicFromTemplate(strHtml)
    Set PE_Content = Nothing
    Set PE_Content = New Soft
    PE_Content.Init
    strHtml = PE_Content.GetCustomFromTemplate(strHtml)
    strHtml = PE_Content.GetPicFromTemplate(strHtml)
    strHtml = PE_Content.GetListFromTemplate(strHtml)
    strHtml = PE_Content.GetSlidePicFromTemplate(strHtml)
    Set PE_Content = Nothing
    Set PE_Content = New Photo
    PE_Content.Init
    strHtml = PE_Content.GetCustomFromTemplate(strHtml)
    strHtml = PE_Content.GetPicFromTemplate(strHtml)
    strHtml = PE_Content.GetListFromTemplate(strHtml)
    strHtml = PE_Content.GetSlidePicFromTemplate(strHtml)
    Set PE_Content = Nothing
    Set PE_Content = New Product
    PE_Content.Init
    strHtml = PE_Content.GetCustomFromTemplate(strHtml)
    strHtml = PE_Content.GetPicFromTemplate(strHtml)
    strHtml = PE_Content.GetListFromTemplate(strHtml)
    strHtml = PE_Content.GetSlidePicFromTemplate(strHtml)
    Set PE_Content = Nothing
End Sub
%>
