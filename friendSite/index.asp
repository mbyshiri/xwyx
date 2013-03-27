<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.Cache.asp"-->
<!--#include file="../Include/PowerEasy.Channel.asp"-->
<!--#include file="../Include/PowerEasy.Common.Front.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

PageTitle = "友情链接"
Dim SpecialID
KindType = PE_CLng(Trim(Request("KindType")))
LinkType = PE_CLng(Trim(Request("LinkType")))
KindID = PE_CLng(Trim(Request("KindID")))
SpecialID = PE_CLng(Trim(Request("SpecialID")))
strFileName = "Index.asp?KindType=" & KindType & "&LinkType=" & LinkType & "&KindID=" & KindID & "&SpecialID=" & SpecialID

If CurrentPage > 1 Or LinkType > 0 Or KindID > 0 Or SpecialID > 0 Or KindType > 0 Then
    Call GetFriendSiteList
Else
    If PE_Cache.CacheIsEmpty("FriendSite_Index") Then
        Call GetFriendSiteList
        PE_Cache.SetValue "FriendSite_Index", strHtml
    Else
        strHtml = PE_Cache.GetValue("FriendSite_Index")
    End If
End If
Response.Write strHtml
Call CloseConn

Sub GetFriendSiteList()
    Dim sqlLink, rsLink, strFriendSite, i
    Dim LinkSiteUrl
    
    ChannelID = 0
    
    strHtml = GetTemplate(ChannelID, 5, 0)
    
    Call ReplaceCommonLabel

    strNavPath = strNavPath & strNavLink & "&nbsp;" & PageTitle

    strHtml = Replace(strHtml, "{$PageTitle}", SiteTitle & " >> " & PageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)

    strHtml = Replace(strHtml, "{$MenuJS}", GetMenuJS("", False))
    strHtml = Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))
    
    Select Case KindType
    Case 1, 2
        strFriendSite = strFriendSite & "<table width='100%' cellSpacing=2 cellPadding=1 border=0 class='Channel_border'>"
        Dim rsFsKind, sqlFsKind
        sqlFsKind = "select KindID,KindName from PE_FsKind where KindType=" & KindType
        If KindID > 0 Then
            sqlFsKind = sqlFsKind & " and KindID=" & KindID
        End If
        sqlFsKind = sqlFsKind & " order by KindID"
        Set rsFsKind = Conn.Execute(sqlFsKind)
        If rsFsKind.BOF And rsFsKind.EOF Then
            strFriendSite = strFriendSite & ("<tr><td align='center'><b>" & XmlText("Site", "Errmsg/FriendSiteErr", "暂时还没有类别或专题的友情链接。") & "</b></td></tr>")
        Else
            Do While Not rsFsKind.EOF
                strFriendSite = strFriendSite & "<tr class='channel_title'><td colspan='10'><a class='LinkFriendSiteList' href='index.asp?KindType=" & KindType & "&KindID=" & rsFsKind("KindID") & "'>" & rsFsKind("KindName") & "</a></td></tr>"
                If KindID > 0 Then
                    sqlLink = "select ID,SiteName,SiteUrl,SiteIntro from PE_FriendSite where Passed=" & PE_True & " "
                Else
                    sqlLink = "select top 20 ID,SiteName,SiteUrl,SiteIntro from PE_FriendSite where Passed=" & PE_True & " "
                End If
                If rsFsKind("KindID") > 0 Then
                    If KindType = 1 Then
                        sqlLink = sqlLink & " and KindID=" & rsFsKind("KindID")
                    Else
                        sqlLink = sqlLink & " and SpecialID=" & rsFsKind("KindID")
                    End If
                End If
    
                sqlLink = sqlLink & " order by ID desc"
                Set rsLink = Conn.Execute(sqlLink)
                strFriendSite = strFriendSite & "<tr class='channl_tdbg'>"
                i = 0
                Do While Not rsLink.EOF
                    If EnableCountFriendSiteHits = True Then
                        LinkSiteUrl = InstallDir & "FriendSite/FriendSiteUrl.asp?ID=" & rsLink("ID")
                    Else
                        LinkSiteUrl = rsLink("SiteUrl")
                    End If
                    strFriendSite = strFriendSite & "<td><a class='LinkFriendSiteList' href='" & LinkSiteUrl & "' target='blank' title='" & rsLink("SiteIntro") & "'>" & rsLink("SiteName") & "</a></td>"
                    i = i + 1
                    If i Mod 4 = 0 Then
                        strFriendSite = strFriendSite & "</tr><tr class='channel_tdbg'>"
                    End If
                    rsLink.MoveNext
                Loop
                strFriendSite = strFriendSite & "</tr><tr><td class='main_shadow'></td></tr>"
                rsFsKind.MoveNext
            Loop
            rsLink.Close
            Set rsLink = Nothing
        End If
        rsFsKind.Close
        Set rsFsKind = Nothing
        strFriendSite = strFriendSite & "</table>"
    Case Else
        strFriendSite = strFriendSite & ("<table width='100%' cellSpacing=2 cellPadding=1 border=0><tr><td>" & XmlText("Site", "ShowFriendSiteList/t1", "分类显示：") & GetLinkType_Option & " " & GetFsKind_Option(1) & " " & GetFsKind_Option(2) & "</td></tr></table>")
        sqlLink = "select * from PE_FriendSite where Passed=" & PE_True & " "
        If LinkType > 0 Then
            sqlLink = sqlLink & " and LinkType=" & LinkType
        End If
        If KindID > 0 Then
            sqlLink = sqlLink & " and KindID=" & KindID
        End If
        If SpecialID > 0 Then
            sqlLink = sqlLink & " and SpecialID=" & SpecialID
        End If
        sqlLink = sqlLink & " order by ID desc"
        Set rsLink = Server.CreateObject("adodb.recordset")
        rsLink.Open sqlLink, Conn, 1, 1
        If rsLink.BOF And rsLink.EOF Then
            strFriendSite = strFriendSite & "<table width='100%' cellSpacing=2 cellPadding=1 border=0 class='Channel_border'><tr><td height='50'>" & XmlText("Site", "ShowFriendSiteList/t2", "共有 0 个友情链接") & "</td></tr></table>"
        Else
            totalPut = rsLink.RecordCount
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
                    rsLink.Move (CurrentPage - 1) * MaxPerPage
                Else
                    CurrentPage = 1
                End If
            End If
            
            i = 0
            strFriendSite = strFriendSite & "<table width='100%' cellSpacing=2 cellPadding=1 border=0 class='Channel_border'>"
            strFriendSite = strFriendSite & "<tr class='channel_title'>"
            strFriendSite = strFriendSite & XmlText("Site", "ShowFriendSiteList/t3", "<td width='60' align='center'>链接类型</td><td width='80' align='center'>网站名称</td><td width='100' align='center'>网站LOGO</td><td align='center'>网站简介</td><td width='60' align='center'>站长</td><td width='60' align='center'>操作</td>")
            strFriendSite = strFriendSite & "</tr>"

            Dim strT4, strT5
            strT4 = XmlText("Site", "ShowFriendSiteList/t4", "LOGO链接")
            strT5 = XmlText("Site", "ShowFriendSiteList/t5", "文字链接")

            Do While Not rsLink.EOF
                If EnableCountFriendSiteHits = True Then
                    LinkSiteUrl = InstallDir & "FriendSite/FriendSiteUrl.asp?ID=" & rsLink("ID")
                Else
                    LinkSiteUrl = rsLink("SiteUrl")
                End If
                strFriendSite = strFriendSite & "<tr class='channel_tdbg'>"
                strFriendSite = strFriendSite & "<td align='center'>"
                If rsLink("LinkType") = 1 Then
                    strFriendSite = strFriendSite & "<a class='LinkFriendSiteList' href='index.asp?LinkType=1'>" & strT4 & "</a>"
                Else
                    strFriendSite = strFriendSite & "<a class='LinkFriendSiteList' href='index.asp?LinkType=2'>" & strT5 & "</a>"
                End If
                strFriendSite = strFriendSite & "</td>"
                strFriendSite = strFriendSite & "<td><a class='LinkFriendSiteList' href='" & LinkSiteUrl & "' target='blank' title='" & rsLink("SiteUrl") & "'>" & rsLink("SiteName") & "</a></td>"
                strFriendSite = strFriendSite & "<td align='center'>"
                If rsLink("LinkType") = 1 Then
                    If rsLink("LogoUrl") <> "" And rsLink("LogoUrl") <> "http://" Then
                        If LCase(Right(rsLink("LogoUrl"), 3)) = "swf" Then
                            strFriendSite = strFriendSite & "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='88' height='31'><param name='movie' value='" & rsLink("LogoUrl") & "'><param name='quality' value='high'><embed src='" & rsLink("LogoUrl") & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='88' height='31'></embed></object>"
                        Else
                            strFriendSite = strFriendSite & "<a class='LinkFriendSiteList' href='" & LinkSiteUrl & "' target='_blank' title='" & rsLink("LogoUrl") & "'><img src='" & rsLink("LogoUrl") & "' width='88' height='31' border='0'></a>"
                        End If
                    Else
                        strFriendSite = strFriendSite & "<a class='LinkFriendSiteList' href='" & LinkSiteUrl & "' target='_blank'><img src='images/nologo.gif' width='88' height='31' border='0'></a>"
                    End If
                Else
                    strFriendSite = strFriendSite & "&nbsp;"
                End If
                strFriendSite = strFriendSite & "</td>"
                strFriendSite = strFriendSite & "<td>" & rsLink("SiteIntro") & "</td>"
                strFriendSite = strFriendSite & "<td align='center'><a class='LinkFriendSiteList' href='mailto:" & rsLink("SiteEmail") & "'>" & rsLink("SiteAdmin") & "</a></td>"
                strFriendSite = strFriendSite & "<td align='center'>"
                strFriendSite = strFriendSite & "<a class='LinkFriendSiteList' href='FriendSiteModify.asp?ID=" & rsLink("ID") & "'>修改</a>&nbsp;"
                strFriendSite = strFriendSite & "<a class='LinkFriendSiteList' href='FriendSiteDel.asp?ID=" & rsLink("ID") & "'>删除</a>"
                strFriendSite = strFriendSite & "</td>"
                strFriendSite = strFriendSite & "</tr>"
                i = i + 1
                If i >= MaxPerPage Then Exit Do
                rsLink.MoveNext
            Loop
            strFriendSite = strFriendSite & "</table>"
        End If
        rsLink.Close
        Set rsLink = Nothing
    End Select
    
    strHtml = Replace(strHtml, "{$FriendSiteList}", strFriendSite)
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("Site", "ShowFriendSiteList/PageChar", "个站点"), False))
    If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("Site", "ShowFriendSiteList/PageChar", "个站点"), False))
    
    regEx.Pattern = "\<(.[^\<\!]*)\>"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        strHtml = Replace(strHtml, Match.value, Replace(Match.value, "= ", "='' "))
    Next
End Sub

%>
