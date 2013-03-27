<!--#include file="Start.asp"-->
<!--#include file="Include/PowerEasy.Common.Front.asp"-->
<!--#include file="Include/PowerEasy.Cache.asp"-->
<!--#include file="Include/PowerEasy.Channel.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

ChannelID = 0
PageTitle = "网站公告列表"
strNavPath = strNavPath & strNavLink & "&nbsp;" & PageTitle
strFileName = "AnnounceList.asp"
strHTML = GetTemplate(ChannelID, 22, 0)
Call ReplaceCommonLabel

strHTML = Replace(strHTML, "{$PageTitle}", SiteTitle & " >> " & PageTitle)
strHTML = Replace(strHTML, "{$ShowPath}", strNavPath)

strHTML = Replace(strHTML, "{$MenuJS}", GetMenuJS("", False))
strHTML = Replace(strHTML, "{$Skin_CSS}", GetSkin_CSS(0))

regEx.Pattern = "【AnnounceList\((.*?)\)】([\s\S]*?)【\/AnnounceList】"
Set Matches = regEx.Execute(strHTML)
For Each Match In Matches
	strHTML = PE_Replace(strHTML, Match.value, GetCustomFromLabel(Match.SubMatches(0), Match.SubMatches(1)))
Next

If InStr(strHtml, "{$ShowPage}") > 0 Then strHTML = Replace(strHTML, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("Site", "ShowAnnounce/PageChar", "个公告"), False))
If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHTML = Replace(strHTML, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("Site", "ShowAnnounce/PageChar", "个公告"), False))

Response.Write strHTML
Call CloseConn


Function GetCustomFromLabel(strTemp, strList)
    Dim arrTemp
    Dim OrderType, OpenType
    
    If strTemp = "" Then
        GetCustomFromLabel = ""
        Exit Function
    End If
    
    Dim sqlCustom, rsCustom, iCount, strCustomList, strThisClass, strLink
    iCount = 0
    sqlCustom = ""
    strThisClass = ""
    strCustomList = ""
    sqlCustom = "select * from PE_Announce"

    arrTemp = Split(strTemp, ",")
    If PE_CLng(Trim(arrTemp(0))) <> 0 Then
        sqlCustom = sqlCustom & " where DateDiff(" & PE_DatePart_D & ",DateAndTime, " & PE_Now & ") <" & PE_CLng(arrTemp(0))
    End If
    If Trim(arrTemp(1)) = 1 Then
        OpenType = " target=_blank"
    Else
        OpenType = " target=_self"
    End If
    If Trim(arrTemp(2)) = 1 Then
        sqlCustom = sqlCustom & " order by DateAndTime Desc"
    Else
        sqlCustom = sqlCustom & " order by DateAndTime asc"
    End If
    Set rsCustom = Server.CreateObject("ADODB.Recordset")
    rsCustom.Open sqlCustom, Conn, 1, 1
    If rsCustom.BOF And rsCustom.EOF Then

        strCustomList = "网站暂时没有任何公告！"
        rsCustom.Close
        Set rsCustom = Nothing
        GetCustomFromLabel = strCustomList
        Exit Function
    End If
    totalPut = rsCustom.RecordCount
    
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
            rsCustom.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If
    Do While Not rsCustom.EOF
        strLink = "<a href=announce.asp?ChannelID="&  rsCustom("ChannelID") &"&ID=" & rsCustom("ID") & OpenType & ">"
        strTemp = PE_Replace(strList, "{$AnnounceTitle}", rsCustom("Title"))
        strTemp = PE_Replace(strTemp, "{$AnnounceContent}", strLink & rsCustom("Content") & "</a>")
        strTemp = PE_Replace(strTemp, "{$AnnounceAuthor}", rsCustom("Author"))
        strTemp = PE_Replace(strTemp, "{$AnnounceDateAndTime}", FormatDateTime(rsCustom("DateAndTime"), 1))
        
        strCustomList = strCustomList & strTemp
        rsCustom.MoveNext
        iCount = iCount + 1
        If iCount >= MaxPerPage Then Exit Do
    Loop
    rsCustom.Close
    Set rsCustom = Nothing
    
    GetCustomFromLabel = strCustomList
End Function

%>
