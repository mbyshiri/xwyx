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
Dim ID, AnnounceNum
Dim sqlAnnounce, rsAnnounce, strAnnounce, AnnounceChannelID

ID = PE_CLng(Trim(Request("ID")))
AnnounceChannelID = PE_CLng(Trim(Request("ChannelID")))

PageTitle = "网站公告"
    
strHTML = GetTemplate(ChannelID, 4, 0)

Call ReplaceCommonLabel

strNavPath = strNavPath & strNavLink & "&nbsp;" & PageTitle

strHTML = Replace(strHTML, "{$PageTitle}", SiteTitle & " >> " & PageTitle)
strHTML = Replace(strHTML, "{$ShowPath}", strNavPath)

strHTML = Replace(strHTML, "{$MenuJS}", GetMenuJS("", False))
strHTML = Replace(strHTML, "{$Skin_CSS}", GetSkin_CSS(0))

sqlAnnounce = "select * from PE_Announce where IsSelected=" & PE_True & " and (ChannelID=-1 or ChannelID=" & AnnounceChannelID & ") and (OutTime=0 or OutTime>DateDiff(" & PE_DatePart_D & ",DateAndTime, " & PE_Now & ")) "
If ID > 0 Then
    sqlAnnounce = sqlAnnounce & " and (ShowType=0 or ShowType=1) and ID=" & ID
Else
    sqlAnnounce = sqlAnnounce & " and (ShowType=0 or ShowType=2)"
End If
sqlAnnounce = sqlAnnounce & " order by ID Desc"
Set rsAnnounce = Server.CreateObject("ADODB.Recordset")
rsAnnounce.Open sqlAnnounce, Conn, 1, 1
If rsAnnounce.BOF And rsAnnounce.EOF Then
    strAnnounce = strAnnounce & XmlText("Site", "ShowAnnounce/AnnounceErr", "<p>&nbsp;&nbsp;没有公告</p>")
Else
    AnnounceNum = rsAnnounce.RecordCount
    Dim i
    Do While Not rsAnnounce.EOF
        strAnnounce = strAnnounce & "<table width='100%'  border='0' cellspacing='0' cellpadding='0' style='word-break:break-all;Width:fixed'>"
        strAnnounce = strAnnounce & "<tr><td align='center' height='24' class='AnnounceTitle'>" & rsAnnounce("title") & "</td></tr>"
        strAnnounce = strAnnounce & "<tr><td align='left' valign='top'>" & rsAnnounce("Content") & "</td></tr>"
        strAnnounce = strAnnounce & "<tr><td align='right' valign='top'><p align=''>" & rsAnnounce("Author") & "&nbsp;&nbsp;<br>" & FormatDateTime(rsAnnounce("DateAndTime"), 1) & "</td></tr>"
        strAnnounce = strAnnounce & "</table>"
        rsAnnounce.MoveNext
        i = i + 1
        If i < AnnounceNum Then strAnnounce = strAnnounce & "<br>"
    Loop
End If
rsAnnounce.Close
Set rsAnnounce = Nothing

strHTML = Replace(strHTML, "{$AnnounceList}", strAnnounce)
If InStr(strHtml, "{$ShowPage}") > 0 Then strHTML = Replace(strHTML, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("Site", "ShowAnnounce/PageChar", "个公告"), False))
If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHTML = Replace(strHTML, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("Site", "ShowAnnounce/PageChar", "个公告"), False))
Response.Write strHTML
Call CloseConn
%>
