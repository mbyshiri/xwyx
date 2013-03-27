<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************


Dim PrevSearchTime
PrevSearchTime = Trim(Session("Search_Time"))
If PrevSearchTime <> "" Then
    PrevSearchTime = CDate(PrevSearchTime)
    If DateDiff("s", PrevSearchTime, Now) < SearchInterval Then
        Response.Write "<br><br><br><p align='center'>为了避免恶意搜索而消耗大量系统资源，请您在 " & SearchInterval & " 秒后再刷新本页！</p>"
        Response.End
    End If
End If

ClassID = PE_CLng(Trim(Request("ClassID")))
SpecialID = PE_CLng(Trim(Request("SpecialID")))
SkinID = DefaultSkinID
PageTitle = "搜索结果"
strFileName = "Search.asp?ModuleName=" & ModuleName & "&Field=" & strField & "&Keyword=" & Keyword & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID
strPageTitle = SiteName & "----" & PageTitle
Call PE_Content.GetHTML_Search
Response.Write strHtml
Session("Search_Time") = Now
Set PE_Content = Nothing
Call CloseConn
%>
