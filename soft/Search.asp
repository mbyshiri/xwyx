<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************


Dim PrevSearchTime
PrevSearchTime = Trim(Session("Search_Time"))
If PrevSearchTime <> "" Then
    PrevSearchTime = CDate(PrevSearchTime)
    If DateDiff("s", PrevSearchTime, Now) < SearchInterval Then
        Response.Write "<br><br><br><p align='center'>Ϊ�˱���������������Ĵ���ϵͳ��Դ�������� " & SearchInterval & " �����ˢ�±�ҳ��</p>"
        Response.End
    End If
End If

ClassID = PE_CLng(Trim(Request("ClassID")))
SpecialID = PE_CLng(Trim(Request("SpecialID")))
SkinID = DefaultSkinID
PageTitle = "�������"
strFileName = "Search.asp?ModuleName=" & ModuleName & "&Field=" & strField & "&Keyword=" & Keyword & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID
strPageTitle = SiteName & "----" & PageTitle
Call PE_Content.GetHTML_Search
Response.Write strHtml
Session("Search_Time") = Now
Set PE_Content = Nothing
Call CloseConn
%>
