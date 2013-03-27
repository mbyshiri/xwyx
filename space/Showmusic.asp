<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

If Action = "savepl" Then
    Call ShowDiary("music")
Else
    If UserID > 0 Or BlogID > 0 Then
        If UserID > 0 Then
            Set UqRs = Conn.Execute("select A.UserID,C.UserName from PE_SpaceMusic A inner join PE_User C on A.UserID=C.UserID Where A.ID=" & UserID)
        Else
            Set UqRs = Conn.Execute("select A.UserID,C.UserName from PE_Space A inner join PE_User C on A.UserID=C.UserID Where A.ID=" & BlogID)
        End If
        If Not (UqRs.BOF And UqRs.EOF) Then
            BlogDir = UqRs(1) & UqRs(0)
        End If
        Set UqRs = Nothing

        If bootnode.length = 0 Then
            If Action <> "xml" Then strtmp = strtmp & "<?xml-stylesheet type=""text/xsl"" href=""" & BlogDir & "/Showmusic.xsl"" version=""1.0""?>"
            MaxPerPage = 10
        Else
            On Error Resume Next
            If Action <> "xml" Then strtmp = strtmp & "<?xml-stylesheet type=""text/xsl"" href=""" & BlogDir & "/" & bootnode(0).selectSingleNode("music/template").text & """ version=""1.0""?>"
            MaxPerPage = PE_CLng(bootnode(0).selectSingleNode("music/MaxPerPage").text)
            If MaxPerPage = 0 Then MaxPerPage = 10
        End If
        Call ShowDiary("music")
    Else
        strtmp = ""
    End If
    Set xmlconfig = Nothing
End If
strtmp = strtmp & XMLDOM.documentElement.xml

Set Node = Nothing
Set SubNode = Nothing
Set XMLDOM = Nothing

Response.Write strtmp

Call CloseConn
%>
