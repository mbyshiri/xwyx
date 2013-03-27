<%@language="vbscript" codepage="936" %>
<%
Option Explicit
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Response.Buffer = True
%>
<!--#include file="../conn.asp"-->
<%
Call OpenConn

Dim ID
ID = Trim(Request("ID"))
If ID <> "" And IsNumeric(ID) Then
    ID = CLng(ID)
    Conn.Execute ("update PE_FriendSite set Hits=Hits+1 where ID=" & ID)
    Dim rsFriendSite, FriendSiteUrl
    Set rsFriendSite = Conn.Execute("select SiteUrl from PE_FriendSite where ID=" & ID)
    If Not (rsFriendSite.bof And rsFriendSite.EOF) Then
        FriendSiteUrl = rsFriendSite("SiteUrl")
    End If
    rsFriendSite.Close
    Set rsFriendSite = Nothing
    If FriendSiteUrl <> "" Then Response.Redirect (FriendSiteUrl)
End If
Call CloseConn
%>
