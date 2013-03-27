<%@language=vbscript codepage=936 %>
<%
Option Explicit
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim ChannelID, SoftID, sql, rs, Hits
%>
<!--#include file="../conn.asp"-->
<!--#include file="Channel_Config.asp"-->
<!--#include file="../Include/PowerEasy.Common.Security.asp"-->
<%
Call OpenConn

SoftID = PE_CLng(Trim(request("SoftID")))
If SoftID = 0 Then
    Hits = 0
Else
    sql = "select BrowseTimes from PE_Soft where Deleted=" & PE_False & " and Status=3 and SoftID=" & SoftID & " and ChannelID=" & ChannelID & ""
    Set rs = server.CreateObject("ADODB.recordset")
    rs.open sql, Conn, 1, 3
    If rs.bof And rs.EOF Then
        Hits = 0
    Else
        Hits = rs(0) + 1
        rs(0) = Hits
        rs.Update
    End If
    rs.Close
    Set rs = Nothing
End If
Response.Write "document.write('" & Hits & "');"
Call CloseConn
%>
