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

Dim ChannelID, PhotoID, Action, sql, rs, Hits
%>
<!--#include file="../conn.asp"-->
<!--#include file="Channel_Config.asp"-->
<!--#include file="../Include/PowerEasy.Common.Security.asp"-->
<%
Call OpenConn

PhotoID = PE_CLng(Trim(request("PhotoID")))
Action = Trim(request("Action"))
If Action = "Count" Then
    sql = "select sum(Hits) from PE_Photo where ChannelID=" & ChannelID
    Set rs = Conn.Execute(sql)
    If IsNull(rs(0)) Then
        Hits = 0
    Else
        Hits = rs(0)
    End If
    rs.Close
    Set rs = Nothing
Else
    If PhotoID = 0 Then
        Hits = 0
    Else
        sql = "select Hits,LastHitTime,DayHits,WeekHits,MonthHits from PE_Photo where Deleted=" & PE_False & " and Status=3 and PhotoID=" & PhotoID & " and ChannelID=" & ChannelID & ""
        Set rs = server.CreateObject("ADODB.recordset")
        rs.open sql, Conn, 1, 3
        If rs.bof And rs.EOF Then
            Hits = 0
        Else
            Hits = rs("Hits") + 1
            rs("Hits") = Hits
            If DateDiff("D", rs("LastHitTime"), Now()) <= 0 Then
                rs("DayHits") = rs("DayHits") + 1
            Else
                rs("DayHits") = 1
            End If
            If DateDiff("ww", rs("LastHitTime"), Now()) <= 0 Then
                rs("WeekHits") = rs("WeekHits") + 1
            Else
                rs("WeekHits") = 1
            End If
            If DateDiff("m", rs("LastHitTime"), Now()) <= 0 Then
                rs("MonthHits") = rs("MonthHits") + 1
            Else
                rs("MonthHits") = 1
            End If
            rs("LastHitTime") = Now()
            rs.Update
        End If
        rs.Close
        Set rs = Nothing
    End If
End If
Response.Write "document.write('" & Hits & "');"
Call CloseConn
%>
