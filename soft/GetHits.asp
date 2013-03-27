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

Dim ChannelID, SoftID, Action, HitsType, sql, rs, Hits
%>
<!--#include file="../conn.asp"-->
<!--#include file="Channel_Config.asp"-->
<!--#include file="../Include/PowerEasy.Common.Security.asp"-->
<%
Call OpenConn
Action = Trim(request("Action"))
SoftID = PE_CLng(Trim(request("SoftID")))
HitsType = PE_CLng(Trim(request("HitsType")))
If Action = "Count" Then
    sql = "select sum(Hits) from PE_Soft where ChannelID=" & ChannelID
    Set rs = Conn.Execute(sql)
    If IsNull(rs(0)) Then
        Hits = 0
    Else
        Hits = rs(0)
    End If
    rs.Close
    Set rs = Nothing
ElseIf Action = "SoftDown" Then
    Hits = ""
    Dim rsSoft
    sql = "select * from PE_Soft where Deleted=" & PE_False & " and Status=3 and SoftID=" & SoftID & " and ChannelID=" & ChannelID & ""
    Set rsSoft = server.CreateObject("ADODB.Recordset")
    rsSoft.open sql, Conn, 1, 3
    If Not (rsSoft.bof And rsSoft.EOF) Then
        rsSoft("Hits") = rsSoft("Hits") + 1
        If DateDiff("D", rsSoft("LastHitTime"), Now()) <= 0 Then
            rsSoft("DayHits") = rsSoft("DayHits") + 1
        Else
            rsSoft("DayHits") = 1
        End If
        If DateDiff("ww", rsSoft("LastHitTime"), Now()) <= 0 Then
            rsSoft("WeekHits") = rsSoft("WeekHits") + 1
        Else
            rsSoft("WeekHits") = 1
        End If
        If DateDiff("m", rsSoft("LastHitTime"), Now()) <= 0 Then
            rsSoft("MonthHits") = rsSoft("MonthHits") + 1
        Else
            rsSoft("MonthHits") = 1
        End If
        rsSoft("LastHitTime") = Now()
        rsSoft.Update
    End If
    rsSoft.Close
    Set rsSoft = Nothing
Else
    If SoftID = 0 Then
        Hits = 0
    Else
        Select Case HitsType
        Case 0
            sql = "select Hits from PE_Soft where SoftID=" & SoftID
        Case 1
            sql = "select DayHits from PE_Soft where SoftID=" & SoftID
        Case 2
            sql = "select WeekHits from PE_Soft where SoftID=" & SoftID
        Case 3
            sql = "select MonthHits from PE_Soft where SoftID=" & SoftID
        Case Else
            sql = "select Hits from PE_Soft where SoftID=" & SoftID
        End Select
        Set rs = server.CreateObject("ADODB.recordset")
        rs.open sql, Conn, 1, 1
        If rs.bof And rs.EOF Then
            Hits = 0
        Else
            Hits = rs(0)
        End If
        rs.Close
        Set rs = Nothing
    End If
End If
Response.Write "document.write('" & Hits & "');"
Call CloseConn
%>
