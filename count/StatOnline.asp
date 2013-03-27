<%@language=vbscript codepage=936 %>
<%
Option Explicit
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!--#include file="conn_counter.asp"-->
<!--#include file="../Include/PowerEasy.Common.Security.asp"-->
<%
Dim RegCount_Fill, OnlineTime
Response.Expires = 0
Call OpenConn_Counter

If IsEmpty(Application("OnlineTime")) Then
    Dim rs
    Set rs = conn_counter.Execute("select * from PE_StatInfoList")
    If Not rs.bof And Not rs.EOF Then
        OnlineTime = rs("OnlineTime")
        Application("OnlineTime") = OnlineTime
    End If
    Set rs = Nothing
Else
    OnlineTime = Application("OnlineTime")
End If

Dim PE_IP, PE_Agent, PE_Page, OnNowTime
PE_IP = ReplaceBadChar(Request.ServerVariables("Remote_Addr"))
PE_Agent = ReplaceBadChar(Request.ServerVariables("HTTP_USER_AGENT"))
PE_Page = ReplaceUrlBadChar(Request.ServerVariables("HTTP_REFERER"))

OnNowTime = DateAdd("s", 0 - OnlineTime, Now())
'response.write "now="&now()&"OnNowTime="&OnNowTime&"OnlineTime="&OnlineTime
Dim rsOnline, rsOd
If CountDatabaseType = "SQL" Then
    Set rsOnline = conn_counter.Execute("select * from PE_StatOnline where LastTime>'" & OnNowTime & "' and UserIP='" & PE_IP & "'")
Else
    Set rsOnline = conn_counter.Execute("select * from PE_StatOnline where LastTime>#" & OnNowTime & "# and UserIP='" & PE_IP & "'")
End If
If rsOnline.EOF Then
    If CountDatabaseType = "SQL" Then
        Set rsOd = conn_counter.Execute("select top 1 id from PE_StatOnline where LastTime<'" & OnNowTime & "'  order by LastTime")
    Else
        Set rsOd = conn_counter.Execute("select top 1 id from PE_StatOnline where LastTime<#" & OnNowTime & "#  order by LastTime")
    End If
    If rsOd.EOF Then
        conn_counter.Execute "insert into PE_StatOnline (UserIP,UserAgent,UserPage,OnTime,LastTime) VALUES('" & PE_IP & "','" & PE_Agent & "','" & PE_Page & "'," & PECount_Now & "," & PECount_Now & ")"
    Else
        conn_counter.Execute "update PE_StatOnline set UserIP='" & PE_IP & "',UserAgent='" & PE_Agent & "',UserPage='" & PE_Page & "',Ontime=" & PECount_Now & ",LastTime=" & PECount_Now & " where id=" & rsOd("id")
    End If
    Set rsOd = Nothing
Else
    If CountDatabaseType = "SQL" Then
        conn_counter.Execute ("update PE_StatOnline set LastTime=" & PECount_Now & ",UserPage='" & PE_Page & "' where LastTime>'" & OnNowTime & "' and UserIP='" & PE_IP & "'")
    Else
        conn_counter.Execute ("update PE_StatOnline set LastTime=" & PECount_Now & ",UserPage='" & PE_Page & "' where LastTime>#" & OnNowTime & "# and UserIP='" & PE_IP & "'")
    End If
End If
Set rsOnline = Nothing
Call CloseConn_Counter
Server.Transfer ("Image/powereasyimg.gif")

%>
