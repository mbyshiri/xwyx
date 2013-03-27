<!--#include file="../Start.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private"
Response.CacheControl = "no-cache"
Response.Charset = "GB2312"
Response.ContentType = "text/html"

If CheckUserLogined() = False Then
    Call CloseConn
    Response.End
End If

Call Receive
Call CloseConn

Sub Receive()
    Dim ArticleID
    Dim sqlReceive, rsReceive

    ArticleID = PE_CLng(Trim(Request("ArticleID")))
    If ArticleID = 0 Then
        FoundErr = True
        Exit Sub
    End If

    sqlReceive = "select * from PE_Article where ArticleID=" & ArticleID
    Set rsReceive = Server.CreateObject("ADODB.Recordset")
    rsReceive.Open sqlReceive, Conn, 1, 3
    If rsReceive.BOF And rsReceive.EOF Then
        FoundErr = True
    Else
        If FoundInArr(rsReceive("ReceiveUser"), UserName, ",") = False Then
            FoundErr = True
        End If
        If FoundInArr(rsReceive("Received"), UserName, "|") = True Then
            FoundErr = True
        End If
    End If
    If FoundErr = True Then
        rsReceive.Close
        Set rsReceive = Nothing
        Exit Sub
    End If
    If rsReceive("Received") = "" Or IsNull(rsReceive("Received")) Then
        rsReceive("Received") = UserName
    Else
        rsReceive("Received") = rsReceive("Received") & "|" & UserName
    End If
    rsReceive.Update
    rsReceive.Close
    Set rsReceive = Nothing

    Dim sqlUser, rsUser, tmpUnsignedItems, tmpArticleID
    Set rsUser = Server.CreateObject("adodb.recordset")
    sqlUser = "select UserID,UserName,UnsignedItems from PE_User where UserName='" & UserName & "'"
    rsUser.Open sqlUser, Conn, 1, 3
    If Not rsUser.EOF Then
        If FoundInArr(rsUser("UnsignedItems"), CStr(ArticleID), ",") = True Then
            tmpUnsignedItems = "," & rsUser("UnsignedItems") & ","
            tmpArticleID = "," & ArticleID & ","
            tmpUnsignedItems = Replace(tmpUnsignedItems, tmpArticleID, ",")
            If tmpUnsignedItems = "," Then
                rsUser("UnsignedItems") = ""
            Else
                rsUser("UnsignedItems") = Mid(tmpUnsignedItems, 2, Len(tmpUnsignedItems) - 2)
            End If
            rsUser.Update
        End If
    End If
    rsUser.Close
    Set rsUser = Nothing
    Response.Write "OK"
End Sub

%>
