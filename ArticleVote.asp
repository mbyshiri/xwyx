<!--#include file="start.asp" -->
<style>
*{text-decoration:none;color:#636363;font-size:12px;border:0px;padding-top:2px;margin:0px;font-weight:normal}
#z{text-decoration:none;color:#636363;width:100px;}
#y{ float:left;text-decoration:none;color:#ff0000; text-align:center; padding-left:10px}
#s{ float:right;width:42px; border:1px #66CCFF solid; background:#C8EDFF; text-align:center}
#s1{ text-decoration:none;color:#ffffff;float:right;width:42px; border:1px #FD3013 solid; background:#C51D02; text-align:center}

</style>
<%
Dim ArticleID, rs, q, cookie, sql
ArticleID = PE_Clng(Trim(request("ArticleID")))
ComeUrl = ReplaceUrlBadChar(request.ServerVariables("HTTP_REFERER"))
cookie=ReplaceBadChar(request.cookies("ArticleVote")(""&ArticleID&""))
Action = ReplaceBadChar(Trim(request("action")))

Const add = "article/shownew.asp" '在这里填入“查看”二字的URL地址，如：http://guest.pasun.cn/article/shownew.asp则填入：article/shownew.asp即可

sql = "select MY_ArticleVote from PE_Article where Deleted=" & PE_False & " and Status=3 and ArticleID=" & ArticleID & ""
If Action = "up" Then
    If cookie = "" Then
        Set rs = server.CreateObject("ADODB.recordset")
        rs.open sql, conn, 1, 3
        If Not (rs.bof And rs.EOF) Then
            rs("MY_ArticleVote") = rs("MY_ArticleVote") + 1
            rs.Update
            response.cookies("ArticleVote")(""&ArticleID&"") = "1"
            Response.Cookies("ArticleVote").Expires = Date + 3650
        End If
        rs.Close
        Set rs = Nothing
    End If
    Response.Write"<script>window.location="""&ComeUrl&""";</script>"
Else
    Set rs = conn.execute(sql)
    Response.Write "<div id='z'>"
    Response.Write "<div id='y'>"
    If rs("MY_ArticleVote")="" or isnull(rs("MY_ArticleVote")) then
        conn.execute("update PE_Article set MY_ArticleVote=0 where ArticleID=" & ArticleID )
        Response.Write rs("MY_ArticleVote")
        Response.Write "<font color='#636363'>&nbsp;票</font>"
    Else
        Response.Write rs("MY_ArticleVote")
        Response.Write "<font color='#636363'>&nbsp;票</font>"
    End If
    Response.Write "</div>"
    If cookie = "" Then
        Response.Write "<div id='s'>"
        Response.Write"<a href="&InstallDir&"ArticleVote.asp?articleid="&ArticleID&"&action=up"">投一票</a>"
    Else
        Response.Write "<div id='s1'>"
        Response.Write "感谢你"
        Response.Write "</div>"
        Response.Write "</div>"

    End If
    Response.Write "</div>"
    rs.Close
    Set rs = Nothing
End If
%>




