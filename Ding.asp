<!--#include file="start.asp" -->
<style>
*{text-decoration:none;color:#000;font-size:12px;border:0px;padding:0px;margin:0px;font-weight:normal}
div{border:#c9c9c9 solid 1px;text-align:center;width:50px}
h1{background:#FDFFAC;width:50px;height:40px;border-bottom:#c9c9c9 solid 1px;font-weight:bold;line-height:40px;font-family:'����';font-size:20px;color:#950}
p{height:20px;line-height:20px}
</style>
<%
Dim ArticleID, rs, q, cookie, sql
ArticleID = PE_Clng(Trim(request("ArticleID")))
Cookie = ReplaceBadChar(Request.Cookies("rhongsheng")("Article_"&ArticleID&""))

Const Add = "article/shownew.asp"  '���������롰�鿴�����ֵ�URL��ַ

sql = "select MY_upart from PE_Article where Deleted=" & PE_False & " and Status=3 and ArticleID=" & ArticleID
If Action = "up" Then 
	If Cookie = "" Then 
		Set rs = server.CreateObject("ADODB.recordset")
		rs.open sql, conn, 1, 3
		If Not (rs.bof and rs.EOF) Then 
			rs("MY_upart") = rs("MY_upart") + 1
			rs.update
			Response.Cookies("rhongsheng")("Article_"&ArticleID&"") = 1
			Response.Cookies("rhongsheng").Expires = Date + 3650
		End If
		rs.Close
		Set rs = nothing
	End If
	Response.Redirect Request.Servervariables("http_referer")
Else 
	Set rs = Conn.Execute(sql)
	Response.Write"<div>"
	Response.Write"<h1>"
        If rs("MY_upart")="" or isnull(rs("MY_upart")) then
        Response.Write 0
            conn.execute("update PE_Article set MY_upart=0 where ArticleID=" & ArticleID )
            
	Else
		Response.Write rs("MY_upart")
	End If
	Response.Write"</h1><p>"
	If Cookie = "" Then
		Response.Write"<a href=""ding.asp?articleid="&ArticleID&"&action=up"">��һ��</a>"
	Else
		Response.Write"<a href="""&InstallDir&add&""" target=""_top"">�鿴</a>"
	End If
	Response.Write"</p></div>"
	rs.Close
	Set rs = Nothing
End If
CloseConn
%>

