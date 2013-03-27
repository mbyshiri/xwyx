<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Sub CheckComeUrl()
    Dim ComeUrl, TrueSiteUrl, cUrl
    ComeUrl = Trim(Request.ServerVariables("HTTP_REFERER"))
    TrueSiteUrl = Trim(Request.ServerVariables("HTTP_HOST"))
    If ComeUrl = "" Then
        Response.Write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许直接输入地址访问本系统的后台管理页面。</font></p>"
        Call WriteEntry(1, "", "直接地址输入访问后台")
        Response.End
    Else
        cUrl = Trim("http://" & TrueSiteUrl) & ScriptName
        If LCase(Left(ComeUrl, InStrRev(ComeUrl, "/"))) <> LCase(Left(cUrl, InStrRev(cUrl, "/"))) Then
            Response.Write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许从外部链接地址访问本系统的后台管理页面。</font></p>"
            Call WriteEntry(1, "", "外部链接访问后台")
            Response.End
        End If
    End If
End Sub

Sub ShowPageTitle(strTitle, UrlID)
    Response.Write "  <tr class='topbg'> " & vbCrLf
    Response.Write "    <td height='22' colspan='10'><table width='100%'><tr class='topbg'><td align='center'><b>" & strTitle & "</b></td><td width='60' align='right'><a href='http://go.powereasy.net/go.aspx?UrlID=" & UrlID & "' target='_blank'><img src='images/help.gif' border='0'></a></td></tr></table></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
End Sub

'=================================================
'过程名：CheckSecretCode
'作  用：效验安全码
'=================================================
Function CheckSecretCode(ByVal iCode)
    Dim j, secritycode, rNum, lcode
    secritycode = ""
    If iCode = "start" Then
        Randomize Timer
        lcode = "0123456789abcdefghijklmnopqrstuvwxyz"
        For j = 0 To 10
            rNum = CInt(35 * Rnd)
            secritycode = secritycode & Mid(lcode, rNum + 1, 1)
        Next
        Session("AdminSecretCode") = secritycode
        CheckSecretCode = secritycode
    Else
        If iCode = "" Or iCode <> Session("AdminSecretCode") Then
            CheckSecretCode = False
        Else
            CheckSecretCode = True
        End If
        Session("AdminSecretCode") = ""
    End If
End Function

'**************************************************
'方法名：SendMessage
'作  用：添加一条短消息
'参  数：InceptUser ----用户名称
'        Title ---- 短消息标题
'        Content ---- 短消息内容
'        SendUser ---- 发布人
'**************************************************
Sub SendMessage(InceptUser, Title, Content, SendUser)
    Dim rsMessage, sqlMessage, arrIncept, i
    arrIncept = Split(InceptUser, ",")
    Set rsMessage = Server.CreateObject("adodb.recordset")
    sqlMessage = "select top 1 * from PE_Message"
    rsMessage.Open sqlMessage, Conn, 1, 3
    For i = 0 To UBound(arrIncept)
        rsMessage.addnew
        rsMessage("Incept") = arrIncept(i)
        rsMessage("Sender") = SendUser
        rsMessage("Title") = Title
        rsMessage("Content") = Content
        rsMessage("SendTime") = Now()
        rsMessage("Flag") = 0
        rsMessage("IsSend") = 1
        rsMessage.Update
        Conn.Execute ("update PE_User set UnreadMsg=UnreadMsg+1 where UserName='" & arrIncept(i) & "'")
    Next
    rsMessage.Close
    Set rsMessage = Nothing
End Sub


'**************************************************
'函数名：Replace_CaseInsensitive
'作  用：替换字符，大小写不敏感
'参  数：expression ---- 字符串表达式 包含要替代的子字符串
'        find ---- 被搜索的子字符串
'        replacewith ---- 用于替换的子字符串
'返回值：处理后的字符串
'**************************************************
Function Replace_CaseInsensitive(expression, find, replacewith)
    regEx.Pattern = find
    Replace_CaseInsensitive = regEx.Replace(expression, replacewith)
End Function
%>
