<!--#include file="CommonCode.asp"-->
<!--#include file="../Include/PowerEasy.SendMail.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim MailType

Select Case MailObject
Case 0
    FoundErr = True
    ErrMsg = ErrMsg & "对不起，服务器没有选定任何邮件发送组件！所以不能使用本功能。"
Case 1
    If Not IsObjInstalled("JMail.Message") Then
        FoundErr = True
        ErrMsg = ErrMsg & "JMail邮件发送组件没有安装！所以不能使用本功能。"
    End If
Case 2
    If Not IsObjInstalled("CDONTS.NewMail") Then
        FoundErr = True
        ErrMsg = ErrMsg & "CDONTS邮件发送组件没有安装！所以不能使用本功能。"
    End If
Case 3
    If Not IsObjInstalled("Persits.MailSender") Then
        FoundErr = True
        ErrMsg = ErrMsg & "ASPEMAIL邮件发送组件没有安装！所以不能使用本功能。"
    End If
Case 4
    If Not IsObjInstalled("easymail.mailsend") Then
        FoundErr = True
        ErrMsg = ErrMsg & "WebEasyMail邮件发送组件没有安装！所以不能使用本功能。"
    End If
Case Else
    FoundErr = True
    ErrMsg = ErrMsg & "对不起，服务器邮件发送组件不对！所以不能使用本功能。"
End Select

ArticleID = PE_CLng(Request("ArticleID"))
If ArticleID = 0 Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>请指定要发送给好友的文章ID！</li>"
End If
If UserLogined = False Then
    FoundErr = True
    ErrMsg = ErrMsg & "<br>&nbsp;&nbsp;&nbsp;&nbsp;你还没注册？或者没有登录？只有本站的注册用户才能使用“告诉好友”功能！<br><br>&nbsp;&nbsp;&nbsp;&nbsp;如果你还没注册，请赶紧<a href='../Reg/User_Reg.asp'><font color=red>点此注册</font></a>吧！<br><br>&nbsp;&nbsp;&nbsp;&nbsp;如果你已经注册但还没登录，请赶紧<a href='../User/User_Login.asp'><font color=red>点此登录</font></a>吧！<br><br>"
End If

If FoundErr <> True Then
    If Action = "MailToFriend" Then
        Call MailToFriend
    Else
        Call SendMailMain
    End If
Else
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Set PE_Content = Nothing
Call CloseConn

Sub SendMailMain()
    Dim rs, sql, Title, Author, UpdateTime
    sql = "Select Title,UpdateTime,Author from PE_Article where ArticleID=" & ArticleID & ""
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 1
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "找不到文章"
        FoundErr = True
    Else
        Title = rs("Title")
        Author = rs("Author")
        UpdateTime = rs("UpdateTime")
    End If
    rs.Close
    Set rs = Nothing
    strHtml = GetTemplate(ChannelID, 20, 0)
    
    Call ReplaceCommonLabel
    
    strHtml = PE_Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))
    strHtml = PE_Replace(strHtml, "{$Title}", Title)
    strHtml = PE_Replace(strHtml, "{$ComeUrl}", ComeUrl)
    strHtml = PE_Replace(strHtml, "{$ArticleID}", ArticleID)
    strHtml = PE_Replace(strHtml, "{$Author}", Author)
    strHtml = PE_Replace(strHtml, "{$UpdateTime}", UpdateTime)
    strHtml = Replace(strHtml, "value= ", "value='' ")
    strHtml = Replace(strHtml, "Value= ", "value='' ")
    Response.Write strHtml
End Sub

Sub MailToFriend()
    Dim MailtoName, MailtoAddress, Subject, MailBody

    MailtoName = Trim(Request.Form("MailToName"))
    MailtoAddress = Trim(Request.Form("MailToAddress"))
    If MailtoName = "" Then
        ErrMsg = ErrMsg & "<li>收信人姓名为空！</li>"
        FoundErr = True
    End If
    If IsValidEmail(MailtoAddress) = False Then
        ErrMsg = ErrMsg & "<li>收信人的Email地址有错误！</li>"
        FoundErr = True
    End If
    If FoundErr Then Exit Sub

    Dim rs, sql, strContent
    sql = "Select A.ChannelID,A.Title,A.Content,A.UpdateTime,A.Author,A.InfoPoint,C.ClassPurview from PE_Article A left join PE_Class C on A.ClassID=C.ClassID where A.ArticleID=" & ArticleID & ""
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 1
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "找不到文章"
    Else
        Subject = Replace(Replace("您的朋友{$UserName}从{$SiteName}给您发来的文章资料料", "{$UserName}", UserName), "{$SiteName}", SiteName)
        If rs("ClassPurview") > 0 Or rs("InfoPoint") > 0 Then
            strContent = "<a href='" & Trim(Request.ServerVariables("HTTP_HOST")) & ChannelUrl_ASPFile & "/ShowArticle.asp?ArticleID=" & ArticleID & "'>点击查看此页面的内容</a>"
        Else
            strContent = Replace(Replace(rs("Content") & "", "[InstallDir_ChannelDir]", Trim(Request.ServerVariables("HTTP_HOST")) & ChannelUrl & "/"), "{$UploadDir}", UploadDir)
        End If
        MailBody = Replace(Replace(Replace(Replace(Replace(Replace("<style>A:visited {  TEXT-DECORATION: none   }A:active  { TEXT-DECORATION: none   }A:hover   { TEXT-DECORATION: underline overline }A:link    { text-decoration: none;}A:visited { text-decoration: none;}A:active  { TEXT-DECORATION: none;}A:hover   { TEXT-DECORATION: underline overline}BODY   {    FONT-FAMILY: 宋体; FONT-SIZE: 9pt;}TD     {    FONT-FAMILY: 宋体; FONT-SIZE: 9pt   }</style><TABLE border=0 width='95%' align=center><TBODY><TR><TD valign=middle align=top>--&nbsp;&nbsp;作者：{$Author}<br>--&nbsp;&nbsp;发布时间：{$Time}<br><br>--&nbsp;&nbsp;{$title}<br>--&nbsp;&nbsp;{$Content}<br></TD></TR></TBODY></TABLE><center><a href='{$SiteUrl}'>{$SiteName}</a>", "{$Author}", rs("Author")), "{$Time}", rs("UpdateTime")), "{$title}", rs("title")), "{$Content}", strContent), "{$SiteUrl}", SiteUrl), "{$SiteName}", SiteName)
    End If
    rs.Close
    Set rs = Nothing

    Dim PE_Mail
    Set PE_Mail = New SendMail
    If ErrMsg <> "" Then
        FoundErr = True
        Set PE_Mail = Nothing
        Exit Sub
    End If
    ErrMsg = PE_Mail.Send(MailtoAddress, MailtoName, Subject, MailBody, UserName, WebmasterEmail, 3)
    Set PE_Mail = Nothing

    If ErrMsg = "" Then
        Call WriteSuccessMsg("已经成功将此文章发送给你的好友！", ComeUrl)
    Else
        FoundErr = True
    End If
End Sub
%>
