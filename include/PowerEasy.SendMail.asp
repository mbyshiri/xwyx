<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Class SendMail

Private MailtoAddress, MailtoName, Subject, MailBody, FromName, MailFrom, Priority

Public Function Send(sMailtoAddress, sMailtoName, sSubject, sMailBody, sFromName, sMailFrom, sPriority)
    Send = ""
    ErrMsg = ""
    MailtoAddress = sMailtoAddress
    MailtoName = sMailtoName
    Subject = sSubject
    MailBody = sMailBody
    FromName = sFromName
    MailFrom = sMailFrom
    Priority = sPriority
    '选择邮件发送组件
    Select Case MailObject
    Case 0
        FoundErr = True
        ErrMsg = ErrMsg & "没有选定任何邮件发送组件！所以不能发送注册确认信。"
    Case 1
        ErrMsg = JSendMail()
    Case 2
        ErrMsg = CdontsMail()
    Case 3
        ErrMsg = Aspemail()
    Case 4
        ErrMsg = WebEasyMail()
    Case Else
        FoundErr = True
        ErrMsg = ErrMsg & "服务器邮件发送组件不对！所以不能发送注册确认信。"
    End Select
    If ErrMsg <> "" Then
        Send = ErrMsg
    End If
End Function

Private Function JSendMail()
    On Error Resume Next
    Dim JMail
    Set JMail = Server.CreateObject("JMail.Message")
    JMail.Charset = "gb2312"        '邮件编码
    JMail.silent = True
    JMail.ContentType = "text/html"     '邮件正文格式
    'JMail.ServerAddress=MailServer     '用来发送邮件的SMTP服务器
    '如果服务器需要SMTP身份验证则还需指定以下参数
    JMail.MailServerUserName = MailServerUserName    '登录用户名
    JMail.MailServerPassWord = MailServerPassWord        '登录密码
    JMail.MailDomain = MailDomain       '域名（如果用“name@domain.com”这样的用户名登录时，请指明domain.com
    JMail.AddRecipient MailtoAddress, MailtoName    '收信人
    JMail.Subject = Subject       '主题
    'JMail.HtmlBody = MailBody     '邮件正文（HTML格式）
    JMail.Body = MailBody        '邮件正文（纯文本格式）
    JMail.FromName = FromName       '发信人姓名
    JMail.From = MailFrom         '发信人Email
    JMail.Priority = Priority            '邮件等级，1为加急，3为普通，5为低级
    JMail.Send (MailServer)
    JSendMail = JMail.ErrorMessage
    JMail.Close
    Set JMail = Nothing
    JSendMail = ""
End Function

Private Function CdontsMail()
    On Error Resume Next
    Dim CDOMail
    Set CDOMail = Server.CreateObject("CDONTS.NewMail")
    CDOMail.From = MailFrom
    CDOMail.To = MailtoAddress
    CDOMail.Subject = Subject
    'CDOMail.Cc = sCc   '抄送给其他人，可以指定多个，用逗号隔开
    'CDOMail.BCc = sBCc '暗抄给其他人，可以指定多个，用逗号隔开
    CDOMail.BodyFormat = 0   '指定邮件为MINE格式
    CDOMail.MailFormat = 0   '指定邮件为HTML格式
    CDOMail.Importance = 0   '指定邮件级别 0: 普通 1：机密 2:绝密
    CDOMail.Body = MailBody
    CDOMail.Send
    Set CDOMail = Nothing
    If Err Then
        CdontsMail = "<li>邮件发送失败!</li>"
        Err.Clear
        Exit Function
    End If
    CdontsMail = ""
End Function

Private Function Aspemail()
    On Error Resume Next
    Dim Mailer
    Set Mailer = Server.CreateObject("Persits.MailSender")
    Mailer.Charset = "gb2312"
    Mailer.IsHTML = True
    Mailer.UserName = MailServerUserName    '登录用户名
    Mailer.Password = MailServerPassWord    '登录密码
    Mailer.Priority = 1
    Mailer.Host = MailServer
    Mailer.Port = 25
    Mailer.From = MailFrom
    Mailer.FromName = FromName
    Mailer.AddAddress MailtoAddress, MailtoAddress
    Mailer.Subject = Subject
    Mailer.Body = MailBody
    Mailer.Send
    If Err Then
        Aspemail = "<li>邮件发送失败!</li>"
        Err.Clear
        Exit Function
    End If
    Aspemail = ""
End Function

Private Function WebEasyMail()
    On Error Resume Next
    Dim mailsend
    Set mailsend = Server.CreateObject("easymail.MailSend")
    mailsend.CreateNew MailFrom, "temp"
    mailsend.MailName = FromName    ' 发信人名称
    mailsend.EM_To = MailtoAddress  ' 收件人邮件地址
    mailsend.EM_Subject = Subject
    mailsend.EM_Text = MailBody     ' 输入纯文本格式邮件内容
    mailsend.EM_HTML_Text = MailBody    ' 输入html邮件内容
    mailsend.useRichEditer = True
    If mailsend.Send() = False Then
        WebEasyMail = "<li>邮件发送失败，请检查WebEasyMail设置！</li>"
    Else
        WebEasyMail = ""
    End If
    Set mailsend = Nothing
End Function

Function TimeDelaySeconds(DelaySeconds)
    Dim SecCount, Sec2
    Dim Sec1
    SecCount = 0
    Sec2 = 0
    While SecCount < DelaySeconds + 1
        Sec1 = Second(Time())
        If Sec1 <> Sec2 Then
            Sec2 = Second(Time())
            SecCount = SecCount + 1
        End If
    Wend
End Function

End Class
%>
