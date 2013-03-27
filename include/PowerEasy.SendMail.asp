<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
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
    'ѡ���ʼ��������
    Select Case MailObject
    Case 0
        FoundErr = True
        ErrMsg = ErrMsg & "û��ѡ���κ��ʼ�������������Բ��ܷ���ע��ȷ���š�"
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
        ErrMsg = ErrMsg & "�������ʼ�����������ԣ����Բ��ܷ���ע��ȷ���š�"
    End Select
    If ErrMsg <> "" Then
        Send = ErrMsg
    End If
End Function

Private Function JSendMail()
    On Error Resume Next
    Dim JMail
    Set JMail = Server.CreateObject("JMail.Message")
    JMail.Charset = "gb2312"        '�ʼ�����
    JMail.silent = True
    JMail.ContentType = "text/html"     '�ʼ����ĸ�ʽ
    'JMail.ServerAddress=MailServer     '���������ʼ���SMTP������
    '�����������ҪSMTP�����֤����ָ�����²���
    JMail.MailServerUserName = MailServerUserName    '��¼�û���
    JMail.MailServerPassWord = MailServerPassWord        '��¼����
    JMail.MailDomain = MailDomain       '����������á�name@domain.com���������û�����¼ʱ����ָ��domain.com
    JMail.AddRecipient MailtoAddress, MailtoName    '������
    JMail.Subject = Subject       '����
    'JMail.HtmlBody = MailBody     '�ʼ����ģ�HTML��ʽ��
    JMail.Body = MailBody        '�ʼ����ģ����ı���ʽ��
    JMail.FromName = FromName       '����������
    JMail.From = MailFrom         '������Email
    JMail.Priority = Priority            '�ʼ��ȼ���1Ϊ�Ӽ���3Ϊ��ͨ��5Ϊ�ͼ�
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
    'CDOMail.Cc = sCc   '���͸������ˣ�����ָ��������ö��Ÿ���
    'CDOMail.BCc = sBCc '�����������ˣ�����ָ��������ö��Ÿ���
    CDOMail.BodyFormat = 0   'ָ���ʼ�ΪMINE��ʽ
    CDOMail.MailFormat = 0   'ָ���ʼ�ΪHTML��ʽ
    CDOMail.Importance = 0   'ָ���ʼ����� 0: ��ͨ 1������ 2:����
    CDOMail.Body = MailBody
    CDOMail.Send
    Set CDOMail = Nothing
    If Err Then
        CdontsMail = "<li>�ʼ�����ʧ��!</li>"
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
    Mailer.UserName = MailServerUserName    '��¼�û���
    Mailer.Password = MailServerPassWord    '��¼����
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
        Aspemail = "<li>�ʼ�����ʧ��!</li>"
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
    mailsend.MailName = FromName    ' ����������
    mailsend.EM_To = MailtoAddress  ' �ռ����ʼ���ַ
    mailsend.EM_Subject = Subject
    mailsend.EM_Text = MailBody     ' ���봿�ı���ʽ�ʼ�����
    mailsend.EM_HTML_Text = MailBody    ' ����html�ʼ�����
    mailsend.useRichEditer = True
    If mailsend.Send() = False Then
        WebEasyMail = "<li>�ʼ�����ʧ�ܣ�����WebEasyMail���ã�</li>"
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
