<!--#include file="CommonCode.asp"-->
<!--#include file="../Include/PowerEasy.SendMail.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim MailType

Select Case MailObject
Case 0
    FoundErr = True
    ErrMsg = ErrMsg & "�Բ��𣬷�����û��ѡ���κ��ʼ�������������Բ���ʹ�ñ����ܡ�"
Case 1
    If Not IsObjInstalled("JMail.Message") Then
        FoundErr = True
        ErrMsg = ErrMsg & "JMail�ʼ��������û�а�װ�����Բ���ʹ�ñ����ܡ�"
    End If
Case 2
    If Not IsObjInstalled("CDONTS.NewMail") Then
        FoundErr = True
        ErrMsg = ErrMsg & "CDONTS�ʼ��������û�а�װ�����Բ���ʹ�ñ����ܡ�"
    End If
Case 3
    If Not IsObjInstalled("Persits.MailSender") Then
        FoundErr = True
        ErrMsg = ErrMsg & "ASPEMAIL�ʼ��������û�а�װ�����Բ���ʹ�ñ����ܡ�"
    End If
Case 4
    If Not IsObjInstalled("easymail.mailsend") Then
        FoundErr = True
        ErrMsg = ErrMsg & "WebEasyMail�ʼ��������û�а�װ�����Բ���ʹ�ñ����ܡ�"
    End If
Case Else
    FoundErr = True
    ErrMsg = ErrMsg & "�Բ��𣬷������ʼ�����������ԣ����Բ���ʹ�ñ����ܡ�"
End Select

ArticleID = PE_CLng(Request("ArticleID"))
If ArticleID = 0 Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>��ָ��Ҫ���͸����ѵ�����ID��</li>"
End If
If UserLogined = False Then
    FoundErr = True
    ErrMsg = ErrMsg & "<br>&nbsp;&nbsp;&nbsp;&nbsp;�㻹ûע�᣿����û�е�¼��ֻ�б�վ��ע���û�����ʹ�á����ߺ��ѡ����ܣ�<br><br>&nbsp;&nbsp;&nbsp;&nbsp;����㻹ûע�ᣬ��Ͻ�<a href='../Reg/User_Reg.asp'><font color=red>���ע��</font></a>�ɣ�<br><br>&nbsp;&nbsp;&nbsp;&nbsp;������Ѿ�ע�ᵫ��û��¼����Ͻ�<a href='../User/User_Login.asp'><font color=red>��˵�¼</font></a>�ɣ�<br><br>"
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
        ErrMsg = ErrMsg & "�Ҳ�������"
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
        ErrMsg = ErrMsg & "<li>����������Ϊ�գ�</li>"
        FoundErr = True
    End If
    If IsValidEmail(MailtoAddress) = False Then
        ErrMsg = ErrMsg & "<li>�����˵�Email��ַ�д���</li>"
        FoundErr = True
    End If
    If FoundErr Then Exit Sub

    Dim rs, sql, strContent
    sql = "Select A.ChannelID,A.Title,A.Content,A.UpdateTime,A.Author,A.InfoPoint,C.ClassPurview from PE_Article A left join PE_Class C on A.ClassID=C.ClassID where A.ArticleID=" & ArticleID & ""
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 1
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "�Ҳ�������"
    Else
        Subject = Replace(Replace("��������{$UserName}��{$SiteName}��������������������", "{$UserName}", UserName), "{$SiteName}", SiteName)
        If rs("ClassPurview") > 0 Or rs("InfoPoint") > 0 Then
            strContent = "<a href='" & Trim(Request.ServerVariables("HTTP_HOST")) & ChannelUrl_ASPFile & "/ShowArticle.asp?ArticleID=" & ArticleID & "'>����鿴��ҳ�������</a>"
        Else
            strContent = Replace(Replace(rs("Content") & "", "[InstallDir_ChannelDir]", Trim(Request.ServerVariables("HTTP_HOST")) & ChannelUrl & "/"), "{$UploadDir}", UploadDir)
        End If
        MailBody = Replace(Replace(Replace(Replace(Replace(Replace("<style>A:visited {  TEXT-DECORATION: none   }A:active  { TEXT-DECORATION: none   }A:hover   { TEXT-DECORATION: underline overline }A:link    { text-decoration: none;}A:visited { text-decoration: none;}A:active  { TEXT-DECORATION: none;}A:hover   { TEXT-DECORATION: underline overline}BODY   {    FONT-FAMILY: ����; FONT-SIZE: 9pt;}TD     {    FONT-FAMILY: ����; FONT-SIZE: 9pt   }</style><TABLE border=0 width='95%' align=center><TBODY><TR><TD valign=middle align=top>--&nbsp;&nbsp;���ߣ�{$Author}<br>--&nbsp;&nbsp;����ʱ�䣺{$Time}<br><br>--&nbsp;&nbsp;{$title}<br>--&nbsp;&nbsp;{$Content}<br></TD></TR></TBODY></TABLE><center><a href='{$SiteUrl}'>{$SiteName}</a>", "{$Author}", rs("Author")), "{$Time}", rs("UpdateTime")), "{$title}", rs("title")), "{$Content}", strContent), "{$SiteUrl}", SiteUrl), "{$SiteName}", SiteName)
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
        Call WriteSuccessMsg("�Ѿ��ɹ��������·��͸���ĺ��ѣ�", ComeUrl)
    Else
        FoundErr = True
    End If
End Sub
%>
