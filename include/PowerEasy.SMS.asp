<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Function PostSMS(SendNum, Content, Reserve)
    'On Error Resume Next
    Dim SendTiming, SendTime, MD5String
    If SendNum = "" Then
        PostSMS = "接收号码为空"
        Exit Function
    End If

    If Content = "" Then
        PostSMS = "短信内容为空"
        Exit Function
    End If

    SendTiming = "0"
    SendTime = ""
    Dim PE_MD5
    Set PE_MD5 = New Md5_Class
    MD5String = UCase(Trim(PE_MD5.MD5(SMSUserName & SMSKey & SendNum & Content & SendTiming & SendTime)))
    Set PE_MD5 = Nothing

    Err.Clear
    Dim xmlHttp, HttpUrl, PostData
    HttpUrl = "http://sms.powereasy.net/MessageGate/MessageGate2.aspx"
    
    PostData = "UserName=" & Server.UrlEncode(SMSUserName) & "&MD5String=" & MD5String & "&SendTiming=" & SendTiming & "&SendTime=" & SendTime & "&SendNum=" & Server.UrlEncode(SendNum) & "&Content=" & Server.UrlEncode(Content) & "&Reserve=" & Server.UrlEncode(Reserve)
    Set xmlHttp = Server.CreateObject("MSXML2.XMLHTTP")
    If Err.Number <> 0 Then
        Err.Clear
        PostSMS = "不能创建MSXML2.XMLHTTP对象"
        Set xmlHttp = Nothing
        Exit Function
    End If

    xmlHttp.Open "POST", HttpUrl, False
    xmlHttp.setRequestHeader "Content-Length", Len(PostData)
    xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=gb2312"
    xmlHttp.Send PostData
    If Err.Number <> 0 Or xmlHttp.Readystate <> 4 Then
        Err.Clear
        PostSMS = "MSXML2.XMLHTTP对象错误"
        Set xmlHttp = Nothing
        Exit Function
    End If
    PostSMS = xmlHttp.responseText
    
    Set xmlHttp = Nothing
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Function
%>
