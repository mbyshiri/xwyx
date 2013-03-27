<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Sub ShowUserErr()
    Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"
    Response.Write "<body>"
    Response.Write "<user>" & UserName & "</user>"
    Response.Write "<checkstat>err</checkstat>"
    Response.Write "<errsource>" & ErrMsg & "</errsource>"
    If EnableCheckCodeOfLogin = True Then
        Response.Write "<checkcode>1</checkcode>"
    Else
        Response.Write "<checkcode>0</checkcode>"
    End If
    Response.Write "<syskey/><apiurl/><savecookie/>"
    Response.Write "</body>"
End Sub

Sub ShowUserXml(WriteAPI)
    Dim UserPassword
    UserPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("UserPassword")))
    Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"
    Response.Write "<body>"
    Response.Write "<user>" & UserName & "</user>"
    Response.Write "<userid>" & UserID & "</userid>"
    Response.Write "<userpass>" & UserPassword & "</userpass>"
    Response.Write "<usertype>" & UserType & "</usertype>"
    Response.Write "<groupname>" & GroupName & "</groupname>"
    Response.Write "<grouptype>" & GroupType & "</grouptype>"
    Response.Write "<checkstat>ok</checkstat>"
    Response.Write "<balance>" & Balance & "</balance>"
    Response.Write "<exp>" & UserExp & "</exp>"
    Response.Write "<point>"
    Response.Write "    <pointname>" & PointName & "</pointname>"
    Response.Write "    <userpoint>" & UserPoint & "</userpoint>"
    Response.Write "    <unit>" & PointUnit & "</unit>"
    Response.Write "</point>"
    If UserChargeType = 0 Then
        Response.Write "<day>noshow</day>"
    ElseIf ValidNum = -1 Then
        Response.Write "<day>unlimit</day>"
    Else
        Response.Write "<day>" & ValidDays & "</day>"
    End If
    If Trim(UnsignedItems & "") = "" Then
        Response.Write "<article>0</article>"
    Else
        Dim UnsignedItemNum, arrUser
        arrUser = Split(UnsignedItems, ",")
        UnsignedItemNum = UBound(arrUser) + 1
        Response.Write "<article>" & UnsignedItemNum & "</article>"
    End If
    Response.Write "<logined>" & LoginTimes & "</logined>"
    If UnreadMsg > 0 Then
        Response.Write "<message>" & UnreadMsg & "</message>"
        Dim MessageID, rsMessage
        Set rsMessage = Conn.Execute("select id,sender,title,sendtime from PE_Message where incept='" & UserName & "' and delR=0 and flag=0 and IsSend=1")
        If rsMessage.bof And rsMessage.EOF Then
            Response.Write "<unreadmessage><stat>empty</stat></unreadmessage>"
        Else
            Response.Write "<unreadmessage><stat>full</stat>"
            Do While Not rsMessage.EOF
                Response.Write "<item>"
                Response.Write "<id>" & rsMessage("id") & "</id>"
                Response.Write "<sender>" & rsMessage("sender") & "</sender>"
                Response.Write "<title>" & rsMessage("title") & "</title>"
                Response.Write "<time>" & rsMessage("sendtime") & "</time>"
                Response.Write "</item>"
                rsMessage.movenext
            Loop
            Response.Write "</unreadmessage>"
        End If
        rsMessage.Close
        Set rsMessage = Nothing
    Else
        Response.Write "<message>0</message>"
        Response.Write "<unreadmessage><stat>empty</stat></unreadmessage>"
    End If
    If API_Enable = True And WriteAPI = True Then
        sPE_Items(conSyskey, 1) = MD5(UserName & API_Key, 16)
        Response.Write "<syskey>" & sPE_Items(conSyskey, 1) & "</syskey>"
        Dim intIndex
        For intIndex = 0 To UBound(arrUrlsSP2)
            Response.Write "<apiurl><![CDATA[" & arrUrlsSP2(intIndex) & "]]></apiurl>"
        Next
        Response.Write "<savecookie>" & CookieDate & "</savecookie>"
    Else
        Response.Write "<syskey/><apiurl/><savecookie/>"
    End If
    Response.Write "</body>"
End Sub
%>
