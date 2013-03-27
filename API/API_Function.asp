<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************


Dim sMyXmlDoc, sMyXmlHTTP
'����һ����ά���飬��Ԫ������ֵ��������
ReDim sPE_Items(31, 1)
sPE_Items(0, 0) = "appid"
sPE_Items(1, 0) = "action"
sPE_Items(2, 0) = "syskey"
sPE_Items(3, 0) = "status"
sPE_Items(4, 0) = "message"
sPE_Items(5, 0) = "username"
sPE_Items(6, 0) = "password"
sPE_Items(7, 0) = "email"
sPE_Items(8, 0) = "question"
sPE_Items(9, 0) = "answer"
sPE_Items(10, 0) = "savecookie"
sPE_Items(11, 0) = "truename"
sPE_Items(12, 0) = "gender"
sPE_Items(13, 0) = "birthday"
sPE_Items(14, 0) = "qq"
sPE_Items(15, 0) = "msn"
sPE_Items(16, 0) = "mobile"
sPE_Items(17, 0) = "telephone"
sPE_Items(18, 0) = "address"
sPE_Items(19, 0) = "zipcode"
sPE_Items(20, 0) = "homepage"
sPE_Items(21, 0) = "userip"
sPE_Items(22, 0) = "jointime"
sPE_Items(23, 0) = "experience"
sPE_Items(24, 0) = "ticket"
sPE_Items(25, 0) = "valuation"
sPE_Items(26, 0) = "balance"
sPE_Items(27, 0) = "posts"
sPE_Items(28, 0) = "userstatus"
sPE_Items(29, 0) = "province"
sPE_Items(30, 0) = "city"
sPE_Items(31, 0) = "sex"

sPE_Items(0, 1) = "powereasy"
sPE_Items(1, 1) = ""
sPE_Items(2, 1) = ""
sPE_Items(3, 1) = "0"
sPE_Items(4, 1) = "�����ѳɹ���ɣ�"
sPE_Items(5, 1) = ""
sPE_Items(6, 1) = ""
sPE_Items(7, 1) = ""
sPE_Items(8, 1) = ""
sPE_Items(9, 1) = ""
sPE_Items(10, 1) = ""
sPE_Items(11, 1) = ""
sPE_Items(12, 1) = ""
sPE_Items(13, 1) = ""
sPE_Items(14, 1) = ""
sPE_Items(15, 1) = ""
sPE_Items(16, 1) = ""
sPE_Items(17, 1) = ""
sPE_Items(18, 1) = ""
sPE_Items(19, 1) = ""
sPE_Items(20, 1) = ""
sPE_Items(21, 1) = ""
sPE_Items(22, 1) = ""
sPE_Items(23, 1) = ""
sPE_Items(24, 1) = ""
sPE_Items(25, 1) = ""
sPE_Items(26, 1) = ""
sPE_Items(27, 1) = ""
sPE_Items(28, 1) = ""
sPE_Items(29, 1) = ""
sPE_Items(30, 1) = ""
sPE_Items(31, 1) = ""

'�����������Ӧ�ĳ����������ڱ�д����ʱʹ��
Const conAppid = 0
Const conAction = 1
Const conSyskey = 2
Const conStatus = 3
Const conMessage = 4
Const conUsername = 5
Const conPassword = 6
Const conEmail = 7
Const conQuestion = 8
Const conAnswer = 9
Const conSavecookie = 10
Const conTruename = 11
Const conGender = 12
Const conBirthday = 13
Const conQQ = 14
Const conMsn = 15
Const conMobile = 16
Const conTelephone = 17
Const conAddress = 18
Const conZipcode = 19
Const conHomepage = 20
Const conUserip = 21
Const conJointime = 22
Const conExperience = 23
Const conTicket = 24
Const conValuation = 25
Const conBalance = 26
Const conPosts = 27
Const conUserstatus = 28
Const conProvince = 29
Const conCity = 30
Const conSex = 31

'**************************************************
'��������prepareXML(vIsQuest)
'��  �ã�����Ҫ���͵�����
'��  ����vIsQuest True=��������False=��Ӧ����
'**************************************************
Sub prepareXML(vIsQuest)
    'On Error Resume Next
    Dim TemplateFile, intIndex
    If vIsQuest Then
        TemplateFile = Server.MapPath(InstallDir & "API/Request.xml")
    Else
        TemplateFile = Server.MapPath(InstallDir & "API/Response.xml")
    End If
    If Not IsObject(sMyXmlDoc) Then createXmlDom
    sMyXmlDoc.Async = False
    sMyXmlDoc.Load (TemplateFile)
    If Err Then
        Err.Clear
        FoundErr = True
        ErrMsg = "����XMLģ���ļ�����"
    Else
        For intIndex = 0 To UBound(sPE_Items, 1)
            If vIsQuest Then
                '��������������������Ӧ��ר��Ԫ��
                If intIndex <> conStatus And intIndex <> conMessage Then
                    setNodeText sPE_Items(intIndex, 0), sPE_Items(intIndex, 1)
                End If
            Else
                '�������Ӧ���������������ר��Ԫ��
                If intIndex <> conAction And intIndex <> conSyskey And intIndex <> conUsername Then
                    setNodeText sPE_Items(intIndex, 0), sPE_Items(intIndex, 1)
                End If
            End If
        Next
    End If
End Sub

'**************************************************
'��������prepareData(vIsQuest)
'��  �ã���XML�л�ȡ�û���Ϣ
'��  ����vIsQuest True=�����ʽ��False=��Ӧ��ʽ
'**************************************************
Sub prepareData(vIsQuest)
    'On Error Resume Next
    Dim intIndex
    For intIndex = 0 To UBound(sPE_Items, 1)
        If vIsQuest Then
            '��������������������Ӧ��ר��Ԫ��
            If intIndex <> conStatus Or intIndex <> conMessage Then
                sPE_Items(intIndex, 1) = getNodeText(sPE_Items(intIndex, 0))
            End If
        Else
            '�������Ӧ���������������ר��Ԫ��
            If intIndex <> conSyskey Or intIndex <> conUsername Or intIndex <> conPassword Then
                sPE_Items(intIndex, 1) = getNodeText(sPE_Items(intIndex, 0))
            End If
        End If
    Next
End Sub

'**************************************************
'��������getNodeText
'��  �ã���ȡXML�ļ���ָ���ڵ���ı�
'��  ����strNodeName   ----�ڵ�����
'����ֵ�������������ı�ֵ��
'**************************************************
Function getNodeText(strNodeName)
    If IsNull(strNodeName) Or IsEmpty(strNodeName) Or strNodeName = "" Then Exit Function
    If IsNode(strNodeName) Then
        getNodeText = sMyXmlDoc.documentElement.getElementsByTagName(strNodeName).Item(0).Text  
    Else
        getNodeText = ""
    End If
End Function

'**************************************************
'��������setNodeText
'��  �ã�����XML�ļ���ָ���ڵ���ı�
'��  ����strNodeName   ----�ڵ�����
'��������strNodeText   ----Ҫ���õ��ı�
'����ֵ��0 = ���óɹ�; ���򷵻�Err.Description
'**************************************************
Function setNodeText(strNodeName, strNodeText)
    If IsNull(strNodeText) Or IsEmpty(strNodeText) Or strNodeText = "" Then Exit Function
    If IsNull(strNodeName) Or IsEmpty(strNodeName) Or strNodeName = "" Then Exit Function
    If IsNode(strNodeName) Then sMyXmlDoc.documentElement.getElementsByTagName(strNodeName).Item(0).text = strNodeText
End Function

'**************************************************
'��������IsNode
'��  �ã����һ��Node�Ƿ�������ı���Ϊ��
'��  ����strNodeName   ----�ڵ�����
'����ֵ��True or False
'**************************************************
Function IsNode(strNodeName)    
    IsNode = False   
    If strNodeName = "" Then Exit Function   
       If sMyXmlDoc.documentElement.getElementsByTagName(strNodeName).Item(0) Is Nothing Then
        IsNode = False   
    Else   
        IsNode = True   
    End If   
End Function 
'**************************************************
'��������createXmlDom
'��  �ã����������ܸ߰汾��MSXML����
'��  ������
'����ֵ��True - ����sMyXmlDoc�ɹ�
'��������False - ��������֧��MSXML����
'**************************************************
Function createXmlDom()
    'On Error Resume Next
    Set sMyXmlDoc = Server.CreateObject("MSXML2.FreeThreadedDOMDocument")
    If Err Then
        Err.Clear
        createXmlDom = False
        FoundErr = True
        ErrMsg = "��������֧��MSXML2.FreeThreadedDOMDocument����"
    Else
        createXmlDom = True
    End If
End Function
'**************************************************
'��������createXmlHTTP
'��  �ã����������ܸ߰汾��ServerXMLHTTP����
'��  ������
'����ֵ��True - ����sMyXmlDoc�ɹ�
'��������False - ��������֧��ServerXMLHTTP����
'**************************************************
Function createXmlHttp()
    'On Error Resume Next
    Set sMyXmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
    If Err Then
        createXmlHttp = False
        FoundErr = True
        ErrMsg = "��������֧��ServerXMLHTTP����"
    Else
        createXmlHttp = True
    End If
End Function

'**************************************************
'��������SendPost
'��  �ã�����Զ��ϵͳ��ͨѶ�����첽��ʽ��������
'��  ������
'**************************************************
Sub SendPost()
    If createXmlHttp Then
        sPE_Items(conUsername, 1) = getNodeText(sPE_Items(conUsername, 0))
        sPE_Items(conSyskey, 1) = MD5(sPE_Items(conUsername, 1) & API_Key, 16)
        setNodeText sPE_Items(conSyskey, 0), sPE_Items(conSyskey, 1)
        sMyXmlHTTP.setTimeouts API_Timeout, API_Timeout, API_Timeout * 6, API_Timeout * 6
        Dim intIndex
        For intIndex = 0 To UBound(arrUrlsSP2)
            sMyXmlHTTP.Open "POST", arrUrlsSP2(intIndex), False
            sMyXmlHTTP.setRequestHeader "Content-Type", "text/xml; charset=gb2312"
            sMyXmlHTTP.send sMyXmlDoc
            If sMyXmlHTTP.readyState = 4 And sMyXmlHTTP.Status = 200 Then
                'Response.Write BytesToBstr(sMyXmlHTTP.ResponseBody, "gb2312")
                Dim objRecXml
                Set objRecXml = Server.CreateObject("Microsoft.XMLDOM")
                objRecXml.Async = False
                objRecXml.Load (sMyXmlHTTP.ResponseXML)
                If Err Then
                    FoundErr = True
                    ErrMsg = "�û�����Ŀǰ�����á�[APIError-HTTP1-" & intIndex & "]"
                    Err.Clear
                ElseIf objRecXml.parseError.errorCode <> 0 Then
                    FoundErr = True
                    ErrMsg = "�û�����Ŀǰ�����á�[APIError-XmlParse-" & intIndex & "]"
                    Err.Clear
                Else
                    If objRecXml.documentElement.selectSingleNode("//status").Text <> "0" Then
                        FoundErr = True
                        ErrMsg = objRecXml.documentElement.selectSingleNode("//message").Text & " [APIError-API-" & intIndex & "]"
                    End If
                End If
            ElseIf sMyXmlHTTP.readyState = 4 And sMyXmlHTTP.Status <> 200 Then
                FoundErr = True
                'ErrMsg = "�û�����Ŀǰ�����ã� [APIError-HTTP2-" & intIndex & "]"
                ErrMsg = BytesToBstr(sMyXmlHTTP.ResponseBody, "gb2312")
            End If
            If FoundErr Then
                If intIndex > 0 then
                    RollbackUser intIndex
                End If
                Exit For
            End If
        Next
    Else
        FoundErr = True
        ErrMsg = "�û�����Ŀǰ�����ã� [APIError-HTTP-Runtime]"
    End If
End Sub

Sub RollbackUser(startIndex)
    startIndex = startIndex - 1
    Do While startIndex >= 0
        setNodeText "action", "delete"
        sMyXmlHTTP.Open "POST", arrUrlsSP2(startIndex), True
        sMyXmlHTTP.setRequestHeader "Content-Type", "text/xml; charset=gb2312"
        sMyXmlHTTP.send sMyXmlDoc
        startIndex = startIndex - 1
    Loop
End Sub

Sub WriteErrXml()
    Response.Clear
    Response.ContentType = "text/xml"
    Response.Charset = "gb2312"
    Response.Expires = 0
    Response.Expiresabsolute = Now() - 1
    Response.AddHeader "pragma", "no-cache"
    Response.AddHeader "cache-control", "private"
    Response.CacheControl = "no-cache"
    Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>" & vbNewLine
    Response.Write "<root><appid>powereasy</appid><status>1</status><body><message>" & ErrMsg & "</message></body></root>"
    Response.End
End Sub
Sub WriteXml()
    Response.Clear
    Response.ContentType = "text/xml"
    Response.Charset = "gb2312"
    Response.Expires = 0
    Response.Expiresabsolute = Now() - 1
    Response.AddHeader "pragma", "no-cache"
    Response.AddHeader "cache-control", "private"
    Response.CacheControl = "no-cache"
    Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>" & vbNewLine
    Response.Write sMyXmlDoc.documentElement.xml
    Response.End
End Sub

Function exchangeGender(iSex)
    If IsNull(iSex) Or iSex = "" Or Not IsNumeric(iSex) Then
        exchangeGender = 2
        Exit Function
    End If
    If iSex = "1" Then
        iSex = 0
    ElseIf iSex = "0" Then
        iSex = 1
    Else
        iSex = 2
    End If
End Function

Function AnsiToUnicode(ByVal str)
    Dim i, j, c, i1, i2, u, fs, f, p
    AnsiToUnicode = ""
    p = ""
    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        j = AscW(c)
        If j < 0 Then
            j = j + 65536
        End If
        If j >= 0 And j <= 128 Then
            If p = "c" Then
                AnsiToUnicode = " " & AnsiToUnicode
                p = "e"
            End If
            AnsiToUnicode = AnsiToUnicode & c
        Else
            If p = "e" Then
                AnsiToUnicode = AnsiToUnicode & " "
                p = "c"
            End If
            AnsiToUnicode = AnsiToUnicode & ("&#" & j & ";")
        End If
    Next
End Function

Function BytesToBstr(Body, Cset)
    Dim Objstream
    Set Objstream = Server.CreateObject("adodb.stream")
    Objstream.Type = 1
    Objstream.Mode = 3
    Objstream.Open
    Objstream.Write Body
    Objstream.Position = 0
    Objstream.Type = 2
    Objstream.Charset = Cset
    If Err.Number <> 0 Then
        Err.Clear
        Objstream.Close
        Set Objstream = Nothing
        BytesToBstr = "$False$"
        Exit Function
    End If
    BytesToBstr = Objstream.ReadText
    Objstream.Close
    Set Objstream = Nothing
End Function
%>
