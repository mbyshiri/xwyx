<!--#include file="Start.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

'===================================================================================================================
'����ʵ������һ��ģ����ָ��Ƶ������Ŀ�����۵��ã�֧���Զ���Ƶ��
'ModuleName     ģ�����ƣ�����Ϊ"Article","Soft","Photo","Product"
'ChannelID      ChannelID=0��ʾ����ָ��ģ�飨�����Զ���Ƶ�����µ�����,��ChannelIDΪ��ͬ��ֵ��Ӧ��ͬ��Ƶ��
'ClassID        ClassID=0��ʾ����ָ��Ƶ���µ�����,��ClassIDΪ��ֵͬ��Ӧ��ͬ����Ŀ
'Num            ��ʾ�������б��ʾ����������,Ϊ0��ʾ�г����з�������������
'Order          ��������1-��ʱ������ 2-��Ƶ��������Ŀ���� 3-Ƶ�������ʱ������ 4-����Ƶ����������ID����
'OpenUrl        ���ӵ�ַ��0-���ӵ����£�ͼƬ�������1-���ӵ������б�

'ShowPic        ����ͼƬ��־ 0-����ʾ 1-���ţ�2-ͼƬ����ʽһ��
'OpenType       ���£�ͼƬ������򿪷�ʽ��0Ϊ��ԭ���ڴ򿪣�1Ϊ���´��ڴ�
'ShowTime       ��ʾʱ�� 0-����ʾ 1-������+��ʱ�� 2-������ 3-ʱ�� 4-��ʽ�����ʱ��
'ShowUserName   �Ƿ���ʾ�û��� 0-����ʾ 1-��ʾ

'���磺
'1:��ʾArticleģ���е�����
'   ʵ�ֹ��ܣ�  ��ʾ����ģ�飨����������Ƶ��Ϊģ���½���Ƶ�����е�����
'   ���÷�ʽ:   NewComment.asp?ModuleName=Article&ShowUserName=1&ShowTime=2&OpenAddr=1&Order=3&num=30

'2:��ʾArticleģ����Ƶ��ID=1001������
'   ���÷�ʽ��  NewComment.asp?ModuleName=Article&ChannelID=1001&ShowUserName=1&ShowTime=2&OpenAddr=1&Order=3&num=30

'3:��ʾArticleģ����Ƶ��ID=1��ClassID=1������
'   ���÷�ʽ��  NewComment.asp?ModuleName=Article&ChannelID=1001&ClassID=1&ShowUserName=1&ShowTime=2&OpenAddr=1&Order=3&num=30

'4:��ʾ��������
'   ���÷�ʽ��  NewComment.asp?ShowUserName=1&ShowTime=2&OpenAddr=1&Order=3&num=30
'===================================================================================================================


Dim PEurl, opentype, strOrder, Content, OpenAddr
Dim sqlComment, rsComment, Titlelen, Num, Order, ChannelID, ClassID, ShowUserName, ShowTime, ShowPic, ShowContentLen
Dim ModuleName, ModuleId

PEurl = request.ServerVariables("HTTP_HOST") & request.ServerVariables("URL")
PEurl = GetServePath(PEurl)

ModuleName = Trim(request("ModuleName"))
ChannelID = PE_CLng(Trim(request("ChannelID")))
ClassID = PE_CLng(Trim(request("ClassID")))
Num = PE_CLng(Trim(request("Num")))
ShowPic = PE_CLng(Trim(request("ShowPic")))
ShowContentLen = PE_CLng(Trim(request("ShowContentLen")))
ShowUserName = PE_CLng(Trim(request("ShowUserName")))
ShowTime = PE_CLng(Trim(request("ShowTime")))
Titlelen = PE_CLng(Trim(request("Titlelen")))
opentype = PE_CLng(Trim(request("OpenType")))
OpenAddr = PE_CLng(Trim(request("OpenAddr")))

Select Case ModuleName
    Case "Article"
        ModuleId = 1
    Case "Soft"
        ModuleId = 2
    Case "Photo"
        ModuleId = 3
    Case "Product"
        ModuleId = 5
    Case Else
        ModuleName = "Article"
        ModuleId = 1
End Select

If Num = 0 Then Num = 10
If Titlelen = 0 Then Titlelen = 10
Select Case PE_CLng(Trim(request("Order")))
    Case 1
        strOrder = " order by WriteTime desc"
    Case 2
        If ClassID <> 0 Then
            strOrder = " order by C.ModuleType asc,A.ClassID desc,C.WriteTime desc"
        Else
            strOrder = " order by ModuleType asc,InfoID desc,WriteTime desc"
        End If
    Case 3
        strOrder = " order by ModuleType desc,WriteTime desc"
    Case 4
        strOrder = " order by ModuleType desc,C.CommentID desc"
    Case Else
        strOrder = " order by ModuleType desc"
End Select

If ModuleName <> "" Then
    If ChannelID <> 0 Then
        If ClassID <> 0 Then
            sqlComment = "Select top " & Num & " C.*,A.ChannelID from PE_Comment C left join PE_" & ModuleName & " A on C.InfoID=A." & ModuleName & "ID where A.ChannelID= " & ChannelID & " and A.ClassID= " & ClassID & " and C.Passed =" & PE_True '��ȡָ��ģ����ָ��Ƶ��ָ����Ŀ��ǰNum������
        Else
            sqlComment = "Select top " & Num & " C.*,A.ChannelID from PE_Comment C left join PE_" & ModuleName & " A on C.InfoID=A." & ModuleName & "ID where A.ChannelID= " & ChannelID & " and C.Passed =" & PE_True  '��ȡָ��ģ����ָ��Ƶ����ǰNum������
        End If
    Else
        sqlComment = "Select top " & Num & " C.*,A.ChannelID From PE_Comment C left join PE_" & ModuleName & " A on C.InfoID=A." & ModuleName & "ID where C.ModuleType= " & ModuleId & " and C.Passed =" & PE_True '��ȡָ��ģ���е�ǰNum������
    End If
Else
    sqlComment = "Select top " & Num & " * from PE_Comment where Passed =" & PE_True  '��ȡ����ģ���е�ǰNum������
End If
          
sqlComment = sqlComment & strOrder

Set rsComment = Server.CreateObject("ADODB.Recordset")
rsComment.open sqlComment, Conn, 1, 1
If rsComment.bof And rsComment.EOF Then
    Response.Write "document.write(' û���κ�����');"
Else
    Do While Not rsComment.EOF
        Content = rsComment("Content")
        If Len(Content) > Titlelen Then
            Content = Left(Content, Titlelen) & "..."
        End If
        Content = HTMLEncode(Content)
        Select Case ShowPic
            Case 0
            Case 1
                Response.Write "document.write('<font color=#b70000><b>��</b></font>');"
            Case 2
                Response.Write "document.write('<IMG src=" & PEurl & "Images/common1.gif border=0>');"
            Case 3
                Response.Write "document.write('<IMG src=" & PEurl & "Images/common2.gif border=0>');"
            Case 4
                Response.Write "document.write('<IMG src=" & PEurl & "Images/common3.gif border=0>');"
            Case 5
                Response.Write "document.write('<IMG src=" & PEurl & "Images/common4.gif border=0>');"
            Case 6
                Response.Write "document.write('<IMG src=" & PEurl & "Images/common5.gif border=0>');"
            Case 7
                Response.Write "document.write('<IMG src=" & PEurl & "Images/common6.gif border=0>');"
            Case 8
                Response.Write "document.write('<IMG src=" & PEurl & "Images/common7.gif border=0>');"
            Case 9
                Response.Write "document.write('<IMG src=" & PEurl & "Images/common8.gif border=0>');"
            Case 10
                Response.Write "document.write('<IMG src=" & PEurl & "Images/common9.gif border=0>');"
            Case Else
        End Select
        
        Response.Write "document.write('<a href=" & PEurl & GetChannelDir(rsComment("ChannelID"), rsComment("InfoID"), OpenAddr) & SetOpenType(opentype) & " Title=" & HTMLEncode(rsComment("Content")) & ">');"
        Response.Write "document.write('" & Content & "');"
        Response.Write "document.write('</a><I><font color=gray>');"
        
        Response.Write "document.write(' �� ');"
        If ShowUserName = 1 Then
            Response.Write "document.write('" & rsComment("UserName") & "��');"
        End If
        Select Case ShowTime
            Case 0
            Case 1      '�����ڸ�ʽ+��ʱ���ʽ
                Response.Write "document.write('<font color=green>" & FormatDateTime(rsComment("WriteTime"), 0) & "</font>');"
            Case 2      '�����ڸ�ʽ
                Response.Write "document.write('<font color=green>" & TransformDay(FormatDateTime(rsComment("WriteTime"), 2)) & "</font>');"
            Case 3      'ʱ��
                Response.Write "document.write('<font color=green>" & FormatDateTime(rsComment("WriteTime"), 4) & "</font>');"
            Case 4      '��ʽ�����ʱ��
                Response.Write "document.write('<font color=green>" & TransformTime(rsComment("WriteTime")) & "</font>');"
            Case Else
        End Select

        Response.Write "document.write('</font></I><br>');"
        rsComment.movenext
    Loop
End If
rsComment.Close
Set rsComment = Nothing


Function HTMLEncode(ByVal fString)
    If Not IsNull(fString) Then
        fString = Replace(fString, ">", "&gt;")
        fString = Replace(fString, "<", "&lt;")

        fString = Replace(fString, Chr(32), "&nbsp;")
        fString = Replace(fString, Chr(9), "&nbsp;")
        fString = Replace(fString, Chr(34), "&quot;")
        fString = Replace(fString, Chr(39), "&#39;")
        fString = Replace(fString, Chr(13), "")
        fString = Replace(fString, Chr(10) & Chr(10), "</P><P> ")
        fString = Replace(fString, Chr(10), "<BR> ")

        HTMLEncode = fString
    End If
End Function

Function SetOpenType(ByVal opentype)
    If opentype = 0 Then
        SetOpenType = " target=_self "
    Else
        SetOpenType = " target=_blank "
    End If
End Function

Function GetServePath(str)
    Dim tmpstr
    tmpstr = Split(str, "/")
    GetServePath = "http://" & Replace(str, tmpstr(UBound(tmpstr)), "")
End Function

Function PE_CLng(ByVal str1)
    If IsNumeric(str1) Then
        PE_CLng = CLng(str1)
    Else
        PE_CLng = 0
    End If
End Function

Function TransformDay(ByVal strDay)
    Dim strTemp
    If Not IsDate(strDay) Then
        TransformDay = ""
        Exit Function
    End If
    strTemp = Right("0" & Month(strDay), 2) & "-" & Right("0" & Day(strDay), 2)
    TransformDay = strTemp
End Function

Function GetChannelDir(ByVal ChannelID, ByVal InfoID, OpenAddr)
     Dim iModuleType, strRs, strTemp, ChannelDir
     Set strRs = Conn.execute("Select ModuleType,ChannelDir from PE_Channel where ChannelID = " & ChannelID & "")
     If Not strRs.EOF Then
        iModuleType = strRs(0)
        ChannelDir = strRs(1)
        Select Case iModuleType
            Case "1"
                If OpenAddr = 1 Then
                        strTemp = "/Comment.asp?ArticleID=" & InfoID & "&Action=ShowAll"
                Else
                        strTemp = "/ShowArticle.asp?ArticleID=" & InfoID
                End If
            Case 2
                If OpenAddr = 1 Then
                        strTemp = "/Comment.asp?SoftID=" & InfoID & "&Action=ShowAll"
                Else
                        strTemp = "/ShowSoft.asp?SoftID=" & InfoID
                End If
            Case 3
                If OpenAddr = 1 Then
                        strTemp = "/Comment.asp?PhotoID=" & InfoID & "&Action=ShowAll"
                Else
                        strTemp = "/ShowPhoto.asp?PhotoID=" & InfoID
                End If
            Case 5
                If OpenAddr = 1 Then
                        strTemp = "/Comment.asp?ProductID=" & InfoID & "&Action=ShowAll"
                Else
                        strTemp = "/ShowProduct.asp?ProductID=" & InfoID
                End If
        End Select
     End If
     GetChannelDir = ChannelDir & strTemp
End Function

Function TransformTime(ByVal GuestDatetime)
    If Not IsDate(GuestDatetime) Then Exit Function
    Dim thour, tminute, tday, nowday, dnt, dayshow, pshow
    thour = Hour(GuestDatetime)
    tminute = Minute(GuestDatetime)
    tday = DateValue(GuestDatetime)
    nowday = DateValue(Now)
    If thour < 10 Then
        thour = "0" & thour
    End If
    If tminute < 10 Then
        tminute = "0" & tminute
    End If
    dnt = DateDiff("d", tday, nowday)
    If dnt > 2 Then
        dayshow = Year(GuestDatetime)
        If (Month(GuestDatetime) < 10) Then
            dayshow = dayshow & "-0" & Month(GuestDatetime)
        Else
            dayshow = dayshow & "-" & Month(GuestDatetime)
        End If
        If (Day(GuestDatetime) < 10) Then
            dayshow = dayshow & "-0" & Day(GuestDatetime)
        Else
            dayshow = dayshow & "-" & Day(GuestDatetime)
        End If
        TransformTime = dayshow
        Exit Function
    ElseIf dnt = 0 Then
        dayshow = "���� "
    ElseIf dnt = 1 Then
        dayshow = "���� "
    ElseIf dnt = 2 Then
        dayshow = "ǰ�� "
    End If
    TransformTime = dayshow & pshow & thour & ":" & tminute
End Function

Conn.Close
Set Conn = Nothing
%>
