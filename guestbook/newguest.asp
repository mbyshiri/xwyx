<!--#include file="../Start.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

'KindId         KindId=0��ʾ�������������������ԣ�KindIdΪ��ͬ��ֵ��Ӧ��ͬ���KindId=10000ֻ��ʾ��������
'OnlyTitle      Ϊ0��ʾ�������Ժͻظ�,Ϊ1ֻ��ʾ�������ԣ�����ʾ�ظ�
'Num            ��ʾ�������б���ʾ��������������
'Titlelen       ���Գ��ȣ����Ա���ĳ��ȣ���ʾ���ٸ���
'Order          ���Ϊ0 ���������� 1 ������ʱ������

'ShowPic        ����ͼƬ��־ 0 ����ʾ 1 ���ţ�2 ͼƬ����ʽһ��
'ShowKindName   �Ƿ���ʾ�������    Ϊ0����ʾ,Ϊ1��ʾ
'ShowContentLen �Ƿ���ʾ������������ 0 ����ʾ 1 ��ʾ
'ShowTime       ��ʾʱ�� 0 ����ʾ 1 ������+��ʱ�� 2 ������ 3 ʱ�� 4 ��ʽ�����ʱ��
'ShowUserName   �Ƿ���ʾ�û��� 0 ����ʾ 1 ��ʾ

Dim PEurl
PEurl = request.ServerVariables("HTTP_HOST") & request.ServerVariables("URL")
PEurl = GetServePath(PEurl)

Dim sqlGuest, rsGuest, Title
Dim Titlelen, Num, Order, KindID, OnlyTitle, ShowKindName, ShowContentLen, ShowUserName, ShowTime, ShowPic

ShowPic = PE_CLng(Trim(request("ShowPic")))
ShowContentLen = PE_CLng(Trim(request("ShowContentLen")))
ShowUserName = PE_CLng(Trim(request("ShowUserName")))
ShowTime = PE_CLng(Trim(request("ShowTime")))
KindID = PE_CLng(Trim(request("KindID")))
ShowKindName = PE_CLng(Trim(request("ShowKindName")))
Num = PE_CLng(Trim(request("Num")))
Titlelen = PE_CLng(Trim(request("Titlelen")))

If Num = 0 Then Num = 10
If Titlelen = 0 Then Titlelen = 10
If PE_CLng(Trim(request("Order"))) = 1 Then
    Order = "GuestDatetime"
Else
    Order = "GuestMaxID"
End If

If PE_CLng(Trim(request("OnlyTitle"))) = 1 Then
    OnlyTitle = " and GuestID=TopicID"
Else
    OnlyTitle = ""
End If
Select Case KindID
    Case 0
        If ShowKindName = 0 Then
            sqlGuest = "select top " & Num & " * from PE_GuestBook where GuestIsPassed=" & PE_True & OnlyTitle & " Order by " & Order & " desc"
        Else
            sqlGuest = "select top " & Num & " * from PE_GuestBook B left join PE_GuestKind K on B.KindID=K.KindID where GuestIsPassed=" & PE_True & OnlyTitle & " Order by " & Order & " desc"
        End If
    Case 10000
        If ShowKindName = 0 Then
            sqlGuest = "select top " & Num & " * from PE_GuestBook where GuestIsPassed=" & PE_True & " and Quintessence=1 Order by " & Order & " desc"
        Else
            sqlGuest = "select top " & Num & " * from PE_GuestBook B left join PE_GuestKind K on B.KindID=K.KindID where GuestIsPassed=" & PE_True & " and Quintessence=1 Order by " & Order & " desc"
        End If
    Case Else
        If ShowKindName = 1 Then
            sqlGuest = "select top " & Num & " * from PE_GuestBook B left join PE_GuestKind K on B.KindID=K.KindID where GuestIsPassed=" & PE_True & OnlyTitle & " and B.KindID=" & KindID & " Order by " & Order & " desc"
        Else
            sqlGuest = "select top " & Num & " * from PE_GuestBook where GuestIsPassed=" & PE_True & OnlyTitle & " and KindID=" & KindID & " Order by " & Order & " desc"
        End If
End Select

Set rsGuest = Server.CreateObject("ADODB.Recordset")
rsGuest.open sqlGuest, Conn, 1, 1
If rsGuest.bof And rsGuest.EOF Then
    Response.Write "document.write(' û���κ�����');"
Else
    Do While Not rsGuest.EOF
        Title = rsGuest("GuestTitle")
        If Len(Title) > Titlelen Then
            Title = Left(Title, Titlelen) & "..."
        End If
        Title = HTMLEncode(Title)
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
        
        If ShowKindName = 1 Then
            If IsNull(rsGuest("KindName")) Then
                Response.Write "document.write('  ');"
            Else
                Response.Write "document.write('[" & rsGuest("KindName") & "]  ');"
            End If
        End If
        Response.Write "document.write('<a href=" & PEurl & "Guest_Reply.asp?TopicID=" & rsGuest("TopicID") & "  target=_blank Title=" & HTMLEncode(rsGuest("GuestTitle")) & ">');"
        Response.Write "document.write('" & Title & "');"
        Response.Write "document.write('</a><I><font color=gray>');"
        If ShowContentLen = 1 Then
            Response.Write "document.write('(" & rsGuest("GuestContentLength") & "��)');"
        End If
        If ShowUserName = 1 Or ShowTime = 2 Or ShowTime = 3 Or ShowTime = 4 Then
            Response.Write "document.write(' �� ');"
        End If
        If ShowUserName = 1 Then
            Response.Write "document.write('" & rsGuest("GuestName") & "��');"
        End If
        Select Case ShowTime
            Case 0
            Case 1      '�����ڸ�ʽ+��ʱ���ʽ
                Response.Write "document.write('<font color=green>" & FormatDateTime(rsGuest("GuestDatetime"), 0) & "</font>');"
            Case 2      '�����ڸ�ʽ
                Response.Write "document.write('<font color=green>" & TransformDay(FormatDateTime(rsGuest("GuestDatetime"), 2)) & "</font>');"
            Case 3      'ʱ��
                Response.Write "document.write('<font color=green>" & FormatDateTime(rsGuest("GuestDatetime"), 4) & "</font>');"
            Case 4      '��ʽ�����ʱ��
                Response.Write "document.write('<font color=green>" & TransformTime(rsGuest("GuestDatetime")) & "</font>');"
            Case Else
        End Select

        Response.Write "document.write('</font></I><br>');"
        rsGuest.movenext
    Loop
End If
rsGuest.Close
Set rsGuest = Nothing
Call CloseConn


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

%>
