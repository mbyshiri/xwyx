<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Sub CheckComeUrl()
    Dim ComeUrl, TrueSiteUrl, cUrl
    ComeUrl = Trim(Request.ServerVariables("HTTP_REFERER"))
    TrueSiteUrl = Trim(Request.ServerVariables("HTTP_HOST"))
    If ComeUrl = "" Then
        Response.Write "<br><p align=center><font color='red'>�Բ���Ϊ��ϵͳ��ȫ��������ֱ�������ַ���ʱ�ϵͳ�ĺ�̨����ҳ�档</font></p>"
        Call WriteEntry(1, "", "ֱ�ӵ�ַ������ʺ�̨")
        Response.End
    Else
        cUrl = Trim("http://" & TrueSiteUrl) & ScriptName
        If LCase(Left(ComeUrl, InStrRev(ComeUrl, "/"))) <> LCase(Left(cUrl, InStrRev(cUrl, "/"))) Then
            Response.Write "<br><p align=center><font color='red'>�Բ���Ϊ��ϵͳ��ȫ����������ⲿ���ӵ�ַ���ʱ�ϵͳ�ĺ�̨����ҳ�档</font></p>"
            Call WriteEntry(1, "", "�ⲿ���ӷ��ʺ�̨")
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
'��������CheckSecretCode
'��  �ã�Ч�鰲ȫ��
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
'��������SendMessage
'��  �ã����һ������Ϣ
'��  ����InceptUser ----�û�����
'        Title ---- ����Ϣ����
'        Content ---- ����Ϣ����
'        SendUser ---- ������
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
'��������Replace_CaseInsensitive
'��  �ã��滻�ַ�����Сд������
'��  ����expression ---- �ַ������ʽ ����Ҫ��������ַ���
'        find ---- �����������ַ���
'        replacewith ---- �����滻�����ַ���
'����ֵ���������ַ���
'**************************************************
Function Replace_CaseInsensitive(expression, find, replacewith)
    regEx.Pattern = find
    Replace_CaseInsensitive = regEx.Replace(expression, replacewith)
End Function
%>
