<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim rsDownloadErrList, sqlDownloadErrList
Dim rs, sql, ServerName
Dim iTemp, UrlNum

ServerName = ""
Action = Request.Form("action")
SoftID = PE_CLng(Request("SoftID"))
If SoftID = 0 Then
    FoundErr = True
    ErrMsg = ErrMsg & "<li>��ָ����ȷ��SoftID!</li>"
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If
Select Case Action
Case "SaveDownErrorForm"
    Call SaveDownErrorForm
Case Else
    Call ShowDownErrorForm
End Select
Call CloseConn


Sub ShowDownErrorForm()
    Response.Write "<html><head><title>���ص�ַ����</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    Response.Write "<link href='" & InstallDir & "images/Style.css' rel='stylesheet' type='text/css'></head><body>" & vbCrLf
    Response.Write "<form name='myform' method='Post' action='DownError.asp'>"
    Response.Write "<table width='760' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' bgcolor='#0084FE'>"
    Response.Write "    <td width='10%'><font color=#FFFFFF size='2'><strong>ѡ ��</strong></font></td>"
    Response.Write "    <td width='90%' height='24'><font color=#FFFFFF size='2'><strong>�� �� �� ַ</strong></font></td>"
    Response.Write "  </tr>"
    Set rsDownloadErrList = Server.CreateObject("ADODB.Recordset")
    sqlDownloadErrList = "select ChannelID,DownloadUrl from PE_Soft where SoftID=" & SoftID
    rsDownloadErrList.Open sqlDownloadErrList, Conn, 1, 3
    If rsDownloadErrList.BOF And rsDownloadErrList.EOF Then
        rsDownloadErrList.Close
        Set rsDownloadErrList = Nothing
        Response.Write "  <tr class='tdbg'><td colspan='6' align='center'><br>����������ڻ��ѱ�ɾ����<br><br></td></tr></Table>"
        Exit Sub
    End If
    '�ж��Ƿ���ģʽ
    If InStr(rsDownloadErrList("DownloadUrl"), "@@@") > 0 Then
       sql = "select * from PE_DownServer where ChannelID=" & ChannelID
       Set rs = Server.CreateObject("ADODB.Recordset")
       rs.Open sql, Conn, 1, 3
       Do While Not rs.EOF
            iTemp = rs("ServerID")
            If rs("ShowType") = 0 Then
              ServerName = rs("ServerName")
            Else
              ServerName = "<img src='" & rs("ServerLogo") & "'>"
            End If
            Response.Write " <tr align='center' bgcolor='#F0F0F0' class='tdbg'  onmouseout=""this.style.backgroundColor='#F0F0F0'"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
            Response.Write "    <td><input name='UrlID' type='checkbox' id='UrlID' value=" & iTemp & ""
            Response.Write "></td>"
            Response.Write "    <td><font size='2'>" & ServerName & "</font></td>"
            Response.Write "</tr>"
            rs.MoveNext
       Loop
       rs.Close
       Set rs = Nothing
       Response.Write "</table>  "
       Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
       Response.Write "  <tr align='center'>"
       Response.Write "    <td><input name='action' type='hidden' id='action' value='SaveDownErrorForm'>"
       Response.Write "    <td><input name='UrlType' type='hidden' id='UrlType' value='ImgType'>"
    Else
        UrlNum = GetUrlIDNum(rsDownloadErrList("DownloadUrl"))
      'Response.Write "UrlNum=" & UrlNum
        For iTemp = 1 To UrlNum
            Response.Write " <tr align='center' class='tdbg' onmouseout=""this.style.backgroundColor='#F0F0F0'"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
            Response.Write "    <td><input name='UrlID' type='checkbox' id='UrlID' value=" & iTemp & ""
            Response.Write "></td>"
            Response.Write "    <td><font size='2'>" & GetDownUrlName(rsDownloadErrList("DownloadUrl"), iTemp) & "</font></td>"
            Response.Write "</tr>"
        Next
        rsDownloadErrList.Close
        Set rsDownloadErrList = Nothing
        Response.Write "</table>  "
        Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
        Response.Write "  <tr align='center'>"
        Response.Write "    <td><input name='action' type='hidden' id='action' value='SaveDownErrorForm'>"
        Response.Write "    <td><input name='UrlType' type='hidden' id='UrlType' value='ComType'>"
        Response.Write "    <td><input name='UrlNum' type='hidden' id='UrlNum' value=" & UrlNum & ">"
    End If
    Response.Write "    <td><input name='ChannelID' type='hidden' id='ChannelID' value=" & ChannelID & ">"
    Response.Write "    <td><input name='SoftID' type='hidden' id='SoftID' value=" & SoftID & ">"
    Response.Write "    <input name='Submit' type='submit' id='Submit' value='�� ��'>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='Cancel' value='ȡ ��' onClick=""window.location.href='ShowSoft.asp?SoftID=" & SoftID & "'"">"
    Response.Write "  </tr>"
    Response.Write "</table></form>"
    Response.Write "</body></html>"
End Sub

Sub SaveDownErrorForm()
    Dim iUrlID, strUrlType, iUrlNum, iTemp, Times, ID, arrUrlID
    Dim sqlDown, sqlDownError, sqlDownErrorF, sqlDownErrorN
    Dim rsDown, rsDownError, rsDownErrorF, rsDownErrorN
   
    iUrlID = ReplaceBadChar(Trim(Request.Form("UrlID")))
    strUrlType = ReplaceBadChar(Trim(Request.Form("UrlType")))
    iUrlNum = PE_CLng(Request.Form("UrlNum"))
    If iUrlID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ�����ش�����ϢID</li>"
        Response.Write ErrMsg
        Exit Sub
    End If
    If InStr(iUrlID, ",") > 0 Then
         arrUrlID = Split(iUrlID, ",")
         Times = UBound(arrUrlID) + 1
    Else
        Times = 1
    End If
    sqlDown = "select * from PE_DownError"
    Set rsDown = Server.CreateObject("ADODB.Recordset")
    rsDown.Open sqlDown, Conn, 1, 3
    If rsDown.BOF And rsDown.EOF Then '��һ����PE_DownError¼������
        For iTemp = 1 To Times
            Conn.Execute ("update PE_Soft set ErrorTimes=ErrorTimes+1 where SoftID=" & SoftID & "And ChannelID=" & ChannelID & "")
            sqlDownErrorF = "select * from PE_DownError"
            Set rsDownErrorF = Server.CreateObject("ADODB.Recordset")
            rsDownErrorF.Open sqlDownErrorF, Conn, 1, 3
            rsDownErrorF.addnew
            rsDownErrorF("ErrorID") = PE_CLng(Conn.Execute("select max(ErrorID) from PE_DownError")(0)) + 1
            rsDownErrorF("ChannelID") = ChannelID
            rsDownErrorF("InfoID") = SoftID
            If strUrlType = "ImgType" Then
                If Times > 1 Then
                    rsDownErrorF("UrlID") = PE_CLng(arrUrlID(iTemp - 1)) '����ģʽ��ΪServerID
                Else
                    rsDownErrorF("UrlID") = PE_CLng(Times) '������ģʽTimes=1ʱ��arrUrlID(0)�����ڵ����
                End If
            Else
                rsDownErrorF("UrlID") = PE_CLng(iTemp)
            End If
            rsDownErrorF("ErrorTimes") = 1
            rsDownErrorF.Update
        Next
        rsDownErrorF.Close
        Set rsDownErrorF = Nothing
    Else
        For iTemp = 1 To Times
            Conn.Execute ("update PE_Soft set ErrorTimes=ErrorTimes+1 where SoftID=" & SoftID & "And ChannelID=" & ChannelID & "")
            sqlDownError = "select * from PE_DownError where InfoID=" & SoftID
            If strUrlType = "ImgType" Then '����ģʽ��ΪServerID
                If Times > 1 Then
                    sqlDownError = sqlDownError & " And UrlID=" & PE_CLng(arrUrlID(iTemp - 1)) & "And ChannelID=" & ChannelID
                Else
                    sqlDownError = sqlDownError & " And UrlID=" & PE_CLng(Times) & "And ChannelID=" & ChannelID
                End If
            Else
                sqlDownError = sqlDownError & " And UrlID=" & PE_CLng(iTemp) & "And ChannelID=" & ChannelID
            End If
            Set rsDownError = Server.CreateObject("ADODB.Recordset")
            rsDownError.Open sqlDownError, Conn, 1, 3
            If rsDownError.BOF And rsDownError.EOF Then
                rsDownError.addnew
                rsDownError("ErrorID") = PE_CLng(Conn.Execute("select max(ErrorID) from PE_DownError")(0)) + 1
                rsDownError("ChannelID") = ChannelID
                rsDownError("InfoID") = SoftID
                If strUrlType = "ImgType" Then
                    If Times > 1 Then
                        rsDownError("UrlID") = PE_CLng(arrUrlID(iTemp - 1)) '����ģʽ��ΪServerID
                    Else
                        rsDownError("UrlID") = PE_CLng(Times)
                    End If
                Else
                    rsDownError("UrlID") = PE_CLng(iTemp)
                End If
                rsDownError("ErrorTimes") = 1
                rsDownError.Update
                rsDownError.Close
                Set rsDownError = Nothing
            Else
                Conn.Execute ("update PE_DownError set ErrorTimes=ErrorTimes+1 where InfoID=" & SoftID & "And UrlID=" & iTemp & "And ChannelID=" & ChannelID & "")
            End If
        Next
    End If
    rsDown.Close
    Set rsDown = Nothing
    Call WriteSuccessMsg("���Ĵ��󱨸��Ѿ��ύ�ɹ����ǳ���л�������ı���", "")
End Sub

'=================================================
'��������GetUrlIDNum
'��  �ã�ȡ�����ص�ַ�У�������ַ�ĸ���
'��  ����DownloadUrls  ----���ص�ַ
'=================================================
Private Function GetUrlIDNum(DownloadUrl)
   Dim arrDownloadUrl
   If DownloadUrl = "" Then
      GetUrlIDNum = ""
      Exit Function
   End If

   If InStr(DownloadUrl, "$$$") > 1 Then
      arrDownloadUrl = Split(DownloadUrl, "$$$")
      GetUrlIDNum = UBound(arrDownloadUrl) + 1
   Else
      GetUrlIDNum = 1
   End If
End Function

'=================================================
'��������GetDownUrlName
'��  �ã�ȡ�����ص�ַ�о���ĳ��URLID��������ַ
'��  ����DownloadUrls  ----���ص�ַ
'        UrlID ----�����ı��ID
'=================================================
Private Function GetDownUrlName(DownloadUrls, ByVal UrlID)
     Dim DownloadUrl, arrDownloadUrls, arrUrlName, iTemp
     
     If DownloadUrls = "" Or UrlID = "" Then
        GetDownUrlName = ""
        Exit Function
     End If
    
   
    If InStr(DownloadUrls, "$$$") > 1 Then
        arrDownloadUrls = Split(DownloadUrls, "$$$")
        iTemp = UrlID - 1
       
         If iTemp >= 0 And iTemp <= UBound(arrDownloadUrls) Then
             arrUrlName = Split(arrDownloadUrls(iTemp), "|")
             GetDownUrlName = arrUrlName(0)
         End If
       Else
       arrUrlName = Split(DownloadUrls, "|")
      GetDownUrlName = arrUrlName(0)
    End If
End Function
%>
