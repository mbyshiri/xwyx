<!--#include file="../Start.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

If CheckUserLogined() = False Then
    Call CloseConn
    Response.Redirect "User_Login.asp"
End If


Call Read
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Call CloseConn


Sub Read()
    Dim MessageID, rs, rsNext, NextID, NextSender
    
    MessageID = PE_CLng(Trim(Request("MessageID")))
    If MessageID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ָ���Ķ���ϢID����</li>"
        Exit Sub
    End If
    
    Conn.Execute ("update PE_Message set Flag=1 where Incept='" & UserName & "' and ID=" & MessageID)
    Conn.Execute ("update PE_User set UnreadMsg=UnreadMsg-1 where UserName='" & UserName & "'")
    Set rsNext = Conn.Execute("select ID,Sender from PE_Message where Incept='" & UserName & "' and Flag=0 and IsSend=1 and ID>" & MessageID & " order by SendTime")
    If Not (rsNext.BOF And rsNext.EOF) Then
        NextID = rsNext(0)
        NextSender = rsNext(1)
    End If
    Set rsNext = Nothing

    Set rs = Conn.Execute("select * from PE_Message where (Incept='" & UserName & "' or Sender='" & UserName & "') and ID=" & MessageID)
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���Ķ���Ϣ</li>"
        Set rs = Nothing
        Exit Sub
    End If

    Response.Write "<head>"
    Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />"
    Response.Write "<title>�Ķ�����Ϣ</title>"
    Response.Write "<link href=""../Skin/DefaultSkin.css"" rel=""stylesheet"" type=""text/css"">"
    Response.Write "</head>"
    Response.Write "<body  leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' align='center'><strong>�� �� �� �� Ϣ</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='center'>"
    Response.Write "      <a href='User_Message.asp?Action=Delete&MessageID=" & rs("ID") & "' target='_blank'><img src='images/m_delete.gif' border=0 alt='ɾ����Ϣ'></a> &nbsp; "
    Response.Write "      <a href='User_Message.asp?Action=New' target='_blank'><img src='images/m_to.gif' border=0 alt='������Ϣ'></a> &nbsp;"
    Response.Write "      <a href='User_Message.asp?Action=Re&touser={$sender}&MessageID=" & rs("ID") & "' target='_blank'><img src='images/m_re.gif' border=0 alt='�ظ���Ϣ'></a>&nbsp;"
    Response.Write "      <a href='User_Message.asp?Action=Fw&MessageID=" & rs("ID") & "' target='_blank'><img src='images/m_fw.gif' border=0 alt='ת����Ϣ'></a>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'><td><b>�� �� �ˣ�</b>" & rs("Sender") & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td><b>����ʱ�䣺</b>" & rs("SendTime") & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td><b>��Ϣ���⣺</b>" & PE_HTMLEncode(rs("Title")) & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td>" & FilterBadTag(rs("Content"), rs("Sender")) & "</td></tr>"
    If NextID <> "" Then
        Response.Write "  <tr class='tdbg'><td align='right'>"
        Response.Write "   <a href=User_Message.asp?Action=ReadMsg&MessageID=" & NextID & ">[��ȡ��һ����Ϣ]</a>"
        Response.Write "  </td></tr>"
    End If
    Response.Write "</table>"
    Response.Write "</body>"
    Response.Write "</html>"
    rs.Close
    Set rs = Nothing
End Sub
%>
