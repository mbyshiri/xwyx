<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = False   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

Dim resultMsg
Dim CheckType
CheckType = Trim(Request("CheckType"))
If CheckType = "CheckNum" Then
    Call JudgeSameNum
Else
    Call JudgeSameTitle
End If
Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>����ظ�" & ModuleName & "</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "<base target='_self'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body>" & vbCrLf
Response.Write "<Table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border' Height='100%'>" & vbCrLf
Response.Write "<tr class='title' height='22'>" & vbCrLf
Response.Write "    <td align=center>" & vbCrLf
Response.Write "    <b>�����<b>" & vbCrLf
Response.Write "    </td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr class='tdbg'>" & vbCrLf
Response.Write "    <td valign='top' align='left'>" & vbCrLf
Response.Write resultMsg
Response.Write "    </td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr class='title' height='22'>" & vbCrLf
Response.Write "    <td align=center>" & vbCrLf
Response.Write "        <input type=button name='button1' id='button1' value='�رմ���' onClick='javascript:window.close();'>" & vbCrLf
Response.Write "    </td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "</Table>"
Response.Write "</body>" & vbCrLf
Response.Write "</html>" & vbCrLf

Sub JudgeSameNum()
    Dim strSql, rsTitle, Title
    Title = Trim(Request("Title"))
    If Title = "" Then
        resultMsg = "�������Ʒ�ı�Ų�����Ϊ�գ�"
        Exit Sub
    End If

    Title = ReplaceText(Title, 2)
    Title = PE_HTMLEncode(Title)
    strSql = "Select ProductNum,ProductName From PE_Product Where ProductNum='" & Title & "'"
    Set rsTitle = Server.CreateObject("Adodb.Recordset")
    rsTitle.Open strSql, Conn, 1, 1
    If rsTitle.EOF And rsTitle.BOF Then
        resultMsg = "��" & ModuleName & "���û�б�ʹ�ã�"
    Else
        resultMsg = "��" & ModuleName & "����Ѵ��ڣ��������:<br>"
        Do While Not rsTitle.EOF
            resultMsg = resultMsg & "<li>��ţ�" & rsTitle(0) & "&nbsp;&nbsp;��Ʒ��:" & rsTitle(1) & "</li><br>"
            rsTitle.movenext
        Loop
    End If
    rsTitle.Close
    Set rsTitle = Nothing
End Sub
'�жϱ����Ƿ��ظ�
Sub JudgeSameTitle()
    Dim strSql, Title, TableName, ModuleType
    Dim rsTitle
    Title = Trim(Request("Title"))
    ModuleType = PE_CLng(Trim(Request("ModuleType")))
    
    If Title = "" Then
        resultMsg = "��������µı��ⲻ����Ϊ�գ�"
        Exit Sub
    End If

    Title = ReplaceText(Title, 2)
    Title = PE_HTMLEncode(Title)
    Select Case ModuleType
        Case "1"
            strSql = "Select  Title  From PE_Article Where Title='" & Title & "'"
            ModuleName = "����"
        Case "2"
            strSql = "Select SoftName From PE_Soft Where SoftName = '" & Title & "'"
            ModuleName = "�����"
        Case "3"
            strSql = "Select PhotoName From PE_Photo Where PhotoName='" & Title & "'"
            ModuleName = "ͼƬ��"
        Case "5"
            strSql = "Select ProductName From PE_Product Where ProductName='" & Title & "'"
            ModuleName = "��Ʒ��"
        Case Else
            resultMsg = "����Ĵ��ݲ�����"
            Exit Sub
    End Select
    
    Set rsTitle = Server.CreateObject("Adodb.Recordset")
    rsTitle.Open strSql, Conn, 1, 1
    If rsTitle.EOF And rsTitle.BOF Then
        resultMsg = "��" & ModuleName & "û�б�ʹ�ã�"
    Else
        resultMsg = "��" & ModuleName & "�Ѵ��ڣ��������:<br>"
        Do While Not rsTitle.EOF
            resultMsg = resultMsg & "<li>" & rsTitle(0) & "</li><br>"
            rsTitle.movenext
        Loop
    End If
    rsTitle.Close
    Set rsTitle = Nothing
End Sub
%>
