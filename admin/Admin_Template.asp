<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.XmlHttp.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

Dim TemplateType, downright, TempType, ProjectName
Dim NavigationCss '�������
Dim IsOnlinePayment '���ͨ��ģ�幫������֧������
Dim TemplateProjectID, i

'������Ա����Ȩ��
If AdminPurview > 1 Then
    If ChannelID > 0 And ModuleType <> 4 Then
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "Template_" & ChannelDir)
    Else
        PurviewPassed = CheckPurview_Other(AdminPurview_Others, "Template")
    End If
    If PurviewPassed = False Then
        Response.Write "<br><p align='center'><font color='red' style='font-size:9pt'>�Բ�����û�д��������Ȩ�ޡ�</font></p>"
        Call WriteEntry(6, AdminName, "ԽȨ����")
        Response.End
    End If
End If

TemplateType = Trim(Request("TemplateType"))

downright = PE_CLng(Trim(Request("downright")))
ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))
TemplateProjectID = PE_CLng(Trim(Request("TemplateProjectID")))
IsOnlinePayment = PE_CLng(Trim(Request("IsOnlinePayment")))
NavigationCss = "title"

If ProjectName = "" Then
    Dim rs
    Set rs = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")
    If rs.BOF And rs.EOF Then
        Call WriteErrMsg("<li>ϵͳ�л�û��Ĭ�Ϸ���,�뵽��������ָ��Ĭ�Ϸ�����</li>", ComeUrl)
        Response.End
    Else
        ProjectName = rs("TemplateProjectName")
    End If
    Set rs = Nothing
End If

If TemplateType = "" Then
    TemplateType = 1
Else
    TemplateType = PE_CLng(TemplateType)
End If

TempType = PE_CLng(Trim(Request("TempType")))

Response.Write "<html><head><title>" & ChannelName & "����----ģ�����</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>"
Response.Write "</head>"
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"

If ChannelID = 0 Then
    If TempType = 0 Then
        Call ShowPageTitle(ProjectName & "����----ͨ��ģ�����", 10006)
    ElseIf TempType = 1 Then
        Call ShowPageTitle(ProjectName & "����----��վ��Աģ�����", 10006)
    End If
Else
    Call ShowPageTitle(ProjectName & "����----" & ChannelName & "ģ�����", 10006)
End If

Response.Write "      <tr class='tdbg'>"
Response.Write "        <td width='70' height='30'><strong>��������</strong></td><td>"

If TempType = 1 Then
    Response.Write "<a href='Admin_Template.asp?TemplateType=8&TempType=1&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'>"
Else
    Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'>"
End If

Response.Write "ģ�������ҳ</a> | <a href='Admin_Template.asp?ChannelID="
If ChannelID = 0 And IsOnlinePayment = 1 Then
    Response.Write "1000&Action=Add&TemplateType=" & TemplateType & "&ProjectName=" & Server.UrlEncode(ProjectName)
Else
    Response.Write ChannelID & "&Action=Add&TemplateType=" & TemplateType & "&ProjectName=" & Server.UrlEncode(ProjectName)
End If

If TempType > 0 Then Response.Write "&TempType=" & TempType
Response.Write "''>���ģ��</a>"
Response.Write " | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Import&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID
If TempType > 0 Then Response.Write "&TempType=" & TempType
Response.Write "'>����ģ��</a>"
Response.Write " | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Export&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID
If TempType > 0 Then Response.Write "&TempType=" & TempType
Response.Write "'>����ģ��</a>"
Response.Write " | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=ChannelCopyTemplate&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID
If TempType > 0 Then Response.Write "&TempType=" & TempType
Response.Write "'>Ƶ��ģ�帴��</a>"
Response.Write " | <a href='Admin_Template.asp?Action=BatchReplace&ChannelID=" & ChannelID & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID
If TempType > 0 Then Response.Write "&TempType=" & TempType
Response.Write "'>�����滻ģ�����</a>"
Response.Write " | <a href='Admin_Template.asp?Action=Main&ChannelID=" & ChannelID & "&downright=1&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID
If TempType > 0 Then Response.Write "&TempType=" & TempType
Response.Write "'>ģ�����վ����</a>"
Response.Write " | <a href='Admin_Template.asp?Action=BatchDefault&ProjectName=" & Server.UrlEncode(ProjectName)
If TempType > 0 Then Response.Write "&TempType=" & TempType
Response.Write "'>ģ��Ĭ����������</a>"
Response.Write "        </td>"
Response.Write "      </tr>"
Response.Write "    </table>"
Response.Write "    <br>"

If Action = "" Or Action = "SaveAdd" Or Action = "SaveModify" Or Action = "Main" Or Action = "main" Then
    
    Response.Write "<table width='100%' border='0' align='center' "
    If TemplateProjectID <> 0 Then
        Response.Write "cellpadding='2' cellspacing='1' class='border'"
    End If
    Response.Write ">"

    If TemplateProjectID <> 0 Then
        Response.Write "  <tr class='" & NavigationCss & "'>"
        Response.Write "    <td>"
        Response.Write "    | <a href='Admin_Template.asp?Action=Main&ChannelID=0&ProjectName=" & Server.UrlEncode(ProjectName) & "&TempType=0&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' "
        If ChannelID = 0 And TempType = 0 Then
            Response.Write " color='red'"
        End If
         Response.Write ">��վͨ��ģ��</FONT></a>"
        Response.Write "    | <a href='Admin_Template.asp?Action=Main&ChannelID=0&ProjectName=" & Server.UrlEncode(ProjectName) & "&TempType=1&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' "
        If ChannelID = 0 And TempType = 1 Then
            Response.Write " color='red'"
        End If
        Response.Write " >��Աģ��</FONT></a>"
        i = 0
        Set rs = Conn.Execute("SELECT DISTINCT t.ChannelID,c.ChannelName FROM PE_Template t INNER JOIN PE_Channel c ON t.ChannelID = c.ChannelID where c.Disabled=" & PE_False)
        If rs.BOF And rs.EOF Then
            Response.Write " û��ģ�����ȵ�����ģ��"
        Else
            Do While Not rs.EOF
                Response.Write "    | <a href='Admin_Template.asp?Action=Main&ChannelID=" & rs("ChannelID") & "&TemplateType=" & TemplateType & "&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TempType=" & TempType & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(rs("ChannelID"), ChannelID) & ">" & rs("ChannelName") & "Ƶ��ģ��</FONT></a>"
                If i > 3 Then
                    Response.Write " | </td><tr class='" & NavigationCss & "'><td>"
                    i = 0
                Else
                    i = i + 1
                End If

                rs.MoveNext
            Loop
            Response.Write " | "
        End If
        rs.Close
        Set rs = Nothing
        Response.Write "    </td>"
        Response.Write "  </tr>"
        NavigationCss = "tdbg"
    End If
    Response.Write "  <tr class='" & NavigationCss & "'>"
    Response.Write "    <td>"
    Response.Write "      <table width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "       <tr class='" & NavigationCss & "'><td>"
    
    If ChannelID > 0 Then
        If ModuleType = 4 Then
            Response.Write "    | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 1) & ">���԰�ģ��</FONT></a> | "
            Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=3&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 3) & ">��������ģ��</FONT></a> | "
            Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=4&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 4) & ">���Իظ�ģ��</FONT></a> | "
            Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=5&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 5) & ">��������ҳģ��</FONT></a> | "
        Else
            Select Case ModuleType
            Case 6
                Response.Write "    | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 1) & ">Ƶ����ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=2&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 2) & ">��Ŀģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=3&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 3) & ">����ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=4&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 4) & ">ר��ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=5&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 5) & ">����ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=6&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 6) & ">����" & ChannelShortName & "ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=7&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 7) & ">�Ƽ�" & ChannelShortName & "ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=8&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 8) & ">����" & ChannelShortName & "ҳģ��</FONT></a>"
                Response.Write "<tr class='" & NavigationCss & "'><td> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=16&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 16) & ">����" & ChannelShortName & "ҳģ��</FONT></a> | "
            Case 7  '*********************���ӷ���ģ�����********************
                Response.Write "| <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 1) & ">Ƶ����ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=2&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 2) & ">��Ŀģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=7&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 7) & ">�Ƽ�ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=8&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 8) & ">����ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=30&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 30) & ">��������ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=31&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 31) & ">��������ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=32&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 32) & ">������ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=33&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 33) & ">��������ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=34&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 34) & ">��������ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=5&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 5) & ">����ҳģ��</FONT></a> | "
            Case 8 '��Ƹģ��ģ�����
                Response.Write "    | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 1) & ">Ƶ����ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=5&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 2) & ">ְλ����ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=3&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 3) & ">ְλ����ҳģ��</FONT></a> | "
            Case Else
                Response.Write "    | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 1) & ">Ƶ����ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=2&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 2) & ">��Ŀģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=3&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 3) & ">����ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=4&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 4) & ">ר��ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=22&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 22) & ">ר���б�ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=5&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 5) & ">����ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=6&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 6) & ">����" & ChannelShortName & "ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=7&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 7) & ">�Ƽ�" & ChannelShortName & "ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=8&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 8) & ">����" & ChannelShortName & "ҳģ��</FONT></a>"
                Response.Write "<tr class='" & NavigationCss & "'><td> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=16&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 16) & ">����" & ChannelShortName & "ҳģ��</FONT></a> | "
            End Select
            If ModuleType = 1 Then
                Response.Write "  <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=17&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 17) & ">��ӡҳģ��</FONT></a>"
                Response.Write "    | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=20&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 20) & ">���ߺ���ҳģ��</FONT></a> | "
            ElseIf ModuleType = 5 Then
                Response.Write "   <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=9&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 9) & ">���ﳵģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=10&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 10) & ">����̨ģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=11&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 11) & ">����Ԥ��ҳģ��</FONT></a> | "
                Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=12&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 12) & ">�����ɹ�ģ��</FONT></a> | "
                'Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=13&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 13) & ">����֧����һ��ģ��</FONT></a> | "
                'Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=14&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 14) & ">����֧���ڶ���ģ��</FONT></a> | "
                'Response.Write "<tr class='" & NavigationCss & "'><td> | "
                'Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=15&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 15) & ">����֧��������ģ��</FONT></a>"
                Response.Write " <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=19&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 19) & ">�ؼ���Ʒҳģ��</FONT></a> | "
                Response.Write " <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=21&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 21) & ">�̳ǰ���ҳģ��</FONT></a> | "
            End If
        End If
        Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=0&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 0) & ">����ģ��</FONT></a> | </td></tr>"
    Else
        If TempType = 0 Then
            Response.Write " | <a href='Admin_Template.asp?TemplateType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 1) & ">��վ��ҳģ�塡</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=3&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 3) & ">��վ����ҳģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=4&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 4) & ">��վ����ҳģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=22&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 22) & ">�����б�ҳģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=5&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 5) & ">��������ҳģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=6&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 6) & ">��վ����ҳģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=7&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 7) & ">��Ȩ����ҳģ��</FONT></a>"

            Response.Write "</td></tr>"
            Response.Write "<tr class='" & NavigationCss & "'><td>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=10&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 10) & ">������ʾҳģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=11&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 11) & ">�����б�ҳģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=12&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 12) & ">��Դ��ʾҳģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=13&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "&IsOnlinePayment=0'><FONT style='font-size:12px' " & IsFontChecked2(TemplateType, 13, IsOnlinePayment, 0) & ">��Դ�б�ҳģ��</FONT></a>"
            If ShowAnonymous = True Then			
                Response.Write " | <a href='Admin_Template.asp?TemplateType=103&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "&IsOnlinePayment=0'><FONT style='font-size:12px' " & IsFontChecked2(TemplateType, 13, IsOnlinePayment, 0) & ">����Ͷ��ģ��</FONT></a>"	
            End If					
            Response.Write "</td></tr>"
            Response.Write "<tr class='" & NavigationCss & "'><td>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=14&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "&IsOnlinePayment=0'><FONT style='font-size:12px' " & IsFontChecked2(TemplateType, 14, IsOnlinePayment, 0) & ">������ʾҳģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=15&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "&IsOnlinePayment=0'><FONT style='font-size:12px' " & IsFontChecked2(TemplateType, 15, IsOnlinePayment, 0) & ">�����б�ҳģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=16&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 16) & ">Ʒ����ʾҳģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=17&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 17) & ">Ʒ���б�ҳģ��</FONT></a>"
            'Response.Write " | "

            'Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=13&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "&IsOnlinePayment=1'><FONT style='font-size:12px' " & IsFontChecked2(TemplateType, 13, IsOnlinePayment, 1) & ">����֧����һ��ģ��</FONT></a> | "
            'Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=14&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "&IsOnlinePayment=1'><FONT style='font-size:12px' " & IsFontChecked2(TemplateType, 14, IsOnlinePayment, 1) & ">����֧���ڶ���ģ��</FONT></a> | "
            'Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=15&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "&IsOnlinePayment=1'><FONT style='font-size:12px' " & IsFontChecked2(TemplateType, 15, IsOnlinePayment, 1) & ">����֧��������ģ��</FONT></a></td></tr>"
            
            Response.Write "<tr class='" & NavigationCss & "'><td>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=29&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 29) & ">ȫվר���б�ҳģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=30&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 30) & ">ȫվר��ҳģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=101&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 101) & ">�Զ����б�ģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=0&TempType=" & TempType & "&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 0) & " >����ģ��</FONT></a> | </td></tr>"
            Response.Write "</td></tr>"
        Else
            Response.Write " | <a href='Admin_Template.asp?TemplateType=8&TempType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 8) & ">��Ա��Ϣҳģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=9&TempType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 9) & ">��Ա�б�ҳģ��</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=18&TempType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 18) & ">��Աע��ҳģ�壨���Э�飩</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=19&TempType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 19) & ">��Աע��ҳģ�壨ע�����</FONT></a>"
            Response.Write " | <a href='Admin_Template.asp?TemplateType=21&TempType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 21) & ">��Աע��ҳģ�壨ע������</FONT></a>"
            If ShowUserModel = True Then			
                Response.Write " | <a href='Admin_Template.asp?TemplateType=102&TempType=1&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 102) & ">��Ա����ͨ��ģ��</FONT></a>"			
            End If				
            Response.Write "</td></tr>"
            Response.Write "<tr class='" & NavigationCss & "'><td>"
            Response.Write " | <a href='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=0&TempType=" & TempType & "&downright=" & downright & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateProjectID=" & TemplateProjectID & "'><FONT style='font-size:12px' " & IsFontChecked(TemplateType, 0) & " >����ģ��</FONT></a> | </td></tr>"
            Response.Write "</td></tr>"
        End If
    End If
    Response.Write "   </table>"
    Response.Write "  </td>"
    Response.Write " </tr>"
    Response.Write "</table>"
End If

'���ͨ��ģ������̳�����֧��
If IsOnlinePayment > 0 Then
    ChannelID = 1000
End If

Select Case Action
Case "Add"
    Call Add
Case "Modify"
    Call Modify
Case "SaveAdd", "SaveModify"
    Call Save
Case "Set"
    Call SetDefault
Case "Del"
    Call DelTemplate
Case "Export"
    Call Export
Case "DoExport"
    Call DoExport
Case "Import"
    Call Import
Case "Import2"
    Call Import2
Case "DoImport"
    Call DoImport
Case "DoTemplateCopy"
    Call DoTemplateCopy
Case "ChannelCopyTemplate"
    Call ChannelCopyTemplate
Case "DoCopy"
    Call DoCopy
Case "BatchReplace"
    Call BatchReplace
Case "DoBatchReplace"
    Call DoBatchReplace
Case "BatchDefault"
    Call BatchDefault
Case "DoBatchDefault"
    Call DoBatchDefault
Case Else
    Call main
End Select

If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


'=================================================
'��������Main
'��  �ã�ģ����ҳ
'=================================================
Sub main()
    Dim iTemplateType, i
    Dim sql, rs, TempType
    Dim TemplateSelect, TemplateSelectContent
    Dim rsProjectName, SysDefault 'ϵͳĬ��
    
    TempType = PE_CLng(Trim(Request.QueryString("TempType")))
    TemplateSelect = PE_CLng(Trim(Request.Form("TemplateSelect")))

    If TemplateSelect = 1 Then
        TemplateSelectContent = Trim(Request.Form("TemplateSelectContent"))

        If TemplateSelectContent = "" Then
            ErrMsg = ErrMsg & "<li>ģ���ѯ����Ϊ�գ�</li>"
            Call WriteErrMsg(ErrMsg, ComeUrl)
            Exit Sub
        End If
    End If

    '�õ�ϵͳ����Ĭ������
    Set rsProjectName = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")

    If rsProjectName.BOF And rsProjectName.EOF Then
        Call WriteErrMsg("<li>ϵͳ�л�û��Ĭ�Ϸ���,�뵽��������ָ��Ĭ�Ϸ�����</li>", ComeUrl)
        Exit Sub
    Else
        SysDefault = rsProjectName("TemplateProjectName")
    End If

    Set rsProjectName = Nothing

    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "    function CheckAll(thisform){" & vbCrLf
    Response.Write "        for (var i=0;i<thisform.elements.length;i++){" & vbCrLf
    Response.Write "            var e = thisform.elements[i];" & vbCrLf
    Response.Write "            if (e.Name != ""chkAll""&&e.disabled!=true&&e.zzz!=1)" & vbCrLf
    Response.Write "                e.checked = thisform.chkAll.checked;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function mysub()" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        esave.style.visibility=""visible"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "</script>" & vbCrLf

    If TemplateSelect = 1 Then
        If ProjectName = "" Then
            sql = "select * from PE_Template where ProjectName='' or ProjectName is null"
        ElseIf ProjectName = "���з���" Then
            sql = "select * from PE_Template"
        Else
            sql = "select * from PE_Template where ProjectName='" & ProjectName & "'"
        End If

    Else

        If TemplateType = 0 Then
            If TempType = 1 Then
                sql = "select * from PE_Template where ChannelID=" & ChannelID & " and TemplateType in (8,9,18,19,20,21)"
            Else
                sql = "select * from PE_Template where ChannelID=" & ChannelID
            End If

            If downright = 1 Then
                sql = sql & " and Deleted=" & PE_True
            Else
                sql = sql & " and Deleted=" & PE_False
            End If

        Else
            sql = "select * from PE_Template where ChannelID=" & ChannelID & " and TemplateType=" & TemplateType

            If downright = 1 Then
                sql = sql & " and Deleted=" & PE_True
            Else
                sql = sql & " and Deleted=" & PE_False
            End If
        End If

        If ProjectName = "" Then
            sql = sql & " and ProjectName='' or ProjectName is null order by TemplateType,TemplateID"
        ElseIf ProjectName = "���з���" Then
            sql = sql & " order by TemplateType,TemplateID"
        Else
            sql = sql & " and ProjectName='" & ProjectName & "' order by TemplateType,TemplateID"
        End If
    End If

    Set rs = Conn.Execute(sql)

    Response.Write "<form name='form1' method='post' action='Admin_Template.asp'>"

    If ChannelName = "" Then
        ChannelName = "ͨ��ģ��"
    End If

    Response.Write "<IMG SRC='images/img_u.gif' height='12'>�����ڵ�λ�ã�" & ProjectName & "&nbsp;&gt;&gt;&nbsp;"
    If downright = 1 Then
        Response.Write "ģ�����վ&nbsp;&gt;&gt;" & vbCrLf
    End If
    Response.Write ChannelName & "&nbsp;&gt;&gt;&nbsp;" & GetTemplateTypeName(TemplateType, ChannelID)

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "     <tr class='title' height='22'>"
    Response.Write "      <td width='30' align='center'><strong>ѡ��</strong></td>"
    Response.Write "      <td width='30' align='center'><strong>ID</strong></td>"
    Response.Write "      <td width='100' align='center'><strong>��������</strong></td>"
    Response.Write "      <td width='120' align='center'><strong>ģ������</strong></td>"
    Response.Write "      <td height='22' align='center'><strong>ģ������</strong></td>"

    If ProjectName = SysDefault Then
        Response.Write "      <td width='60' align='center'><strong>ϵͳĬ��</strong></td>"
    Else
        Response.Write "      <td width='60' align='center'><strong>����Ĭ��</strong></td>"
    End If

    Response.Write "      <td width='260' align='center'><strong>����</strong></td>"
    Response.Write "     </tr>"
    iTemplateType = 0
    i = 0

    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg'><td colspan='10' align='center' height='50'>��ģ�������л�û��ģ��</td></tr>"
    Else

        Do While Not rs.EOF

            If TemplateSelect <> 1 Or (TemplateSelect = 1 And InStr(rs("TemplateContent"), TemplateSelectContent) > 0) Then
                If i > 0 And rs("TemplateType") <> iTemplateType Then
                    Response.Write "<tr height='10'><td colspan='6'></td></tr>"
                End If

                iTemplateType = rs("TemplateType")
                i = i + 1
                Response.Write "    <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
                Response.Write "  <td width=""30"" align=""center"" height=""30"">" & vbCrLf
                Response.Write "    <input type=""checkbox"" value=" & rs("TemplateID") & " name=""TemplateID"""

                If rs("IsDefault") = True Or rs("IsDefaultInProject") = True Then Response.Write "disabled"
                Response.Write "> " & vbCrLf
                Response.Write "  </td>" & vbCrLf
                Response.Write "      <td width='30' align='center'>" & rs("TemplateID") & "</td>"
                Response.Write "      <td width='100' align='center'>" & rs("ProjectName") & "</td>"
                Response.Write "      <td width='120' align='center'>" & GetTemplateTypeName(rs("TemplateType"), ChannelID) & "</td>"
                Response.Write "      <td align='center'><a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Modify&ProjectName=" & Server.UrlEncode(rs("ProjectName")) & "&TemplateID=" & rs("TemplateID") & "'>" & rs("TemplateName") & "</a></td>"

                If ProjectName = SysDefault Then
                    Response.Write "      <td width='60' align='center'><b>"

                    If rs("IsDefault") = True Then
                        Response.Write "<FONT style='font-size:12px' color='#008000'>��</FONT>"
                    Else
                    End If

                    Response.Write "</td>"
                Else
                    Response.Write "      <td width='60' align='center'><b>"

                    If rs("IsDefaultInProject") = True Then
                        Response.Write "��"
                    Else
                    End If
                End If

                Response.Write "</td>"
                Response.Write "      <td width='260' align='center'>"

                If rs("Deleted") = True Then
                    If rs("IsDefault") = False Then
                        Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Del&TemplateID=" & rs("TemplateID") & "&downright=1" & "&ProjectName=" & Server.UrlEncode(ProjectName)

                        If TempType > 0 Then Response.Write "&TempType=" & TempType
                        Response.Write "' onClick=""return confirm('ȷ��Ҫ����ɾ���˰������ģ���𣿸�ģ��ɾ���󲻿ɻָ�,ɾ���˰������ģ���ԭʹ�ô˰������ģ������½���Ϊʹ��ϵͳĬ�ϰ������ģ�塣');"">����ɾ��ģ��</a>&nbsp;&nbsp;"
                    Else
                        Response.Write "<font color='gray'>����ɾ��ģ��&nbsp;&nbsp;</font>"
                    End If

                    Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Del&TemplateID=" & rs("TemplateID") & "&TemplateType=" & rs("TemplateType") & "&downright=3" & "&ProjectName=" & Server.UrlEncode(ProjectName)

                    If TempType > 0 Then Response.Write "&TempType=" & TempType
                    Response.Write "'>��ԭģ��</a><br>"
                Else

                    '��ΪϵͳĬ��
                    If ProjectName = SysDefault Then
                        If rs("IsDefault") = False And ProjectName = SysDefault Then
                            Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Set&DefaultType=1&TemplateType=" & rs("TemplateType") & "&TemplateID=" & rs("TemplateID") & "&ProjectName=" & Server.UrlEncode(ProjectName)

                            If TempType > 0 Then Response.Write "&TempType=" & TempType
                            Response.Write "'>&nbsp;��ΪϵͳĬ��</a>&nbsp;&nbsp;"
                        Else
                            Response.Write "<font color='gray'>&nbsp;��ΪϵͳĬ��&nbsp;&nbsp;</font>"
                        End If

                    Else

                        If rs("IsDefaultInProject") = False Then
                            Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Set&DefaultType=2&TemplateType=" & rs("TemplateType") & "&TemplateID=" & rs("TemplateID") & "&ProjectName=" & Server.UrlEncode(ProjectName)

                            If TempType > 0 Then Response.Write "&TempType=" & TempType
                            Response.Write "'>&nbsp;��Ϊ����Ĭ��</a>&nbsp;&nbsp;"
                        Else
                            Response.Write "<font color='gray'>&nbsp;��Ϊ����Ĭ��&nbsp;&nbsp;</font>"
                        End If
                    End If

                    Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Modify&ProjectName=" & Server.UrlEncode(rs("ProjectName")) & "&TemplateID=" & rs("TemplateID")

                    If TempType > 0 Then Response.Write "&TempType=" & TempType
                    Response.Write "'>�޸�ģ��</a>&nbsp;&nbsp;"
                    If rs("IsDefault") = False And rs("IsDefaultInProject") = False Then
                        Response.Write "<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Del&TemplateID=" & rs("TemplateID") & "&ProjectName=" & Server.UrlEncode(ProjectName)
                        If TempType > 0 Then Response.Write "&TempType=" & TempType
                        Response.Write "' onClick=""return confirm('ȷ��Ҫɾ���˰������ģ����ɾ��������Դӻ���վ��ԭ����');"">ɾ��ģ��</a>"
                    Else
                        Response.Write "<font color='gray'>ɾ��ģ��</font>"
                    End If

                    Response.Write "&nbsp;&nbsp;<a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=DoTemplateCopy&TemplateName=" & Server.UrlEncode(rs("TemplateName")) & "&TemplateType=" & rs("TemplateType") & "&TemplateID=" & rs("TemplateID") & "&ProjectName=" & Server.UrlEncode(ProjectName)

                    If TempType > 0 Then Response.Write "&TempType=" & TempType
                    Response.Write "'>����ģ��</a><br>"
                End If

                Response.Write " </td>"
                Response.Write "</tr>"
            End If

            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "</table><br>" & vbCrLf
    Response.Write "        <input name=""Action"" type=""hidden""  value=""Del"">   " & vbCrLf
    Response.Write "        <input name=""ChannelID"" type=""hidden""  value=" & ChannelID & ">" & vbCrLf

    If TempType > 0 Then Response.Write "        <input name=""TempType"" type=""hidden""  value=" & TempType & ">" & vbCrLf
    Response.Write "        <input name=""chkAll"" type=""checkbox"" id=""chkAll"" onclick=CheckAll(this.form) value=""checkbox"" >ѡ������ģ��" & vbCrLf
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf

    If downright = 1 Then
        Response.Write "        <input name='downright' value='1' type='hidden'>"
        Response.Write "        <input type=""submit"" value="" ����ɾ�� "" name=""Del"" onclick='return confirm(""ȷ��Ҫ����ɾ��ѡ�е�ģ���𣿳���ɾ���󲻿ɻָ���"");' >&nbsp;&nbsp;" & vbCrLf
        Response.Write "        &nbsp;&nbsp;" & vbCrLf
        Response.Write "        <input type=""submit"" value="" ��ջ���վ "" name=""Del"" onClick=""document.form1.downright.value='2'"">&nbsp;&nbsp;" & vbCrLf
        Response.Write "        &nbsp;&nbsp;" & vbCrLf
        Response.Write "        <input type=""submit"" value="" �� ԭ "" name=""Del"" onClick=""document.form1.downright.value='3'"">&nbsp;&nbsp;" & vbCrLf
        Response.Write "        <input type=""submit"" value="" ȫ����ԭ "" name=""Del"" onClick=""document.form1.downright.value='4'"">&nbsp;&nbsp;" & vbCrLf
    Else
        Response.Write "        <input type=""submit"" value=""����ɾ�� "" name=""Del"" onclick='return confirm(""ȷ��Ҫɾ��ѡ�е�ģ����ɾ��������Դӻ���վ��ԭ����"");' >&nbsp;&nbsp;" & vbCrLf
        Response.Write "        <input type=""submit"" value="" �������� "" name=""ChannelCopyTemplate"" onClick=""document.form1.Action.value='DoTemplateCopy'"">&nbsp;&nbsp;" & vbCrLf
        Response.Write "        <input type=""submit"" value="" �����滻 "" name=""BatchReplace"" onClick=""document.form1.Action.value='BatchReplace'"">&nbsp;&nbsp;" & vbCrLf
    End If

    Response.Write "                        <Input TYPE='hidden' Name='BatchTypeName' value='�ƶ�'>" & vbCrLf
    Response.Write "                        <Input TYPE='hidden' Name='ProjectName' value='" & ProjectName & "'>" & vbCrLf
    Response.Write "                        <Input TYPE='hidden' Name='TemplateProjectID' value='" & TemplateProjectID & "'>" & vbCrLf

    If downright = 0 Then
        If TemplateType > 0 Then
            Response.Write "<input type='button' name='buttonm' value='���" & GetTemplateTypeName(TemplateType, ChannelID) & "' onclick=""window.location.href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Add&TemplateType=" & TemplateType & "&ProjectName=" & ProjectName
            If TempType > 0 Then Response.Write "&TempType=" & TempType
            Response.Write "'"">"
        End If
    End If

    Response.Write "<br><br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write " <tr class=""tdbg"">"
    Response.Write "   <td width='20%' align='right'>"
    Response.Write "  ģ�����ݲ�ѯ��</td>"
    Response.Write "   <td width='30%'> <TEXTAREA NAME='TemplateSelectContent'  style='width:300px;height:40px' onMouseOver=""this.select()"" onClick=""javascript:{if (form1.TemplateSelectContent.value == '���ڲ�ѯ��������Ҫ���ҵ��ַ�')form1.TemplateSelectContent.value=''; };"">���ڲ�ѯ��������Ҫ���ҵ��ַ�</TEXTAREA><input name='TemplateSelect' value='0' type='hidden'></td>"
    Response.Write "   <td width='50%' align='left'> <input type='submit' value=' �� ѯ ' onClick=""document.form1.TemplateSelect.value='1';document.form1.Action.value='Main'"">&nbsp;&nbsp; <font color='blue'>ע��</font> �����ܿɲ�ѯ��Ӧ��������Щģ����ʹ�ù���</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"

End Sub

'=================================================
'��������CommonLabel
'��  �ã����ó��ú�����ǩ
'=================================================
Sub CommonLabel(ByVal TemplateType)
    Response.Write "        <table align='left' border='0' id='CommonLabel" & TemplateType & "' cellpadding='0' cellspacing='1' width='550' height='100%' >"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "           <td width='120'> ���ó���������ǩ:</td>"
    Response.Write "           <td width='1' bgcolor='#ACA899'></td><td width='2'></td>"
    Response.Write "           <td>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetArticleList','�����б�����ǩ',1,'GetList',800,700," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetArticleList.gif' border='0' width='18' height='18' alt='��ʾ���±������Ϣ'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetPicArticle','��ʾͼƬ���±�ǩ',1,'GetPic',700,500," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetPicArticle.gif' border='0' width='18' height='18' alt='��ʾͼƬ����'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetSlidePicArticle','��ʾ�õ�Ƭ���±�ǩ',1,'GetSlide',700,500," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetSlidePicArticle.gif' border='0' width='18' height='18' alt='��ʾ�õ�Ƭ����'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_CustomListLabel.asp','CustomListLable','�����Զ����б��ǩ',1,'GetArticleCustom',720,700," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetArticleCustom.gif' border='0' width='18' height='18' alt='�����Զ����б�'></a>"
    Response.Write "           </td>"
    Response.Write "           <td width='1' bgcolor='#ACA899'></td><td width='2'></td>"
    Response.Write "           <td>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetSoftList','�����б�����ǩ',2,'GetList',800,700," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetSoftList.gif' border='0' width='18' height='18' alt='��ʾ�������'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetPicSoft','��ʾͼƬ���ر�ǩ',2,'GetPic',700,500," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetPicSoft.gif' border='0' width='18' height='18' alt='��ʾͼƬ����'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetSlidePicSoft','��ʾ�õ�Ƭ���ر�ǩ',2,'GetSlide',700,500," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetSlidePicSoft.gif' border='0' width='18' height='18' alt='��ʾ�õ�Ƭ����'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_CustomListLabel.asp','CustomListLable','�����Զ����б��ǩ',2,'GetSoftCustom',720,700," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetSoftCustom.gif' border='0' width='18' height='18' alt='�����Զ����б�'></a>"
    Response.Write "           </td>"
    Response.Write "           <td width='1' bgcolor='#ACA899'></td><td width='2'></td>"
    Response.Write "           <td>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetPhotoList','ͼƬ�б�����ǩ',3,'GetList',800,700," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetPhotoList.gif' border='0' width='18' height='18' alt='��ʾͼƬ����'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetPicPhoto','��ʾͼƬͼ�ı�ǩ',3,'GetPic',700,550," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetPicPhoto.gif' border='0' width='18' height='18' alt='��ʾͼƬ'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetSlidePicPhoto','��ʾ�õ�ƬͼƬ��ǩ',3,'GetSlide',700,550," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetSlidePicPhoto.gif' border='0' width='18' height='18' alt='��ʾͼƬ�õ�Ƭ'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_CustomListLabel.asp','CustomListLable','ͼƬ�Զ����б��ǩ',3,'GetPhotoCustom',720,700," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetPhotoList.gif' border='0' width='18' height='18' alt='ͼƬ�Զ����б�'></a>"
    Response.Write "           </td>"
    Response.Write "           <td width='1' bgcolor='#ACA899'></td><td width='2'></td>"
    Response.Write "           <td>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetProductList','�̳��б�����ǩ',5,'GetList',800,750," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetProductList.gif' border='0' width='18' height='18' alt='��ʾ��Ʒ����'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetPicProduct','��ʾͼƬ�̳Ǳ�ǩ',5,'GetPic',700,600," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetPicProduct.gif' border='0' width='18' height='18' alt='��ʾ��ƷͼƬ'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_label.asp','GetSlidePicProduct','��ʾ�õ�Ƭ�̳Ǳ�ǩ',5,'GetSlide',700,460," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetSlidePicProduct.gif' border='0' width='18' height='18' alt='��ʾ��Ʒ�õ�Ƭ'></a>"
    Response.Write "            <a href=""javascript:SuperFunctionLabel('../Editor/editor_CustomListLabel.asp','CustomListLable','�̳��Զ����б��ǩ',5,'GetProductCustom',720,700," & TemplateType & ")"" ><img src='../Editor/images/LabelIco/GetProductCustom.gif' border='0' width='18' height='18' alt='��Ʒ�Զ����б�'></a>"
    Response.Write "           </td>"
    Response.Write "           <td width='1' bgcolor='#ACA899'></td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
End Sub
'=================================================
'��������ADD
'��  �ã����ģ��
'=================================================
Sub Add()

    Dim Num, strHead, Content
    Dim rsTemplateProject
    
    TemplateType = Request.QueryString("TemplateType")
    ProjectName = Request.QueryString("ProjectName")

    '����js���� num Ϊ �����С��
    If TemplateType = 2 Then
        Num = 2
    Else
        Num = 1
    End If
    
    '����ģ��Ԥ��ͷ�� �����ʱ�õ�
    
    strHead = "<html>" & vbCrLf
    strHead = strHead & "<head>" & vbCrLf
    strHead = strHead & "<title>��ģ�����</title>" & vbCrLf
    strHead = strHead & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
    strHead = strHead & "{$Skin_CSS} {$MenuJS}" & vbCrLf
    strHead = strHead & "</head>" & vbCrLf
    strHead = strHead & "<body leftmargin=0 topmargin=0 onmousemove='HideMenu()'>" & vbCrLf
    strHead = strHead & vbCrLf & "<!-- ��������Ҫ��ƵĴ��� -->" & vbCrLf
    strHead = strHead & vbCrLf & "</body>" & vbCrLf
    strHead = strHead & "</html>" & vbCrLf
        
    '�滻ͷ����ǩ Content Ϊ�滻��ͷ���ļ������ڱ༭����ʾcss
    Content = Replace(strHead, "{$Skin_CSS}", "<link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>")
    Content = Replace(Content, "{$MenuJS}", "<script language='JavaScript' type='text/JavaScript' src='" & InstallDir & "js/menu.js'></script>")
    Content = Replace(Content, "{$InstallDir}", InstallDir)
    'Ԥд3000��
    Dim strContenttemp, i
    For i = 1 To 3000
        If strContenttemp = "" Then
            strContenttemp = i & vbCrLf
        Else
            strContenttemp = strContenttemp & i & vbCrLf
        End If
    Next
   
    '����js�������
    Call StrJS_Template
            
    Response.Write "<form name='form1' method='post' action='Admin_Template.asp' >"
    Response.Write "  <table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>�� �� �� ģ ��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>&nbsp;&nbsp;<strong> ѡ�񷽰��� </strong><select name='ProjectName' id='ProjectName' onChange=""window.location.href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Add&TemplateType=" & TemplateType & "&TempType=" & TempType & "&ProjectName='+this.value"">" & GetProject_Option(ProjectName) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>&nbsp;&nbsp;<strong> ģ�����ͣ� </strong><select name='TemplateType' id='TemplateType' onChange=""window.location.href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Add&TempType=" & TempType & "&ProjectName=" & Server.UrlEncode(ProjectName) & "&TemplateType='+this.value"">" & GetTemplate_Option(PE_CLng(TemplateType)) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>&nbsp;&nbsp;<strong> ģ�����ƣ� </strong><input name='TemplateName' type='text' id='TemplateName' value='' size='50' maxlength='50'></td>"
    Response.Write "    </tr>"
    Response.Write "    <a name='#TemplateStart1'></a>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td valign='top'align=center id='Navigation1' style='display:'>"
    
    If TemplateType = 2 Then
        Response.Write "<b>����ģ�壺</b>����Ŀ��������Ŀʱ���ͻ���ô˴�������ʾ��"
    Else
        Response.Write "<b> ģ �� �� �� ��</b>"
    End If

    Response.Write "    </td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "        <table align='left' width='200'>"
    Response.Write "          <tr id=OpenNavigation1 ><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateStart1' onclick=""OpenNavigation(1)"">&nbsp;ʹ�ø���ı�ǩ&nbsp;</a></td></tr>"
    Response.Write "          <tr id=CloseNavigation1 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateStart1' onclick=""CloseNavigation(1)"">&nbsp;�رձ�ǩ������</a></td></tr>"
    Response.Write "        </table>"

    Call CommonLabel(1)

    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg' id=showAlgebra>"
    Response.Write "      <td>"
    Response.Write "       <table>"
    Response.Write "        <tr >"
    Response.Write "          <td width='20'><table id=showLabel style='display:none'><tr><td><iframe marginwidth=0 marginheight=0 frameborder=0  width='180' height='440' src='" & InstallDir & "editor/editor_tree.asp?ChannelID=" & ChannelID & "&ModuleType=" & ModuleType & "&insertTemplate=1&insertTemplateType=1'></iframe></td></tr></table></td>"
    Response.Write "          <td >"
    Response.Write "           <table width='100%'><tr><td width='20'>" & vbCrLf
    Response.Write "              <textarea id='txt_ln' name='rollContent'  COLS='5' ROWS='31' class='RomNumber' readonly>" & strContenttemp & "</textarea>" & vbCrLf
    Response.Write "            </td><td width='700'>"
    Response.Write "             <textarea name='Content' id='txt_main'  ROWS='30' COLS='117'  class='txt_main' wrap='OFF'  onkeydown='editTab()' onscroll=""show_ln('txt_ln','txt_main')""  wrap='on' onMouseUp=""setContent('get',1);setContent2(1)"">" & Server.HTMLEncode(strHead) & "</textarea></td></tr>"
    Response.Write "             <script>for(var i=3000; i<=3000; i++) document.getElementById('txt_ln').value += i + '\n';</script>" & vbCrLf
    Response.Write "            </td></tr>"
    Response.Write "           </table>"
    Response.Write "          </td>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "        <table align='left' width='200'>"
    Response.Write "          <tr id=OpenNavigation3 ><td>&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateEnd1'  onclick=""OpenNavigation(1)"">&nbsp;ʹ�ø���ı�ǩ&nbsp;</a></td></tr>"
    Response.Write "          <tr id=CloseNavigation3 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateEnd1' onclick=""CloseNavigation(1)"">&nbsp;�رձ�ǩ������</a></td></tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg' >"
    Response.Write "    <td><table><tr>"
    Response.Write "       <td width='95%'>&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "         <input name='EditorAlgebra' type='button' id='EditorAlgebra' value=' ����ģʽ '  onclick='LoadEditorAlgebra(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='EditorMix' type='button' id='EditorMix' value=' ���ģʽ '  disabled onclick='LoadEditorMix(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='EditorEdit' type='button' id='EditorEdit' value=' �༭ģʽ ' disabled onclick='LoadEditorEdit(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='Copy' type='button' id='Copy' value=' ���ƴ��� ' onclick='copy(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='Editorfullscreen' type='button' id='Editorfullscreen' value=' ȫ���༭ ' onclick='fullscreen(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='EditorSkin' type='button' id='EditorSkin' value=' �޸ķ�� ' onClick=""return Templateskin()"" onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "       </td>"
    Response.Write "       <td align='right' width='5%'><img  src='../Editor/images/sizeplus.gif' width='20' height='20' onclick=""sizeContent(5,'Content');sizeContent(5,'rollContent')"">&nbsp<img  src='../Editor/images/sizeminus.gif' width='20' height='20' onclick=""sizeContent(-5,'Content');sizeContent(-5,'rollContent')"">&nbsp;&nbsp;</td></tr>"
    Response.Write "     </tr></table>"
    Response.Write "    </td></tr>"
    Response.Write "    <a name='#TemplateEnd1'></a>"
    Response.Write "    <tr class='tdbg' id=showeditor style='display:none'>"
    Response.Write "      <td valign='top'>"
    Response.Write "       <table >"
    Response.Write "        <tr><td width='20'><td>"
    Response.Write "       <textarea name='EditorContent' style='display:none' >" & Server.HTMLEncode(Content) & "</textarea>"
    Response.Write "       <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=1&tContentid=EditorContent&TemplateType=1' frameborder='1' scrolling='no' width='790' height='400' ></iframe>"
    Response.Write "       </td></tr></table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    If TemplateType = 2 Then
        Response.Write "    <a name='#TemplateStart2'></a>"
        Response.Write "    <tr class='tdbg' >"
        Response.Write "      <td valign='top' align='center'>"
        Response.Write "     <b>С��ģ�壺</b>����Ŀû������Ŀʱ���ͻ���ô˴�������ʾ</td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td align='center'  align='left' valign='top'>"
        Response.Write "        <table align='left' width='200'  id='Navigation12' style='display:'>"
        Response.Write "          <tr id=OpenNavigation2 ><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateStart2' onclick=""OpenNavigation(2)"">&nbsp;ʹ�ø���ı�ǩ&nbsp;</a></td></tr>"
        Response.Write "          <tr id=CloseNavigation2 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateStart2' onclick=""CloseNavigation(2)"">&nbsp;�رձ�ǩ������</a></td></tr>"
        Response.Write "        </table>"

        Call CommonLabel(2)

        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg' id=showAlgebra2>"
        Response.Write "      <td>"
        Response.Write "       <table>"
        Response.Write "        <tr  >"
        Response.Write "          <td width='20'><table id=showLabel2 style='display:none'><tr><td><iframe marginwidth=0 marginheight=0 frameborder=0 width='180' height='440' src='" & InstallDir & "editor/editor_tree.asp?ChannelID=" & ChannelID & "&ModuleType=" & ModuleType & "&insertTemplate=1&insertTemplateType=2'></iframe></td></tr></table></td>"
        Response.Write "          <td >"
        Response.Write "           <table width='100%'><tr><td width='20'>" & vbCrLf
        Response.Write "           <textarea id='txt_ln2' name='rollContent2'  COLS='5' ROWS='31' class=RomNumber readonly>" & strContenttemp & "</textarea>" & vbCrLf
        Response.Write "            </td><td width='700'>"
        Response.Write "           <textarea name='Content2' id='txt_main2'  ROWS='30' COLS='117' wrap='OFF' id='TemplateContent2' class='txt_main'  onkeydown='editTab()' onscroll=""show_ln('txt_ln2','txt_main2')"" onMouseUp=""setContent('get',2);setContent2(2)"">" & Server.HTMLEncode(strHead) & "</textarea></td></tr>"
        Response.Write "           <script>for(var i=3000; i<=3000; i++) document.getElementById('txt_ln2').value += i + '\n';</script>" & vbCrLf
        Response.Write "            </td></tr>"
        Response.Write "           </table>"
        Response.Write "          </td>"
        Response.Write "       </table>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg' >"
        Response.Write "    <td><table><tr>"
        Response.Write "       <td width='95%'>&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "         <input name='EditorAlgebra' type='button' id='EditorAlgebra2' value=' ����ģʽ '  onclick='LoadEditorAlgebra(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='EditorMix2' type='button' id='EditorMix2' value=' ���ģʽ ' disabled onclick='LoadEditorMix(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='EditorEdit2' type='button' id='EditorEdit2' value=' �༭ģʽ '  disabled onclick='LoadEditorEdit(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='Copy2' type='button' id='Copy2' value=' ���ƴ��� '  onclick='copy(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='Editorfullscreen2' type='button' id='Editorfullscreen2' value=' ȫ���༭ ' onclick='fullscreen(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='EditorSkin2' type='button' id='EditorSkin' value=' �޸ķ�� ' onClick=""return Templateskin()"" onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "       </td>"
        Response.Write "       <td align='right' width='5%'><img  src='../Editor/images/sizeplus.gif' width='20' height='20' onclick=""sizeContent(5,'Content')"">&nbsp<img  src='../Editor/images/sizeminus.gif' width='20' height='20' onclick=""sizeContent(-5,'Content')"">&nbsp;&nbsp;</td></tr>"
        Response.Write "     </tr></table>"
        Response.Write "    <a name='#TemplateEnd2'></a>"
        Response.Write "    </td></tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td>"
        Response.Write "        <table align='left' width='200'>"
        Response.Write "          <tr id=OpenNavigation4 ><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateEnd2' onclick=""OpenNavigation(2)"">&nbsp;ʹ�ø���ı�ǩ&nbsp;</a></td></tr>"
        Response.Write "          <tr id=CloseNavigation4 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateEnd2' onclick=""CloseNavigation(2)"">&nbsp;�رձ�ǩ������</a></td></tr>"
        Response.Write "        </table>"
        Response.Write "      </td>"
        Response.Write "    </tr>"

        Response.Write "    <tr class='tdbg' id=showeditor2 style='display:none'>"
        Response.Write "      <td valign='top'>"
        Response.Write "       <table >"
        Response.Write "        <tr><td width='20'><td>"
        Response.Write "       <textarea name='EditorContent2' style='display:none' >" & Server.HTMLEncode(Content) & "</textarea>"
        Response.Write "       <iframe ID='editor2' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=1&tContentid=EditorContent2&TemplateType=2' frameborder='1' scrolling='no' width='790' height='400' ></iframe>"
        Response.Write "       </td></tr></table>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
    End If

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td valign='top'>&nbsp;&nbsp;<input name='IsDefault' type='checkbox' id='IsDefault' value='Yes'> ����ģ����Ϊ"
    Set rsTemplateProject = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")
    If rsTemplateProject.BOF And rsTemplateProject.EOF Then
        Call WriteErrMsg("<li>ϵͳ�л�û��Ĭ�Ϸ���,�뵽��������ָ��Ĭ�Ϸ�����</li>", ComeUrl)
        Exit Sub
    Else
        If ProjectName = rsTemplateProject("TemplateProjectName") Then
            Response.Write "ϵͳ"
        Else
            Response.Write "����"
        End If
    End If
    Set rsTemplateProject = Nothing
    Response.Write "Ĭ��ģ��</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='50'  align='center'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"

    If TempType > 0 Then Response.Write "<input name='TempType' type='hidden' id='TempType' value='" & TempType & "'>"
    Response.Write "       <input name='Action' type='hidden' id='Action' value='SaveAdd'><input type='button' name='button' value=' �� �� ' onClick='return CheckForm(" & Num & ");'>"
    Response.Write "     </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write " </td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    
End Sub

'=================================================
'��������Modify
'��  �ã��޸�ģ��
'=================================================
Sub Modify()
    
    Dim TemplateID, TemplateContent, TemplateContent2
    Dim arrContent, rs, sql, Num, Content, Content2
    Dim strTemp
    Dim rsTemplateProject
    '��ȡģ��ID
    TemplateID = PE_CLng(Trim(Request("TemplateID")))
 
    If TemplateID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��TemplateID</li>"
        Exit Sub
    End If

    '�õ�ģ������
    sql = "select * from PE_Template where TemplateID=" & TemplateID
    
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ����ģ�壡</li>"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    TemplateContent = rs("TemplateContent")
    
    If rs("TemplateType") = 2 Then
        '��ǰ̨jsΪ ��Ŀ����
        Num = 2
        arrContent = Split(TemplateContent, "{$$$}")
        TemplateContent = arrContent(0)
        TemplateContent2 = arrContent(1)
        Content = ShiftCharacter(TemplateContent)
        Content2 = ShiftCharacter(TemplateContent2)
    Else
        Num = 1
        Content = ShiftCharacter(TemplateContent)
    End If

    '��4.03 ���� �滻����ҳ��ʾ���ÿ�����
    If rs("TemplateType") = 3 Then
        regEx.Pattern = "(\<noscript)([\s\S]*?)(\<\/noscript\>)"
        Set Matches = regEx.Execute(Content)
        For Each Match In Matches
            strTemp = Match.value
            Content = Replace(Content, strTemp, "")
        Next
    End If

    'Ԥд3000��
    Dim strContenttemp, i
    For i = 1 To 3000
        If strContenttemp = "" Then
            strContenttemp = i & vbCrLf
        Else
            strContenttemp = strContenttemp & i & vbCrLf
        End If
    Next

    '����ǰ̨js
    Call StrJS_Template
    
    Response.Write "<form name='form1' method='post' action='Admin_Template.asp' >"
    Response.Write "  <table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22'  align='center'><strong> �� �� ģ �� �� ��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>&nbsp;&nbsp;<strong> ѡ�񷽰��� </strong><select name='ProjectName' id='ProjectName' disabled>" & GetProject_Option(ProjectName) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>&nbsp;&nbsp;<strong> ģ�����ͣ� </strong><select name='TemplateType' disabled>" & GetTemplate_Option(PE_CLng(rs("TemplateType"))) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>&nbsp;&nbsp;<strong> ģ�����ƣ� </strong><input name='TemplateName' type='text' id='TemplateName' value='" & rs("TemplateName") & "' size='50' maxlength='50'></td>"
    Response.Write "    </tr>"
    Response.Write "    <a name='#TemplateStart1'></a>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align=center>"
    If rs("TemplateType") = 2 Then
        Response.Write "<b>����ģ�壺</b>����Ŀ��������Ŀʱ���ͻ���ô˴�������ʾ��"
    Else
        Response.Write "<b> ģ �� �� �� ��</b>"
    End If
    Response.Write "    </td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td valign='top'align=center id='Navigation1' style='display:'>"
    Response.Write "        <table align='left' width='200'>"
    Response.Write "          <tr id=OpenNavigation1 ><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateStart1' onclick=""OpenNavigation(1)"">&nbsp;ʹ�ø���ı�ǩ&nbsp;</a></td></tr>"
    Response.Write "          <tr id=CloseNavigation1 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateStart1' onclick=""CloseNavigation(1)"">&nbsp;�رձ�ǩ������</a></td></tr>"
    Response.Write "        </table>"

    Call CommonLabel(1)

    Response.Write "    </td></tr>"
    Response.Write "    <tr class='tdbg' id=showAlgebra>"
    Response.Write "      <td>"
    Response.Write "       <table>"
    Response.Write "        <tr >"
    Response.Write "          <td width='20'><table id=showLabel style='display:none'><tr><td><iframe marginwidth=0 marginheight=0 frameborder=0  width='180' height='440' src='" & InstallDir & "editor/editor_tree.asp?ChannelID=" & ChannelID & "&ModuleType=" & ModuleType & "&insertTemplate=1&insertTemplateType=1'></iframe></td></tr></table></td>"
    Response.Write "          <td>"
    Response.Write "           <table width='100%'><tr><td width='20'>" & vbCrLf
    Response.Write "            <textarea id='txt_ln' name='rollContent'  COLS='5' ROWS='31'   class=RomNumber readonly>" & strContenttemp & "</textarea>" & vbCrLf
    Response.Write "            </td><td width='700'>"
    Response.Write "            <textarea name='Content' id='txt_main'  ROWS='30' COLS='117'  wrap='OFF'  onkeydown='editTab()' onscroll=""show_ln('txt_ln','txt_main')"" wrap='on' onMouseUp=""setContent('get',1);setContent2(1)"" class='txt_main'>" & Server.HTMLEncode(TemplateContent) & "</textarea></td></tr>"
    Response.Write "            <script>for(var i=3000; i<=3000; i++) document.getElementById('txt_ln').value += i + '\n';</script>" & vbCrLf
    Response.Write "            </td></tr>"
    Response.Write "           </table>"
    Response.Write "         </td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "    <td><table><tr>"
    Response.Write "       <td width='95%'>&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "         <input name='EditorAlgebra' type='button' id='EditorAlgebra' value=' ����ģʽ ' onclick='LoadEditorAlgebra(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='EditorMix' type='button' id='EditorMix' value=' ���ģʽ ' disabled onclick='LoadEditorMix(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='EditorEdit' type='button' id='EditorEdit' value=' �༭ģʽ ' disabled onclick='LoadEditorEdit(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='Copy' type='button' id='Copy' value=' ���ƴ��� ' onclick='copy(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='Editorfullscreen' type='button' id='Editorfullscreen' value=' ȫ���༭ ' onclick='fullscreen(1);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "         &nbsp;<input name='EditorSkin' type='button' id='EditorSkin' value=' �޸ķ�� ' onClick=""return Templateskin()""  onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
    Response.Write "       </td>"
    Response.Write "       <td align='right' width='5%'><img  src='../Editor/images/sizeplus.gif' width='20' height='20' onclick=""sizeContent(5,'Content');sizeContent(5,'rollContent')"">&nbsp<img  src='../Editor/images/sizeminus.gif' width='20' height='20' onclick=""sizeContent(-5,'Content');sizeContent(-5,'rollContent')"">&nbsp;&nbsp;</td></tr>"
    Response.Write "     </tr></table>"
    Response.Write "    </td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td>"
    Response.Write "        <table align='left' width='200'>"
    Response.Write "          <tr id=OpenNavigation3 ><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateEnd1' onclick=""OpenNavigation(1)"">&nbsp;ʹ�ø���ı�ǩ&nbsp;</a></td></tr>"
    Response.Write "          <tr id=CloseNavigation3 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateEnd1' onclick=""CloseNavigation(1)"">&nbsp;�رձ�ǩ������</a></td></tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg' id=showeditor style='display:none'>"
    Response.Write "      <td valign='top' >"
    Response.Write "       <table >"
    Response.Write "        <tr><td width='20'><td>"
    Response.Write "       <textarea name='EditorContent' style='display:none' >" & Server.HTMLEncode(Content) & "</textarea>"
    Response.Write "       <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=1&tContentid=EditorContent&TemplateType=1' frameborder='1' scrolling='no' width='790' height='400' ></iframe>"
    Response.Write "       </td></tr></table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <a name='#TemplateEnd1'></a>"
    '����С��ģ��ʱ
    If rs("TemplateType") = 2 Then
        Response.Write "    <a name='#TemplateStart2'></a>"
        Response.Write "<tr class='tdbg'>"
        Response.Write "   <td align='center'  align='left' valign='top'>"
        Response.Write "    <b>С��ģ�壺</b>����Ŀû������Ŀʱ���ͻ���ô˴�������ʾ</td>"
        Response.Write "   </td>"
        Response.Write "</tr>"
        Response.Write "<tr class='tdbg'>"
        Response.Write "   <td align='center'  align='left' valign='top'>"
        Response.Write "    <table align='left' width='200' id='Navigation12' style='display:'>"
        Response.Write "      <tr id=OpenNavigation2 ><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateStart2' onclick=""OpenNavigation(2)"">&nbsp;ʹ�ø���ı�ǩ&nbsp;</a></td></tr>"
        Response.Write "      <tr id=CloseNavigation2 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateStart2' onclick=""CloseNavigation(2)"">&nbsp;�رձ�ǩ������</a></td></tr>"
        Response.Write "    </table>"

        Call CommonLabel(2)

        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "   <tr class='tdbg' id=showAlgebra2>"
        Response.Write "      <td>"
        Response.Write "       <table>"
        Response.Write "        <tr >"
        Response.Write "          <td width='20'><table id=showLabel2 style='display:none'><tr><td><iframe marginwidth=0 marginheight=0 frameborder=0  width='180' height='440' src='" & InstallDir & "editor/editor_tree.asp?ChannelID=" & ChannelID & "&ModuleType=" & ModuleType & "&insertTemplate=1&insertTemplateType=2'></iframe></td></tr></table></td>"
        Response.Write "          <td>"
        Response.Write "           <table width='100%'><tr><td width='20'>" & vbCrLf
        Response.Write "            <textarea id='txt_ln2' name='rollContent2'  COLS='5' ROWS='31'   class=RomNumber readonly>" & strContenttemp & "</textarea>" & vbCrLf
        Response.Write "            </td><td width='700'>"
        Response.Write "            <textarea name='Content2' id='txt_main2'  ROWS='30' COLS='117' wrap='OFF' id='TemplateContent2' class='txt_main' onkeydown=""editTab()"" onscroll=""show_ln('txt_ln2','txt_main2')"" onMouseUp=""setContent('get',2);setContent2(2)"">" & Server.HTMLEncode(TemplateContent2) & "</textarea></td></tr>"
        Response.Write "            <script>for(var  i=3000; i<=3000; i++) document.getElementById('txt_ln2').value += i + '\n';</script>" & vbCrLf
        Response.Write "            </td></tr>"
        Response.Write "           </table>"
        Response.Write "       </table>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td>"
        Response.Write "        <table align='left' width='200'>"
        Response.Write "          <tr id=OpenNavigation4 ><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_open.gif' BORDER='0'   ALT=''><a href='#TemplateEnd2' onclick=""OpenNavigation(2)"">&nbsp;ʹ�ø���ı�ǩ&nbsp;</a></td></tr>"
        Response.Write "          <tr id=CloseNavigation4 style='display:none'><td >&nbsp;&nbsp;&nbsp;<IMG SRC='Images/admin_close.gif' BORDER='0'   ALT=''><a href='#TemplateEnd2' onclick=""CloseNavigation(2)"">&nbsp;�رձ�ǩ������</a></td></tr>"
        Response.Write "        </table>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg' >"
        Response.Write "    <td><table><tr>"
        Response.Write "       <td width='95%'>&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "         <input name='EditorAlgebra2' type='button' id='EditorAlgebra2' value=' ����ģʽ ' onclick='LoadEditorAlgebra(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='EditorMix2' type='button' id='EditorMix2' value=' ���ģʽ ' disabled onclick='LoadEditorMix(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='EditorEdit2' type='button' id='EditorEdit2' value=' �༭ģʽ ' disabled onclick='LoadEditorEdit(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='Copy2' type='button' id='Copy2' value=' ���ƴ��� ' onclick='copy(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='Editorfullscreen2' type='button' id='Editorfullscreen2' value=' ȫ���༭ ' onclick='fullscreen(2);' onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "         &nbsp;<input name='EditorSkin' type='button' id='EditorSkin' value=' �޸ķ�� ' onClick=""return Templateskin()"" onmouseover=""this.style.backgroundColor='#BFDFFF'"" onmouseout=""this.style.backgroundColor=''"">"
        Response.Write "       </td>"
        Response.Write "       <td align='right' width='5%'><img  src='../Editor/images/sizeplus.gif' width='20' height='20' onclick=""sizeContent(5,'Content2');sizeContent(5,'rollContent2')"">&nbsp<img  src='../Editor/images/sizeminus.gif' width='20' height='20' onclick=""sizeContent(-5,'Content2');sizeContent(-5,'rollContent2')"">&nbsp;&nbsp;</td></tr>"
        Response.Write "     </tr></table>"
        Response.Write "    </td></tr>"
        Response.Write "  <tr class='tdbg'id=showeditor2 style='display:none'>"
        Response.Write "   <td valign='top' >"
        Response.Write "     <table >"
        Response.Write "      <tr><td width='20'><td>"
        Response.Write "       <textarea name='EditorContent2' style='display:none' >" & Server.HTMLEncode(Content2) & "</textarea>"
        Response.Write "       <iframe ID='editor2' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=1&tContentid=EditorContent2&TemplateType=2' frameborder='1' scrolling='no' width='790' height='400' ></iframe>"
        Response.Write "       </td>"
        Response.Write "      </tr>"
        Response.Write "     </table>"
        Response.Write "   </td>"
        Response.Write "</tr>"
    
    End If

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td valign='top'>&nbsp;&nbsp;<input name='IsDefault' type='checkbox' id='IsDefault' value='Yes'"

    If rs("IsDefault") = True Then Response.Write " checked"
    Response.Write "> ����ģ����Ϊ"

    Set rsTemplateProject = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")
    If rsTemplateProject.BOF And rsTemplateProject.EOF Then
        Call WriteErrMsg("<li>ϵͳ�л�û��Ĭ�Ϸ���,�뵽��������ָ��Ĭ�Ϸ�����</li>", ComeUrl)
        Exit Sub
    Else
        If ProjectName = rsTemplateProject("TemplateProjectName") Then
            Response.Write "ϵͳ"
        Else
            Response.Write "����"
        End If
    End If
    Set rsTemplateProject = Nothing

    Response.Write "Ĭ��ģ��</td>"
    Response.Write "    </tr>"
    Response.Write "    <a name='#TemplateEnd2'></a>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td height='50'  align='center'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'><input name='TemplateID' type='hidden' id='TemplateID' value='" & TemplateID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModify'><input name='TemplateType' type='hidden' id='Action' value='" & rs("TemplateType") & "'>"

    If TempType > 0 Then Response.Write "<input name='TempType' type='hidden' id='TempType' value='" & TempType & "'>"
    Response.Write "        <input type='button' name='Submit2' value=' �����޸Ľ�� ' onClick='return CheckForm(" & Num & ");'></td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"

    rs.Close
    Set rs = Nothing
End Sub

'=================================================
'��������Save
'��  �ã�����ģ��
'=================================================
Sub Save()
    
    Dim rs, sql, Action
    Dim TemplateID, ProjectName, TemplateName, IsDefault, IsDefaultInProject
    Dim DefaultType, setUpdateItem
    Dim TemplateContent, TemplateContent2, i
    
    '�õ�ģ��ID ���� ����
    TemplateID = Trim(Request.Form("TemplateID"))
    TemplateName = Trim(Request.Form("TemplateName"))
    Action = Trim(Request.Form("Action"))
    TemplateType = Trim(Request.Form("TemplateType"))
    ProjectName = Trim(Request.Form("ProjectName"))
    '������
    If Action = "SaveModify" Then
        If TemplateID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��TemplateID</li>"
        Else
            TemplateID = PE_CLng(TemplateID)
        End If
        Set rs = Conn.Execute("Select TemplateID,ProjectName From PE_Template Where TemplateID=" & TemplateID & "")
        If rs.BOF And rs.EOF Then
            Call WriteErrMsg("<li>ϵͳ�л�û��ģ�壡</li>", ComeUrl)
            Exit Sub
        Else
            ProjectName = rs("ProjectName")
        End If
        Set rs = Nothing
    End If
    
    If TemplateName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ģ�����Ʋ���Ϊ�գ�</li>"
    End If
         
    For i = 1 To Request.Form("Content").Count
        TemplateContent = TemplateContent & Request.Form("Content")(i)
    Next
    
    For i = 1 To Request.Form("Content2").Count
        TemplateContent2 = TemplateContent2 & Request.Form("Content2")(i)
    Next
    
    If TemplateType <> 2 Then
        TemplateContent = ShiftCharacterSave(TemplateContent)
    Else
        TemplateContent = ShiftCharacterSave(TemplateContent)
        TemplateContent2 = ShiftCharacterSave(TemplateContent2)
    End If
    
    If Len(TemplateName) > 50 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ģ����ⲻ�ܳ���50���ַ��� </li>"
    End If
        
    If InStr(TemplateContent, "rsClass_") > 0 And Not ((TemplateType = 1 And ChannelID <> 0) Or TemplateType = 2 Or TemplateType = 101) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ƶ����ҳ����Ŀ����ģ���⣬����ģ���еı�ǩ��������ʹ�� rsClass_��ͷ�ı�ǩ������</li>"
    End If

    If InStr(TemplateContent2, "rsClass_") > 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ĿС��ģ���в������ʹ�� rsClass_��ͷ�ı�ǩ������ </li>"
    End If

    If TemplateType = 101 Then
        If InStr(TemplateContent, "��/ArticleList��") > 0 Or InStr(TemplateContent, "��/ProductList��") Or InStr(TemplateContent, "��/SoftList��") Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�Զ����б�ģ�岻�ܰ����Լ����͵�ģ�壡 </li>"
        End If
    End If

    If UBound(Split(TemplateContent, "<!--")) > UBound(Split(TemplateContent, "-->")) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����¼��ģ�� &lt;!-- �������� --&gt; ����������ע�ͺ���ģ�岻������ </li>"
    End If

    If TemplateType = 2 Then
        If UBound(Split(TemplateContent2, "<!--")) > UBound(Split(TemplateContent2, "-->")) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����С��ģ������ &lt;!-- �������� --&gt; ����������ע�ͺ���ģ�岻������ </li>"
        End If
    End If
            
    If FoundErr = True Then Exit Sub
    '��� �Ƿ� ����ΪĬ��
    IsDefault = Trim(Request("IsDefault"))

    '�ж��Ƿ�Ĭ��
    If IsDefault = "Yes" Then
        IsDefault = True
    Else
        IsDefault = False
    End If

    'ִ��Ĭ��ѡ��
    If IsDefault = True Then
        '----------------------------------------------------
        '�ж��Ƿ�ϵͳĬ�Ϸ���
        Set rs = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")
        If rs.BOF And rs.EOF Then
            Call WriteErrMsg("<li>ϵͳ�л�û��Ĭ�Ϸ���,�뵽��������ָ��Ĭ�Ϸ�����</li>", ComeUrl)
            Exit Sub
        Else
            If ProjectName = rs("TemplateProjectName") Then
                DefaultType = 1
            Else
                DefaultType = 2
            End If
        End If
        Set rs = Nothing

        If DefaultType = 1 Then
            setUpdateItem = "IsDefault=" & PE_False & ",IsDefaultInProject=" & PE_False
        ElseIf DefaultType = 2 Then
            setUpdateItem = "IsDefaultInProject=" & PE_False
        End If
        Conn.Execute ("update PE_Template set " & setUpdateItem & " where ChannelID=" & ChannelID & " and TemplateType=" & TemplateType & " and ProjectName='" & ProjectName & "'")
    End If
    '----------------------------------------------------
    '��ӱ���
    If Action = "SaveAdd" Then
        sql = "select top 1 * from PE_Template"
        Set rs = Server.CreateObject("Adodb.RecordSet")
        rs.Open sql, Conn, 1, 3
        rs.addnew
        rs("ChannelID") = ChannelID
        rs("Templatetype") = TemplateType
        rs("ProjectName") = ProjectName
        rs("TemplateName") = TemplateName
        
        '��С��ģ����ж�
        If TemplateType = 2 Then
            rs("TemplateContent") = TemplateContent & "{$$$}" & TemplateContent2
        Else
            rs("TemplateContent") = TemplateContent
        End If
        If IsDefault = True Then
            If DefaultType = 1 Then
                rs("IsDefault") = True
                rs("IsDefaultInProject") = True
            Else
                rs("IsDefault") = False
                rs("IsDefaultInProject") = True
            End If
        Else
            rs("IsDefault") = False
            rs("IsDefaultInProject") = False
        End If
        rs.Update
        rs.Close
        Set rs = Nothing
        Call WriteSuccessMsg("�ɹ�����µ�ģ�壺" & Trim(Request("TemplateName")), ComeUrl)
    Else
        '�޸ı���
        sql = "select * from PE_Template where TemplateID=" & TemplateID
        Set rs = Server.CreateObject("Adodb.RecordSet")
        rs.Open sql, Conn, 1, 3

        If rs.BOF And rs.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�Ҳ���ָ���İ������ģ�壡</li>"
        Else

            If rs("TemplateType") = 2 Then
                If TemplateContent2 = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>С��ģ�����ݲ���Ϊ�գ�</li>"
                    rs.Close
                    Set rs = Nothing
                    Exit Sub
                End If

                rs("TemplateContent") = TemplateContent & "{$$$}" & TemplateContent2
            Else
                rs("TemplateContent") = TemplateContent
            End If

            rs("TemplateName") = TemplateName

            If IsDefault = True Then
                If DefaultType = 1 Then
                    rs("IsDefault") = True
                    rs("IsDefaultInProject") = True
                Else
                    rs("IsDefault") = False
                    rs("IsDefaultInProject") = True
                End If
            End If
            rs.Update
            Call WriteSuccessMsg("����ģ��ɹ���", ComeUrl)
        End If

        rs.Close
        Set rs = Nothing

        If IsDefault = True And TemplateType = 1 Then
            If ChannelID = 0 Then
                
                Dim FileExt_SiteIndex, FileName_Index
                FileExt_SiteIndex = arrFileExt(FileExt_SiteIndex)
                FileName_Index = "Index" & FileExt_SiteIndex
                
                If FileName_Index = "Index.asp" Then
                    ErrMsg = ErrMsg & "<li>��Ϊ��վ������δ������վ��ҳ����HTML���ܣ����Բ���������ҳ��</li>"
                    Response.Write ErrMsg
                    Exit Sub
                End If
                
                If ObjInstalled_FSO = True Then
                    Response.Write "<br><iframe  width='100%' height='210' frameborder='0' src='Admin_CreateSiteIndex.asp'></iframe>"
                Else
                    ErrMsg = ErrMsg & "<li>��Ϊ��վ��֧��FSO �� ����FSO�Ѹ�����</li>"
                    Response.Write ErrMsg
                    Exit Sub
                End If

            Else

                If UseCreateHTML > 0 And ObjInstalled_FSO = True Then
                    Response.Write "<br><iframe  width='100%' height='210' frameborder='0' src='Admin_Create" & ModuleName & ".asp?ChannelID=" & ChannelID & "&CreateType=1&Action=CreateIndex&ShowBack=No'></iframe>"
                End If
            End If
        End If
    End If

    Call ClearSiteCache(0)
End Sub

'=================================================
'��������SetDefault
'��  �ã�����ָ����Ĭ��ģ��
'=================================================
Sub SetDefault()
    Dim TemplateID, DefaultType, setUpdateItem, setUpdateItem2, strTemp, ProjectName
    TemplateID = PE_CLng(Trim(Request("TemplateID")))
    DefaultType = PE_CLng(Trim(Request("DefaultType")))
    ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))

    If TemplateID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��TemplateID</li>"
        Exit Sub
    End If
        
    If DefaultType = 1 Then
        setUpdateItem = "IsDefault=" & PE_False & ",IsDefaultInProject=" & PE_False
        setUpdateItem2 = "IsDefault=" & PE_True & ",IsDefaultInProject=" & PE_True
        strTemp = "<li>�ɹ���ѡ����ģ��,����Ϊ<FONT style='font-size:12px' color='#008000'>ϵͳĬ��</FONT>ģ��.</li>"
    ElseIf DefaultType = 2 Then
        setUpdateItem = "IsDefaultInProject=" & PE_False
        setUpdateItem2 = "IsDefaultInProject=" & PE_True
        strTemp = "<li>�ɹ���ѡ����ģ��,����Ϊ<FONT style='font-size:12px' color='#3366FF'>����Ĭ��</FONT>ģ��.</li>"
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�趨��Ĭ�����Ͳ���!</li>"
        Exit Sub
    End If

    Conn.Execute ("update PE_Template set " & setUpdateItem & " where ChannelID=" & ChannelID & " and TemplateType=" & TemplateType & " and ProjectName='" & ProjectName & "'")
    Conn.Execute ("update PE_Template set " & setUpdateItem2 & " where ChannelID=" & ChannelID & " and TemplateType=" & TemplateType & " and TemplateID=" & TemplateID)
    Call WriteSuccessMsg(strTemp, ComeUrl)
    Call ClearSiteCache(0)
End Sub

'=================================================
'��������DelTemplate
'��  �ã�ɾ��ָ��ģ��
'=================================================
Sub DelTemplate()
    Dim TemplateID, rs, trs, sql, downright

    FoundErr = False

    'downright 0 ɾ�������ݿ� 1 ѡ��ģ�峹��ɾ�� 2 ��ջ���վ 3 ѡ���Ļ�ԭ 4 ȫ����ԭ
    downright = PE_CLng(Trim(Request("downright")))
    TemplateID = Trim(Request("TemplateID"))
	If IsValidID(TemplateID) = False Then
		TemplateID = ""
	End If

    If downright = 2 Or downright = 4 Then
        sql = "select * from PE_Template where Deleted=" & PE_True
    Else

        If TemplateID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��TemplateID</li>"
            Exit Sub
        End If

        If InStr(TemplateID, ",") > 0 Then
            sql = "select * from PE_Template where TemplateID In(" & TemplateID & ")"
        Else
            sql = "select * from PE_Template where TemplateID=" & PE_CLng(TemplateID)
        End If
    End If

    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 3

    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���İ������ģ�壡</li>"
    Else

        Do While Not rs.EOF

            If downright = 1 Or downright = 2 Then
                If rs("IsDefault") = True Or rs("IsDefaultInProject") = True Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>" & rs("TemplateName") & "��ǰģ��Ϊ����Ĭ��ģ�壬����ɾ����</li><li>���Ƚ�Ĭ��ģ���Ϊ����ģ�������ɾ����ģ�塣</li>"
                Else
                    Set trs = Conn.Execute("select TemplateID from PE_Template where ChannelID=" & ChannelID & " and IsDefault=" & PE_True & " and TemplateType=" & rs("TemplateType"))

                    If trs.BOF And trs.EOF Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>" & rs("TemplateName") & "�Ҳ��������õ�Ĭ��ģ�壬���Բ���ɾ����ǰģ�塣���Ƚ�����һ��ģ���ΪĬ��ģ�������ɾ����ģ�塣</li>"
                    Else
                        Select Case rs("TemplateType")

                            Case 1
                                Conn.Execute ("update PE_Channel set Template_Index=0 where ChannelID=" & ChannelID & " and Template_Index=" & rs("TemplateID"))

                            Case 2
                                Conn.Execute ("update PE_Class set TemplateID=0 where ChannelID=" & ChannelID & " and TemplateID=" & rs("TemplateID"))

                            Case 3
                                Conn.Execute ("update PE_Article set TemplateID=0 where ChannelID=" & ChannelID & " and TemplateID=" & rs("TemplateID"))

                            Case 4
                                Conn.Execute ("update PE_Special set TemplateID=0 where ChannelID=" & ChannelID & " and TemplateID=" & rs("TemplateID"))
                        End Select

                        TemplateType = rs("TemplateType")

                        If downright = 1 Then
                            ErrMsg = ErrMsg & "<li>�ɹ�ɾ�� <font color=red>" & rs("TemplateName") & "</font>ģ�塣����ʹ�ô�ģ�����Ŀ�����¸�Ϊʹ��Ĭ��ģ�塣</li><br>"
                        End If

                        rs.Delete
                        rs.Update
                    End If

                    Set trs = Nothing
                End If

            ElseIf downright = 3 Or downright = 4 Then
                rs("Deleted") = False
                If downright = 3 Then
                    ErrMsg = "<FONT color='blue'>" & rs("TemplateName") & "</FONT>ģ���Ѿ���ԭ��"
                End If
            Else

                If rs("IsDefaultInProject") = True Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>" & rs("TemplateName") & "��ǰģ��Ϊ����Ĭ��ģ�壬����ɾ����</li><li>���Ƚ�Ĭ��ģ���Ϊ����ģ�������ɾ����ģ�塣</li>"
                Else
                    rs("Deleted") = True
                    ErrMsg = ErrMsg & "�ɹ�ɾ��<font color=red>" & rs("TemplateName") & "</font>ģ�塣��������ģ�����վ�ָ�����<br>"
                End If
            End If

            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing


    If FoundErr = False Then
        If downright = 2 Then
            ErrMsg = "<li>�ɹ��������ģ�����վ��</li>"
            
        ElseIf downright = 4 Then
            ErrMsg = "<li>�ɹ���ȫ��ģ�廹ԭ��</li>"
        End If
        Call WriteSuccessMsg(ErrMsg, ComeUrl)
    End If

End Sub

'=================================================
'��������BatchDefault
'��  �ã���������Ĭ��
'=================================================
Sub BatchDefault()
    Dim sql, rs
    Dim iTemplateType, iChannelID, i, Num
    Dim rsTemplateProject, sqlTemplateProject, IsProjectDefault
    iChannelID = 0
    iTemplateType = 0
    i = 0
    Num = 1

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td>"
    Response.Write "    | <a href='Admin_Template.asp?Action=BatchDefault&ChannelID=0&ProjectName=" & Server.UrlEncode(ProjectName) & "'><FONT style='font-size:12px' " & IsFontChecked(ChannelID, 0) & ">��վͨ��ģ��</FONT></a>"
    i = 0
    sql = "SELECT DISTINCT t.ChannelID,c.ChannelName FROM PE_Template t INNER JOIN PE_Channel c ON t.ChannelID = c.ChannelID where c.Disabled=" & PE_False
    Set rs = Conn.Execute(sql)
        
    If rs.BOF And rs.EOF Then
        IsProjectDefault = False
    Else

        Do While Not rs.EOF
            Response.Write "    | <a href='Admin_Template.asp?Action=BatchDefault&ChannelID=" & rs("ChannelID") & "&ProjectName=" & Server.UrlEncode(ProjectName) & "'><FONT style='font-size:12px' " & IsFontChecked(rs("ChannelID"), ChannelID) & ">" & rs("ChannelName") & "Ƶ��ģ��</FONT></a>"

            If i > 3 Then
                Response.Write " | </td><tr class='title'><td>"
                i = 0
            Else
                i = i + 1
            End If

            rs.MoveNext
        Loop

        Response.Write " | "
    End If

    rs.Close
    Set rs = Nothing

    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    
    '�ǲ��ǵ�ǰϵͳĬ�Ϸ���
    Set rs = Conn.Execute("select * from PE_TemplateProject Where IsDefault=" & PE_True & "")
    If rs.BOF And rs.EOF Then
        IsProjectDefault = False
    Else
        If rs("TemplateProjectName") = ProjectName Then
            IsProjectDefault = True
        Else
            IsProjectDefault = False
        End If
    End If
    Set rs = Nothing

    sql = "select * from PE_Template where Deleted=" & PE_False & " and ChannelID=" & ChannelID & " and ProjectName='" & ProjectName & "' order by TemplateType,ChannelID"
        
    Set rs = Conn.Execute(sql)

    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "    function CheckAll(thisform){" & vbCrLf
    Response.Write "        for (var i=0;i<thisform.elements.length;i++){" & vbCrLf
    Response.Write "            var e = thisform.elements[i];" & vbCrLf
    Response.Write "            if (e.Name != ""chkAll""&&e.disabled!=true&&e.zzz!=1)" & vbCrLf
    Response.Write "                e.checked = thisform.chkAll.checked;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<form name='form1' method='post' action='Admin_Template.asp'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "     <tr class='title' height='22'>"
    Response.Write "      <td width='30' align='center'><strong>ѡ��</strong></td>"
    Response.Write "      <td width='30' align='center'><strong>ID</strong></td>"
    Response.Write "      <td width='150' align='center'><b>ģ������</b></td>"
    Response.Write "      <td height='22' align='center'><strong>ģ������</strong></td>"
    Response.Write "      <td width='85' align='center'><strong>�Ƿ�"
    If IsProjectDefault = True Then
        Response.Write "ϵͳ"
    Else
        Response.Write "����"
    End If
    Response.Write "Ĭ��</strong></td>"
    Response.Write "     </tr>"
    i = 0

    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg'><td width='100%' colspan='6' align='center'> �� �� �� �� û �� ģ ��</td></tr>"
    Else

        Do While Not rs.EOF

            If i > 0 And rs("TemplateType") <> iTemplateType Or i > 0 And rs("ChannelID") <> iChannelID Then
                Num = Num + 1
                Response.Write "<tr height='10'><td colspan='6'></td></tr>"
            End If

            iChannelID = rs("ChannelID")
            iTemplateType = rs("TemplateType")
            i = i + 1

            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "  <td width=""30"" align=""center"" height=""30"">" & vbCrLf
            Response.Write "    <input TYPE='radio' value='" & rs("TemplateID") & "' name=""TemplateID" & Num & """"
            If IsProjectDefault = True Then
                If rs("IsDefault") = True Then Response.Write "checked"
            Else
                If rs("IsDefaultInProject") = True Then Response.Write "checked"
            End If
            Response.Write "> " & vbCrLf
            Response.Write "  </td>" & vbCrLf
            Response.Write "      <td width='30' align='center'>" & rs("TemplateID") & "</td>"
            Response.Write "      <td width='150' align='center'>" & GetTemplateTypeName(rs("TemplateType"), rs("ChannelID")) & "</td>"
            Response.Write "      <td align='center'><a href='Admin_Template.asp?ChannelID=" & ChannelID & "&Action=Modify&TemplateID=" & rs("TemplateID") & "'>" & rs("TemplateName") & "</a></td>"
            Response.Write "      <td width='80' align='center'><b>"

            If IsProjectDefault = True Then
                If rs("IsDefault") = True Then
                    Response.Write "��"
                Else
                    Response.Write "��"
                End If
            Else
                If rs("IsDefaultInProject") = True Then
                    Response.Write "��"
                Else
                    Response.Write "��"
                End If
            End If

            Response.Write "</td>"
            Response.Write "</tr>"

            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "    <tr class=""tdbg""> " & vbCrLf
    Response.Write "      <td colspan=6 height=""30"" align='left'>" & vbCrLf
    Response.Write "        <input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>ѡ������ģ��&nbsp;<FONT style='font-size:12px' color='blue'>��ע��ѡ������ģ������ģ�������ж��ģ�壬ϵͳ����ѡ������͵����һ��ģ�壩</FONT> &nbsp;&nbsp;"
    Response.Write "        <input name=""ProjectName"" type=""hidden""  value=" & ProjectName & ">   " & vbCrLf
    Response.Write "        <input name=""ContentNum"" type=""hidden""  value=" & Num & ">   " & vbCrLf
    Response.Write "        <input name=""Action"" type=""hidden""  value=""DoBatchDefault"">   " & vbCrLf
    Response.Write "        <input name=""ChannelID"" type=""hidden""  value=" & ChannelID & ">" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr> " & vbCrLf
    Response.Write "</table>  "
    Response.Write "<br><center><input type=""submit"" value="" �� �� �� �� �� �� Ĭ �� ""></center><br>" & vbCrLf
    Response.Write "</form>"
End Sub

'=================================================
'��������DoBatchDefault
'��  �ã��������ô���
'=================================================
Sub DoBatchDefault()
    Dim ContentNum, ProjectName, arrTemplateID, arrContent, i, DefaultType

    ContentNum = PE_CLng(Trim(Request("ContentNum")))
    ProjectName = ReplaceBadChar(Trim(Request("ProjectName")))
        
    For i = 1 To ContentNum
        arrTemplateID = arrTemplateID & PE_CLng(Trim(Request("TemplateID" & i & ""))) & ","
    Next

    '�ж��Ƿ�ϵͳĬ�Ϸ���
    Set rs = Conn.Execute("Select TemplateProjectName From PE_TemplateProject Where IsDefault=" & PE_True & "")

    If rs.BOF And rs.EOF Then
        Call WriteErrMsg("<li>ϵͳ�л�û��Ĭ�Ϸ���,�뵽��������ָ��Ĭ�Ϸ�����</li>", ComeUrl)
        Exit Sub
    Else

        If ProjectName = rs("TemplateProjectName") Then
            DefaultType = "IsDefault"
        Else
            DefaultType = "IsDefaultInProject"
        End If
    End If

    Set rs = Nothing

    arrTemplateID = Left(arrTemplateID, Len(arrTemplateID) - 1)
    If DefaultType = "IsDefaultInProject" Then
        Conn.Execute ("update PE_Template set " & DefaultType & "=" & PE_False & " where ChannelID=" & ChannelID & " and ProjectName='" & ProjectName & "'")
        Conn.Execute ("update PE_Template set " & DefaultType & "=" & PE_True & " where ChannelID=" & ChannelID & " and ProjectName='" & ProjectName & "' and TemplateID in (" & arrTemplateID & " )")
    Else
        Conn.Execute ("update PE_Template set " & DefaultType & "=" & PE_False & ",IsDefaultInProject=" & PE_False & " where ChannelID=" & ChannelID & " and ProjectName='" & ProjectName & "'")
        Conn.Execute ("update PE_Template set " & DefaultType & "=" & PE_True & ",IsDefaultInProject=" & PE_True & " where ChannelID=" & ChannelID & " and ProjectName='" & ProjectName & "' and TemplateID in (" & arrTemplateID & " )")
    End If
    Call WriteSuccessMsg("�ɹ���ѡ����ģ������ΪĬ��ģ��", ComeUrl)
    Call ClearSiteCache(0)
End Sub

'=================================================
'��������Export
'��  �ã�����ģ��
'=================================================
Sub Export()
    
    Dim rs, sql
    Dim trs, iCount, ModuleType, ProjectName
    
    '999999 Ϊ����
    ModuleType = Trim(Request.Form("ModuleType"))
    If ReplaceBadChar(Trim(Request.QueryString("ProjectName"))) = "" Then
        ProjectName = ReplaceBadChar(Trim(Request.Form("ProjectName")))
    Else
        ProjectName = ReplaceBadChar(Trim(Request.QueryString("ProjectName")))
    End If


    If ModuleType = "" Then
        ModuleType = 999999
    End If
 
    Response.Write "<form name='myform' method='post' action='Admin_Template.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>ģ�嵼��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'><td height='10'></td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='center'>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "            <select name='ProjectName' id='ProjectName' style='width:150px;'  onChange='document.myform.submit();' >"
    sql = "select * from PE_TemplateProject"
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>û���κη�����Ŀ</option>"
    Else
        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TemplateProjectName") & "' " & OptionValue(rs("TemplateProjectName"), ProjectName) & ">" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Response.Write "            </select>"
    Response.Write "            <br>"
    Response.Write "             <select name='ModuleType' onChange='document.myform.submit();'>"
    Call GetAllModule("5.0", ModuleType)
    Response.Write "             </select>"
    Response.Write "            </td><td></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "           <td>"
    Response.Write "            <select name='TemplateID' size='2' multiple style='height:300px;width:450px;'>"
    
    '�ж��Ƿ������У�999999������ҳ��0����ָ����ģ��
    If CLng(ModuleType) = 999999 Then
        sql = "select ChannelID,TemplateID,TemplateName,ProjectName from PE_Template where Deleted=" & PE_False & " And ProjectName='" & ProjectName & "'"
        If FoundInArr(AllModules, "Supply", ",") = False Then
            sql = sql & " And ChannelID <>999 "
        End If
        If FoundInArr(AllModules, "House", ",") = False Then
            sql = sql & " And ChannelID <>998 "
        End If
        If FoundInArr(AllModules, "Job", ",") = False Then
            sql = sql & " And ChannelID <>997 "
        End If
    ElseIf CLng(ModuleType) = 0 Then
        sql = "select ChannelID,TemplateID,TemplateName,ProjectName from PE_Template where ChannelID=0 And ProjectName='" & ProjectName & "'"
    Else
        sql = "select t.ChannelID,t.TemplateID,t.TemplateName,t.ProjectName,c.ModuleType from PE_Template t inner join PE_Channel c on t.ChannelID=c.ChannelID where c.ChannelType < 2 and c.Disabled=" & PE_False & " and c.ModuleType=" & PE_CLng(ModuleType) & " And t.ProjectName='" & ProjectName & "' Order by t.ChannelID asc"
    End If

    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 1

    If rs.BOF And rs.EOF Then
        Response.Write "         <option value=''>û���κ�ģ��</option>"
        '�ر��ύ��ť
        iCount = 0
    Else
        iCount = rs.RecordCount

        Do While Not rs.EOF
            Response.Write "     <option value='" & rs("TemplateID") & "'>" & rs("TemplateName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "         </select>"
    Response.Write "       </td>"
    Response.Write "       <td align='left'>&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='Submit' value=' ѡ������ ' onclick='SelectAll()'>"
    Response.Write "       <br><br>&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='Submit' value=' ȡ��ѡ�� ' onclick='UnSelectAll()'><br><br><br><b>&nbsp;��ʾ����ס��Ctrl����Shift�������Զ�ѡ</b></td>"
    Response.Write "      </tr>"
    Response.Write "      <tr height='30'>"
    Response.Write "        <td colspan='2'>Ŀ�����ݿ⣺<input name='TemplateMdb' type='text' id='TemplateMdb' value='../temp/Template.mdb' size='20' maxlength='50'>&nbsp;&nbsp;<INPUT TYPE='checkbox' NAME='FormatConn' value='yes' id='id' checked> �����Ŀ�����ݿ�</td>"
    Response.Write "      </tr>"
    Response.Write "      <tr height='50'>"
    Response.Write "         <td colspan='2' align='center'><input type='submit' name='Submit' value='ִ�е�������' onClick=""document.myform.Action.value='DoExport';"">"
    Response.Write "                  <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"

    If TempType > 0 Then Response.Write "<input name='TempType' type='hidden' id='TempType' value='" & TempType & "'>"
    Response.Write "                  <input name='Action' type='hidden' id='Action' value='Export'>"
    Response.Write "         </td>"
    Response.Write "        </tr>"
    Response.Write "    </table>"
    Response.Write "   </td>"
    Response.Write " </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.TemplateID.length;i++){" & vbCrLf
    Response.Write "    document.myform.TemplateID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.TemplateID.length;i++){" & vbCrLf
    Response.Write "    document.myform.TemplateID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

End Sub

'=================================================
'��������DoExport
'��  �ã�����ģ�崦��
'=================================================
Sub DoExport()
    On Error Resume Next
    Dim mdbname, tconn, trs, strSql, Table_PE_lable
    Dim TemplateID, rs, sql, FormatConn, rsLabel
    TemplateID = Trim(Request("TemplateID"))
    FormatConn = Request.Form("FormatConn")
    mdbname = Replace(Trim(Request.Form("Templatemdb")), "'", "")
    If IsValidID(TemplateID) = False Then
        TemplateID = ""
    End If

    If TemplateID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ������ģ��</li>"
    End If

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д����ģ�����ݿ���"
    End If
        
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        FoundErr = True
        Set tconn = Nothing
        ErrMsg = ErrMsg & "<li>���ݿ����ʧ�ܣ����Ժ����ԣ�����ԭ��" & Err.Description
        Err.Clear
    End If

    If FoundErr = True Then
        Exit Sub
    End If

    Table_PE_lable = True
    tconn.Execute ("select LabelName from PE_Label")

    If Err Then
        Table_PE_lable = False
    End If

    '�ж�PE_Label ���Ƿ����
    If Table_PE_lable = False Then
        strSql = "        CREATE TABLE PE_Label  ("
        strSql = strSql & "  LabelID counter PRIMARY KEY,"
        strSql = strSql & "  LabelName text(50),"
        strSql = strSql & "  LabelClass text(50),"
        strSql = strSql & "  PageNum int,"
        strSql = strSql & "  LabelType int,"
        strSql = strSql & "  reFlashTime int,"
        strSql = strSql & "  fieldlist text(50),"
        strSql = strSql & "  LabelIntro text(255),"
        strSql = strSql & "  Priority int,"
        strSql = strSql & "  LabelContent Memo,"
        strSql = strSql & "  AreaCollectionID int"
        strSql = strSql & " )"
        Set trs = tconn.Execute(strSql)
        Set trs = Nothing
    End If
      
    If FormatConn <> "" Then
        tconn.Execute ("delete from PE_Template")
        tconn.Execute ("delete from PE_Label")
    End If

    Set rs = Conn.Execute("select t.ChannelID,t.TemplateID,t.TemplateName,t.TemplateType,t.TemplateContent,t.IsDefault,t.ProjectName,c.ModuleType from PE_Template t left join PE_Channel c on t.ChannelID=c.ChannelID where t.TemplateID in (" & TemplateID & ")  order by t.TemplateID")
 
    Dim i, iVersion
    iVersion = 4
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_Template", tconn, 1, 3

    For i = 0 To trs.Fields.Count - 1
        If LCase(trs.Fields(i).name) = "moduletype" Then
            iVersion = 5
            Exit For
        End If
    Next

    If iVersion = 4 Then
        trs.Close
        tconn.Execute ("alter table [PE_Template]  add COLUMN ModuleType int null")
        trs.Open "select * from PE_Template", tconn, 1, 3
    End If

    Do While Not rs.EOF
        trs.addnew
        trs("TemplateID") = rs("TemplateID")
        trs("ChannelID") = rs("ChannelID")

        If rs("ModuleType") <> "" And Not IsNull(rs("ModuleType")) Then
            trs("ModuleType") = rs("ModuleType")
        Else
            trs("ModuleType") = 0
        End If

        trs("TemplateName") = rs("TemplateName")
        trs("TemplateType") = rs("TemplateType")
        trs("TemplateContent") = rs("TemplateContent")
        trs("IsDefault") = rs("IsDefault")
        trs.Update
        rs.MoveNext
    Loop

    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    
    '���ǵ�����ǩ
    Set trs = Conn.Execute("select * from PE_Label")
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open "select * from PE_Label", tconn, 1, 3
    
    If Not trs.EOF Then
        Do While Not trs.EOF
            rs.addnew
            rs("LabelName") = trs("LabelName")
            rs("LabelClass") = trs("LabelClass")
            rs("LabelType") = trs("LabelType")
            rs("PageNum") = trs("PageNum")
            rs("reFlashTime") = trs("reFlashTime")
            rs("fieldlist") = trs("fieldlist")
            rs("LabelIntro") = trs("LabelIntro")
            rs("Priority") = trs("Priority")
            rs("LabelContent") = trs("LabelContent")
            rs("AreaCollectionID") = trs("AreaCollectionID")
            rs.Update
            trs.MoveNext
        Loop
    End If

    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    
    tconn.Close
    Set tconn = Nothing
    Call WriteSuccessMsg("�Ѿ��ɹ�����ѡ�е�ģ�����õ�����ָ�������ݿ��У�", ComeUrl)
End Sub

'=================================================
'��������Import
'��  �ã�����ģ���һ��
'=================================================
Sub Import()
    Response.Write "<form name='myform' method='post' action='Admin_Template.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>ģ�嵼�루��һ����</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;������Ҫ�����ģ�����ݿ���ļ����� "
    Response.Write "        <input name='TemplateMdb' type='text' id='TemplateMdb' value='../temp/Template.mdb' size='20' maxlength='50'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' ��һ�� '><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='Import2'>"
    Response.Write "        <input name='ProjectName' type='hidden' id='Action' value='" & ProjectName & "'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

'=================================================
'��������Import2
'��  �ã�����ģ��ڶ���
'=================================================
Sub Import2()
    On Error Resume Next

    Dim rs, sql
    Dim mdbname, tconn, trs, iCount
    Dim ModuleType, ChannelName
    
    '�������Ƶ������ 999999 ��ʾ����
    ModuleType = Trim(Request.Form("ModuleType"))

    If ModuleType = "" Then
        ModuleType = 999999
    Else
        ModuleType = PE_CLng(ModuleType)
    End If
    
    '��õ���ģ�����ݿ�·��
    mdbname = Replace(Trim(Request.Form("Templatemdb")), "'", "")

    If mdbname = "" Then
        mdbname = Replace(Trim(Request.QueryString("Templatemdb")), "'", "")
    End If

    mdbname = Replace(mdbname, "��", "/") '��ֹ�ⲿ���Ӱ�ȫ����

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д����ģ�����ݿ���"
        Exit Sub
    End If

    '��������ģ�����ݿ�
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݿ����ʧ�ܣ����Ժ����ԣ�����ԭ��" & Err.Description
        Err.Clear
        Exit Sub
    End If

    Dim i, iVersion
    iVersion = 4
    Set trs = tconn.Execute("select top 1 * from PE_Template")

    For i = 0 To trs.Fields.Count - 1

        If LCase(trs.Fields(i).name) = "moduletype" Then
            iVersion = 5
            Exit For
        End If

    Next

    Set trs = Nothing
    
    Response.Write "<form name='myform' method='post' action='Admin_Template.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>ģ�嵼�루�ڶ�����</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td height='100' align='center'>"
    Response.Write "        <br>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr align='center'>"
    Response.Write "            <td><strong>��Ҫ�����ģ��</strong></td>"
    Response.Write "            <td></td>"
    Response.Write "            <td><strong>Ҫ���뵽�Ǹ�Ƶ��</strong></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td colspan='3'height='25'>"

    If iVersion = 4 Then
        Response.Write "��ѡ��Ƶ����"
    Else
        Response.Write "��ѡ��ģ�飺"
    End If

    Response.Write "              <select name='ModuleType' onChange='document.myform.submit();'>"

    If iVersion = 4 Then
        Call GetAllModule("4.03", ModuleType)
    Else
        Call GetAllModule("5.0", ModuleType)
    End If

    Response.Write "              </select>"
    Response.Write "              <br>"
    Response.Write "             </td>"
    Response.Write "           </tr>"
    Response.Write "           <tr>"
    Response.Write "            <td>"
    
    '��ʾģ��
    Response.Write "              <select name='TemplateID' size='2' multiple style='height:300px;width:250px;'>"
    
    '������ģ��Ϊ4.03��ʱ
    If iVersion = 4 Then
        '��ѯѡ���� ָ�� ���� �û��Զ��壨��2�� ���� ��ҳ��0�� ����ȫ��
        If ModuleType <> 999999 And ModuleType <> -2 Then
            sql = "select * from PE_Template where ChannelID = " & ModuleType & " Order by ChannelID asc"
        ElseIf ModuleType = -2 Then
            sql = "select ChannelID,TemplateID,TemplateName from PE_Template where ChannelID not in (0,1,2,3)"
        ElseIf ModuleType = 0 Then
            sql = "select ChannelID,TemplateID,TemplateName from PE_Template where ChannelID=0"
        Else
            sql = "select * from PE_Template"
        End If
    Else
        '5.0  ��ѯѡ���� ָ�� ���� ��ҳ��0�� ����ȫ��
        If ModuleType <> 999999 And ModuleType <> -1 Then
            sql = "select * from PE_Template where ModuleType = " & ModuleType & " Order by ChannelID asc"
        ElseIf ModuleType = 0 Then
            sql = "select ChannelID,TemplateID,TemplateName from PE_Template where ChannelID=0"
        Else
            sql = "select * from PE_Template"
        End If
    End If

    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, tconn, 1, 1

    If rs.BOF And rs.EOF Then
        'û��ģ��ʱָ���ر��ύ��ť
        Response.Write "                <option value='0'>û���κ�ģ��</option>"
        iCount = 0
    Else
        iCount = rs.RecordCount
        Do While Not rs.EOF
            Response.Write "            <option value='" & rs("TemplateID") & "'>" & rs("TemplateName") & "</option>"
            rs.MoveNext
        Loop
    End If

    rs.Close
    Set rs = Nothing
    Response.Write "                   </select>"
    Response.Write "                  </td>"
    Response.Write "                  <td width='80'>&nbsp;&nbsp;���뵽&gt;&gt;</td>"
    Response.Write "                  <td>"
    Response.Write "                    <select name='TargetChannelID' size='2' multiple style='height:300px;width:250px;'"

    If iCount = 0 Then Response.Write " disabled"
    Response.Write "                             >"
    
    If CLng(ModuleType) = 0 Then
        Response.Write "               <option value='0'>ͨ��ģ��</option>" & vbCrLf
    Else
        '�����У�������ҳ������ָ��ģ��
        If CLng(ModuleType) = 999999 Or CLng(ModuleType) = -2 Then
            sql = "select ChannelID,ChannelName from PE_Channel where ChannelType < 2 and Disabled=" & PE_False
        Else
            sql = "select ChannelID,ChannelName,ModuleType from PE_Channel where  ChannelType < 2 and Disabled=" & PE_False & " and ModuleType=" & ModuleType & "  Order by ChannelID asc"
        End If

        Set rs = Conn.Execute(sql)

        If rs.BOF And rs.EOF Then
            Response.Write "              <option value='0'>����û�н��������͵�Ƶ��</option>"
        Else
            If CLng(ModuleType) = 999999 Then
                Response.Write "            <option value='0'>ͨ��ģ��</option>" & vbCrLf
            End If

            Do While Not rs.EOF
                Response.Write "           <option value='" & rs("ChannelID") & "'>" & rs("ChannelName") & "</option>"
                rs.MoveNext
            Loop
        End If

        rs.Close
        Set rs = Nothing
    End If
    
    Response.Write "                    </select>"
    Response.Write "                   </td>"
    Response.Write "                  </tr>"
    Response.Write "                  <tr><td colspan='3' height='10'></td></tr>"
    Response.Write "                  <tr>"
    Response.Write "                    <td height='25' align='center'><b> ��ʾ����ס��Ctrl����Shift�������Զ�ѡ</b></td>"
    Response.Write "                    <td height='25' align='center'></td>"
    Response.Write "                    <td height='25' align='center'><b> ��ʾ����ס��Ctrl����Shift�������Զ�ѡ</b></td>"
    Response.Write "                  </tr>"
    Response.Write "                  <tr><td colspan='3' height='20'></td></tr>"
    Response.Write "                  <tr><td colspan='3' height='25' align='center'><input type='submit' name='Submit' value=' ����ģ�� ' onClick=""document.myform.Action.value='DoImport';"""

    If iCount = 0 Then Response.Write " disabled"
    Response.Write "                      ></td>"
    Response.Write "                 </tr>"
    Response.Write "               </table>"
    Response.Write "               <input name='TemplateMdb' type='hidden' id='TemplateMdb' value='" & mdbname & "'>"
    Response.Write "               <input name='Action' type='hidden' id='Action' value='Import2'>"
    Response.Write "               <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"

    If TempType > 0 Then Response.Write "<input name='TempType' type='hidden' id='TempType' value='" & TempType & "'>"
    Response.Write "               <br>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "       </table>"
    Response.Write "</form>"
End Sub

'=================================================
'��������DoImport
'��  �ã�����ģ�屣��
'=================================================
Sub DoImport()
    On Error Resume Next
    
    Dim trs, crs, mdbname, tconn
    Dim TemplateID, TargetChannelID, rs, sql, rsLabel, Table_PE_lable
    TemplateID = Trim(Request("TemplateID"))
    TargetChannelID = Trim(Request("TargetChannelID"))
    mdbname = Replace(Trim(Request.Form("Templatemdb")), "'", "")
    If IsValidID(TemplateID) = False Then
        TemplateID = ""
    End If
    If IsValidID(TargetChannelID) = False Then
        TargetChannelID = ""
    End If

    If TemplateID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����ģ��</li>"
    End If

    If TargetChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����Ƶ��ģ��</li>"
    End If

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д����ģ�����ݿ���"
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If
        
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        ErrMsg = ErrMsg & "<li>���ݿ����ʧ�ܣ����Ժ����ԣ�����ԭ��" & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    If FoundInArr(TargetChannelID, 0, ",") = True Then
      
        Set trs = tconn.Execute(" select * from PE_Template where TemplateID in (" & TemplateID & ") and ChannelID=0 order by TemplateID")
        Set rs = Server.CreateObject("adodb.recordset")
        rs.Open "select top 1 * from PE_Template", Conn, 1, 3

        Do While Not trs.EOF
            rs.addnew
            rs("ChannelID") = 0
            rs("TemplateName") = trs("TemplateName")
            rs("TemplateType") = trs("TemplateType")
            rs("TemplateContent") = trs("TemplateContent")
            rs("ProjectName") = ProjectName
            rs("IsDefault") = False
            rs("IsDefaultInProject") = False
            rs("Deleted") = False
            rs.Update
            trs.MoveNext
        Loop
    
        Set trs = Nothing
        rs.Close
        Set rs = Nothing
    End If
       
    Dim i, iVersion
    iVersion = 4
    Set crs = Conn.Execute("select ChannelID,ChannelName,ModuleType from PE_Channel where ChannelID in (" & TargetChannelID & ")")
    Set trs = tconn.Execute(" select * from PE_Template where TemplateID in (" & TemplateID & ")  order by TemplateID")

    For i = 0 To trs.Fields.Count - 1

        If LCase(trs.Fields(i).name) = "moduletype" Then
            iVersion = 5
            Exit For
        End If

    Next

    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open "select top 1 * from PE_Template", Conn, 1, 3
    
    Do While Not crs.EOF
        trs.MoveFirst
        If Not trs.EOF Then
            Do While Not trs.EOF
                If iVersion = 5 Then   '�����5.0���ģ�����ݿ�
                    If trs("ModuleType") = crs("ModuleType") Then
                        rs.addnew
                        rs("ChannelID") = crs("ChannelID")
                        rs("TemplateName") = trs("TemplateName")
                        rs("TemplateType") = trs("TemplateType")
                        rs("TemplateContent") = trs("TemplateContent")
                        rs("ProjectName") = ProjectName
                        rs("IsDefault") = False
                        rs("IsDefaultInProject") = False
                        rs("Deleted") = False
                        rs.Update
                    End If
                Else  '�����4.0���ģ�����ݿ�
                    rs.addnew
                    rs("ChannelID") = crs("ChannelID")
                    rs("TemplateName") = trs("TemplateName")
                    rs("TemplateType") = trs("TemplateType")
                    rs("TemplateContent") = trs("TemplateContent")
                    rs("ProjectName") = ProjectName
                    rs("IsDefault") = False
                    rs("IsDefaultInProject") = False
                    rs("Deleted") = False
                    rs.Update
                End If
                trs.MoveNext
            Loop
        End If
        crs.MoveNext
    Loop
    Set crs = Nothing
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    
    '�ж�PE_Label ���Ƿ����
    Table_PE_lable = True
    tconn.Execute ("select LabelName from PE_Label")

    If Err Then
        Table_PE_lable = False
    End If

    If Table_PE_lable = True Then
        '���ǵ����ǩ
        Set rsLabel = tconn.Execute("select * from PE_Label")
        Set trs = Server.CreateObject("adodb.recordset")
        trs.Open "select * from PE_Label", Conn, 1, 3
        
        Do While Not rsLabel.EOF
            Set rs = Conn.Execute("select LabelName from PE_Label where LabelName='" & rsLabel("LabelName") & "'")

            If rs.BOF And rs.EOF Then
                trs.addnew
                trs("LabelName") = rsLabel("LabelName")
                trs("LabelType") = rsLabel("LabelType")
                trs("LabelIntro") = rsLabel("LabelIntro")
                trs("Priority") = rsLabel("Priority")
                trs("LabelContent") = rsLabel("LabelContent")
                trs.Update
            End If

            rsLabel.MoveNext
        Loop
        
        Set trs = Nothing
        Set rsLabel = Nothing
        rs.Close
        Set rs = Nothing
    End If
    
    tconn.Close
    Set tconn = Nothing
       
    Call WriteSuccessMsg("�Ѿ��ɹ���ָ�������ݿ��е���ѡ�е�ģ�壡", ComeUrl & "?ChannelID=" & ChannelID & "&TempType=" & TempType & "&Action=Import2&Templatemdb=" & Replace(mdbname, "/", "��") & "")

End Sub

'=================================================
'��������ChannelCopyTemplate
'��  �ã�Ƶ������ģ��
'=================================================
Sub ChannelCopyTemplate()

    Dim rs, sql
    Dim trs, iCount
    Dim TemplateChannelID, ModuleType
    
    '���ģ��Ƶ��ID
    TemplateChannelID = Trim(Request.Form("TemplateChannelID"))
    If IsValidID(TemplateChannelID) = False Then
        TemplateChannelID = ""
    End If

    If TemplateChannelID = "" Then
        TemplateChannelID = 999999
    End If
    
    Response.Write "<form name='myform' method='post' action='Admin_Template.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>����Ƶ��ģ�帴��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'> "
    Response.Write "      <td height='100' align='center'>"
    Response.Write "        <br>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr align='center'>"
    Response.Write "            <td><strong>ѡ��Ҫ���Ƶ�Ƶ��ģ��</strong></td>"
    Response.Write "            <td></td>"
    Response.Write "            <td><strong>Ҫ���Ƶ��Ǹ�Ƶ��</strong></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td colspan='3'height='25'>"
    '��ʾ����ϵͳ���е�Ƶ��
    sql = "select ChannelID,ChannelName,ModuleType from PE_Channel where ModuleType <> 0"
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 3
    Response.Write "             <select name='TemplateChannelID' onChange='document.myform.submit();'>"

    If rs.BOF And rs.EOF Then
        Response.Write "           <option value=''>����û�н���Ƶ��</option>"
    Else

        Do While Not rs.EOF
            Response.Write "       <option value='" & rs("ChannelID") & "'"

            If rs("ChannelID") = PE_CLng(TemplateChannelID) Then Response.Write " selected"
            Response.Write ">" & rs("ChannelName") & "</option>" & vbCrLf
            rs.MoveNext
        Loop

        Response.Write "           <option value='999999'"

        If PE_CLng(TemplateChannelID) = 999999 Then Response.Write " selected"
        Response.Write ">����Ƶ��</option>" & vbCrLf
    End If

    Response.Write "         </select>"
    rs.Close
    Set rs = Nothing
    
    Response.Write "             <br>"
    Response.Write "            </td>"
    Response.Write "           </tr>"
    Response.Write "           <tr>"
    Response.Write "             <td>"
    Response.Write "               <select name='TemplateID' size='2' multiple style='height:300px;width:250px;'>"
           
    '�ж������л���ָ��
    If PE_CLng(TemplateChannelID) = 999999 Then
        sql = "select t.ChannelID,t.TemplateID,t.TemplateName,c.ModuleType from PE_Template t inner join PE_Channel c on t.ChannelID=c.ChannelID where t.ProjectName='" & ProjectName & "'"
    Else
        sql = "select t.ChannelID,t.TemplateID,t.TemplateName,c.ModuleType from PE_Template t inner join PE_Channel c on t.ChannelID=c.ChannelID where t.ChannelID=" & TemplateChannelID & " And t.ProjectName='" & ProjectName & "'  Order by t.ChannelID asc"
    End If

    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 1

    If rs.BOF And rs.EOF Then
        Response.Write "      <option value=''>û���κ�ģ��</option>"
        '�ر��ύ��ť
        iCount = 0
        '-999999 Ϊ û��ģ��
        ModuleType = -999999
    Else
        '�õ�ֵ֤�������ݿ����ύ��ť
        iCount = rs.RecordCount
        '�õ�ģ�������
        ModuleType = rs("ModuleType")

        Do While Not rs.EOF
            Response.Write "  <option value='" & rs("TemplateID") & "'>" & rs("TemplateName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "               </select>"
    Response.Write "              </td>"
    Response.Write "              <td width='80'>&nbsp;&nbsp;���Ƶ�&gt;&gt;></td>"
    Response.Write "              <td>"
    Response.Write "                <select name='TargetChannelID' size='2' multiple style='height:300px;width:250px;'>"

    '��û��ģ��ʱ
    If ModuleType = -999999 Then
        Response.Write "              <option value=''>û�пɸ��Ƶ�ģ��</option>" & vbCrLf
    Else

        '�ж���ȫ�� ���� ָ��
        If PE_CLng(TemplateChannelID) = 999999 Then
            sql = "select ChannelID,ChannelName from PE_Channel where ModuleType<>0"
        Else
            sql = "select ChannelID,ChannelName,ModuleType from PE_Channel where ChannelID<>" & TemplateChannelID & " and ModuleType=" & ModuleType & "  Order by ChannelID asc"
        End If

        Set rs = Conn.Execute(sql)

        If rs.BOF And rs.EOF Then
            iCount = 0
            Response.Write "          <option value=''>����û�н�����ͬ���͵�Ƶ��</option>"
        Else

            Do While Not rs.EOF
                Response.Write "      <option value='" & rs("ChannelID") & "'>" & rs("ChannelName") & "</option>"
                rs.MoveNext
            Loop

        End If

        rs.Close
        Set rs = Nothing
    End If

    Response.Write "                </select>"
    Response.Write "               </td>"
    Response.Write "             </tr>"
    Response.Write "             <tr><td colspan='3' height='10'></td></tr>"
    Response.Write "             <tr>"
    Response.Write "              <td height='25' align='center'><b> ��ʾ����ס��Ctrl����Shift�������Զ�ѡ</b></td>"
    Response.Write "              <td height='25' align='center'></td>"
    Response.Write "              <td height='25' align='center'><b> ��ʾ����ס��Ctrl����Shift�������Զ�ѡ</b></td>"
    Response.Write "             </tr>"
    Response.Write "             <tr><td colspan='3' height='20'></td></tr>"
    Response.Write "             <tr><td colspan='3' height='25' align='center'><input type='submit' name='Submit' value=' ����ģ�� ' onClick=""document.myform.Action.value='DoCopy';"""

    If iCount = 0 Then Response.Write " disabled"
    Response.Write "                  ></td>"
    Response.Write "             </tr>"
    Response.Write "           </table>"
    Response.Write "           <input name='Action' type='hidden' id='Action' value='ChannelCopyTemplate'>"
    Response.Write "          <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"

    If TempType > 0 Then Response.Write "<input name='TempType' type='hidden' id='TempType' value='" & TempType & "'>"
    Response.Write "          <br>"
    Response.Write "        </td>"
    Response.Write "      </tr>"
    Response.Write "   </table>"
    Response.Write "</form>"
End Sub

'=================================================
'��������DoCopy
'��  �ã�Ƶ������ģ�屣��
'=================================================
Sub DoCopy()
    ' On Error Resume Next
    Dim trs, crs
    Dim TemplateID, TargetChannelID, rs, sql
    TemplateID = Trim(Request("TemplateID"))
    TargetChannelID = Trim(Request("TargetChannelID"))
    If IsValidID(TemplateID) = False Then
        TemplateID = ""
    End If
    If IsValidID(TargetChannelID) = False Then
        TargetChannelID = ""
    End If

    If TemplateID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ���Ƶ�ģ��</li>"
    End If

    If TargetChannelID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ŀ��Ƶ��</li>"
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If

    Set crs = Conn.Execute("select ChannelID,ChannelName,ModuleType from PE_Channel where ChannelID in (" & TargetChannelID & ")")
    Set trs = Conn.Execute("select T.ProjectName,T.TemplateName,T.TemplateType,T.TemplateContent,T.ChannelID,C.ChannelName,C.ModuleType from PE_Template T inner join PE_Channel c on T.ChannelID=C.ChannelID where T.TemplateID in (" & TemplateID & ")  order by T.TemplateID")

    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open "select top 1 * from PE_Template", Conn, 1, 3
    
    Do While Not crs.EOF

        If Not trs.EOF Then

            Do While Not trs.EOF

                If trs("ChannelID") <> crs("ChannelID") And trs("ModuleType") = crs("ModuleType") Then
                    rs.addnew
                    rs("ChannelID") = crs("ChannelID")
                    rs("TemplateName") = trs("TemplateName")
                    rs("TemplateType") = trs("TemplateType")
                    rs("TemplateContent") = trs("TemplateContent")
                    rs("IsDefault") = False
                    rs("ProjectName") = trs("ProjectName")
                    rs("IsDefaultInProject") = False
                    rs("Deleted") = False
                    rs.Update
                End If

                trs.MoveNext
            Loop

            trs.MoveFirst
        End If

        crs.MoveNext
    Loop

    Set crs = Nothing
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    Call WriteSuccessMsg("�Ѿ��ɹ������ģ�帴�ƣ�", ComeUrl & "?ChannelID=" & ChannelID & "&TempType=" & TempType & "&Action=ChannelCopyTemplate")
End Sub

'=================================================
'��������DoTemplateCopy
'��  �ã�ģ�帴�ƴ���
'=================================================
Sub DoTemplateCopy()
    Dim sql, rs, trs, TemplateID, TemplateName, FoundErr, ErrMsg
    FoundErr = False

    TemplateID = Trim(Request("TemplateID"))
    TemplateName = Trim(Request("TemplateName"))
    ProjectName = Trim(Request("ProjectName"))
    TemplateProjectID = PE_CLng(Trim(Request("TemplateProjectID")))
    If IsValidID(TemplateID) = False Then
        TemplateID = ""
    End If

    
    If TemplateID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����������Ŀ��ID���ԣ�</li>"
    End If
    
    If FoundErr <> True Then
        If InStr(TemplateID, ",") = 0 Then
            Set trs = Conn.Execute("Select TemplateID,ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,Deleted,IsDefaultInProject from PE_Template Where TemplateID=" & TemplateID)
        Else
            Set trs = Conn.Execute("Select TemplateID,ChannelID,TemplateName,TemplateType,TemplateContent,IsDefault,ProjectName,Deleted,IsDefaultInProject from PE_Template Where TemplateID in (" & TemplateID & ")")
        End If

        If trs.BOF And trs.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<br>��������û���ҵ���ģ����Ŀ��"
        Else
            Set rs = Server.CreateObject("adodb.recordset")
            rs.Open "select top 1 * from PE_Template", Conn, 1, 3

            Do While Not trs.EOF
                rs.addnew
                rs("ChannelID") = trs("ChannelID")
                rs("TemplateName") = trs("TemplateName") & " ����"
                rs("TemplateType") = trs("TemplateType")
                rs("TemplateContent") = trs("TemplateContent")
                rs("IsDefault") = False
                rs("ProjectName") = trs("ProjectName")
                rs("IsDefaultInProject") = False
                rs("Deleted") = trs("Deleted")
                rs.Update
                ErrMsg = ErrMsg & "<br>�µ�ģ�屣��Ϊ��<font color=red>" & rs("TemplateName") & "</font>"
                trs.MoveNext
            Loop

            rs.Close
            Set rs = Nothing
        End If

        Set trs = Nothing
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    Else
        Response.Write "<br>"
        Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
        Response.Write "  <tr align='center' class='title'><td height='22'><strong>��ϲ����</strong></td></tr>" & vbCrLf
        Response.Write "  <tr class='tdbg'><td height='100' valign='top'><br>" & TemplateName & " ģ�屸�����." & ErrMsg & "<br></td></tr>" & vbCrLf
        Response.Write "  <tr align='center' class='tdbg'><td>"
        Response.Write "</td></tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
      '  Response.Write "<meta http-equiv='refresh' content=3;url='Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=" & TemplateType & "&TempType=" & TempType & "&ProjectName=" & ProjectName & "&TemplateProjectID=" & TemplateProjectID & "'>"
        Call Refresh("Admin_Template.asp?ChannelID=" & ChannelID & "&TemplateType=" & TemplateType & "&TempType=" & TempType & "&ProjectName=" & ProjectName & "&TemplateProjectID=" & TemplateProjectID,3)
    End If

    Call CloseConn
End Sub

'=================================================
'��������BatchReplace
'��  �ã������滻
'=================================================
Sub BatchReplace()

    Dim rs, sql
    Dim ModuleType, TemplateID, TemplateChannelID
    Dim BatchType, BatchContent, TemplateType, TemplateReplace, TemplateReplaceStart, TemplateReplaceEnd, TemplateReplaceResult
    Dim ProjectName

    TemplateType = PE_CLng(Trim(Request.Form("TemplateType")))
    TemplateID = ReplaceBadChar(Trim(Request.Form("TemplateID")))
    TemplateChannelID = ReplaceBadChar(Trim(Request.Form("TemplateChannelID")))
    BatchType = PE_CLng(Trim(Request.Form("BatchType")))
    BatchContent = PE_CLng(Trim(Request.Form("BatchContent")))
    TemplateReplace = Trim(Request.Form("TemplateReplace"))
    TemplateReplaceStart = Trim(Request.Form("TemplateReplaceStart"))
    TemplateReplaceEnd = Trim(Request.Form("TemplateReplaceEnd"))
    TemplateReplaceResult = Trim(Request.Form("TemplateReplaceResult"))

    If ReplaceBadChar(Trim(Request.QueryString("ProjectName"))) = "" Then
        ProjectName = ReplaceBadChar(Trim(Request.Form("ProjectName")))
    Else
        ProjectName = ReplaceBadChar(Trim(Request.QueryString("ProjectName")))
    End If

    If TemplateType = 0 Then
        TemplateType = 1
    End If

    If BatchType = 0 Then
        BatchType = 1
    End If
        
    '�������Ƶ������ 999999 ��ʾ����
    ModuleType = Trim(Request.Form("ModuleType"))

    If ModuleType = "" Then
        ModuleType = 999999
    Else
        ModuleType = PE_CLng(ModuleType)
    End If

    Response.Write "<form method=""post"" action=""Admin_Template.asp"" name=""form1"" >" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr class='title'>" & vbCrLf
    Response.Write "    <td height='22' colspan='2' align='center'><b>ģ�������滻����</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td class='tdbg' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "          <tr class='tdbg5'>" & vbCrLf
    Response.Write "            <td align='left' ><INPUT TYPE='radio' NAME='TemplateType' value='1' " & IsRadioChecked(TemplateType, 1) & ""
    Response.Write " onClick=""javascript:TemplateID.style.display='';TemplateChannelID.style.display='none';ProjectName.style.display='none';ModuleType.style.display='none';BatchTemplateID.style.display='none';"""
    Response.Write " ><b>ѡ��Ҫ���滻<FONT color='red'>ģ��</FONT>ID</b></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr  class='tdbg5'>" & vbCrLf
    Response.Write "          <td align='left' ><INPUT TYPE='radio' NAME='TemplateType' value='2' " & IsRadioChecked(TemplateType, 2) & ""
    Response.Write " onClick=""javascript:TemplateID.style.display='none';TemplateChannelID.style.display='';ProjectName.style.display='none';ModuleType.style.display='none';BatchTemplateID.style.display='none';"""
    Response.Write " ><b>ѡ��Ҫ���滻<FONT color='blue'>Ƶ��</FONT>ID</b></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr  class='tdbg5'>" & vbCrLf
    Response.Write "            <td align='left'><INPUT TYPE='radio' NAME='TemplateType' value='3' " & IsRadioChecked(TemplateType, 3) & ""
    Response.Write " onClick=""javascript:TemplateID.style.display='none';TemplateChannelID.style.display='none';ProjectName.style.display=''; ModuleType.style.display='';BatchTemplateID.style.display='';"""
    Response.Write "><b>ѡ��Ҫ���滻��<FONT color='#339900'>����</Font></b></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr  class='tdbg5'>" & vbCrLf
    Response.Write "            <td align='left' ><INPUT TYPE='Text' NAME='TemplateID' id='TemplateID' value='" & TemplateID & " ' size='40' " & IsStyleDisplay(TemplateType, 1) & "></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr  class='tdbg5'>" & vbCrLf
    Response.Write "            <td align='left'>" & vbCrLf
    sql = "SELECT DISTINCT t.ChannelID, c.ChannelName FROM PE_Template t INNER JOIN PE_Channel c ON t.ChannelID = c.ChannelID"
    Set rs = Conn.Execute(sql)
    Response.Write "<select name='TemplateChannelID' id='TemplateChannelID' size='2' multiple style='height:300px;width:250px;'  " & IsStyleDisplay(TemplateType, 2) & ">"

    If rs.BOF And rs.EOF Then
        Response.Write "<option value="" selected>��û�����Ƶ����</option> "
    Else
        Response.Write "<option selected value=" & rs("ChannelID") & ">" & rs("ChannelName") & "</option>"
        rs.MoveNext

        Do While Not rs.EOF
            Response.Write "<option value=" & rs("ChannelID") & ">" & rs("ChannelName") & "</option>"
            rs.MoveNext
        Loop

    End If

    Response.Write "</select>"
    rs.Close
    Set rs = Nothing
    Response.Write "            </td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf

    Response.Write "          <tr  class='tdbg5'>" & vbCrLf
    Response.Write "            <td>" & vbCrLf
    Response.Write "            <select name='ProjectName' id='ProjectName' style='width:150px;'  onChange='document.form1.submit();' " & IsStyleDisplay(TemplateType, 3) & ">"
    sql = "select * from PE_TemplateProject"
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>û���κη�����Ŀ</option>"
    Else
        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TemplateProjectName") & "' " & OptionValue(rs("TemplateProjectName"), ProjectName) & ">" & rs("TemplateProjectName") & "</option>"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Response.Write "            </select>"
    Response.Write "            <br>"
    '��ʾ����ϵͳ���е�Ƶ��
    Response.Write "              <select name='ModuleType' id='ModuleType' onChange='document.form1.submit();'  " & IsStyleDisplay(TemplateType, 3) & ">"
    Call GetAllModule("5.0", ModuleType)
    Response.Write "              </select>"
    Response.Write "             <br>"
    '��ʾģ��
    Response.Write "              <select name='BatchTemplateID' id='BatchTemplateID' size='2' multiple style='height:300px;width:250px;'  " & IsStyleDisplay(TemplateType, 3) & ">"

    '5.0  ��ѯѡ���� ָ�� ���� ��ҳ��0�� ����ȫ��
    If ModuleType <> 999999 And ModuleType <> -1 And ModuleType <> 0 Then
        sql = "select t.ChannelID,t.TemplateID,t.TemplateName,t.TemplateType,t.TemplateContent,t.IsDefault,c.ModuleType,t.ProjectName from PE_Template t left join PE_Channel c on t.ChannelID=c.ChannelID where c.ModuleType=" & ModuleType & " And t.ProjectName='" & ProjectName & "' order by t.TemplateID"
    ElseIf ModuleType = 0 Then
        sql = "select ChannelID,TemplateID,TemplateName,ProjectName from PE_Template where ChannelID=0  And ProjectName='" & ProjectName & "'"
    Else
        sql = "select * from PE_Template  Where ProjectName='" & ProjectName & "'"
        If FoundInArr(AllModules, "Supply", ",") = False Then
            sql = sql & " And ChannelID <> 999 "
        End If
        If FoundInArr(AllModules, "House", ",") = False Then
            sql = sql & " And ChannelID <> 998 "
        End If
        If FoundInArr(AllModules, "Job", ",") = False Then
            sql = sql & " And ChannelID <> 997 "
        End If
    End If

    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, Conn, 1, 1

    If rs.BOF And rs.EOF Then
        'û��ģ��ʱָ���ر��ύ��ť
        Response.Write "                <option value='0'>û���κ�ģ��</option>"
    Else
        Do While Not rs.EOF
            Response.Write "            <option value='" & rs("TemplateID") & "'>" & rs("TemplateName") & "</option>"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Response.Write "              </select>"
    Response.Write "            </td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td valign='top'>" & vbCrLf
    Response.Write "       <table width='100%' height='400' border='0' cellpadding='0' cellspacing='1'>"
    Response.Write "         <tr class='tdbg'>" & vbCrLf
    Response.Write "           <td width='150' align=""right""><strong>�滻���ݣ�</strong></td>" & vbCrLf
    Response.Write "           <td align='left'>" & vbCrLf
    Response.Write "            <INPUT TYPE='radio' NAME='BatchContent' value='1' " & IsRadioChecked(BatchContent, 1) & " >ģ������&nbsp;&nbsp;"
    Response.Write "            <INPUT TYPE='radio' NAME='BatchContent' value='2' " & IsRadioChecked(BatchContent, 2) & " >ģ������</td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class='tdbg'>" & vbCrLf
    Response.Write "           <td width='150' align=""right""><strong>�滻���ͣ�</strong></td>" & vbCrLf
    Response.Write "           <td align='left'>" & vbCrLf
    Response.Write "            <INPUT TYPE='radio' NAME='BatchType' value='1' onClick=""javascript:PE_TemplateReplaceStart.style.display='none';PE_TemplateReplaceEnd.style.display='none';PE_TemplateReplace.style.display='';"" " & IsRadioChecked(BatchType, 1) & ">���滻&nbsp;&nbsp;"
    Response.Write "            <INPUT TYPE='radio' NAME='BatchType' value='2'  onClick=""javascript:PE_TemplateReplaceStart.style.display='';PE_TemplateReplaceEnd.style.display='';PE_TemplateReplace.style.display='none';"" " & IsRadioChecked(BatchType, 2) & ">�߼��滻</td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_TemplateReplace' " & IsStyleDisplay(BatchType, 1) & "> " & vbCrLf
    Response.Write "           <td width=""150"" class=""tdbg"" align=""right""><strong> Ҫ�滻�Ĵ��룺&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='TemplateReplace' ROWS='' COLS='' style='width:400px;height:100px'>" & TemplateReplace & "</TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_TemplateReplaceStart' " & IsStyleDisplay(BatchType, 2) & "> " & vbCrLf
    Response.Write "           <td width=""150"" class=""tdbg"" align=""right"" ><strong> Ҫ�滻�Ŀ�ʼ���룺&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='TemplateReplaceStart' ROWS='' COLS='' style='width:400px;height:100px'>" & TemplateReplaceStart & "</TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_TemplateReplaceEnd' " & IsStyleDisplay(BatchType, 2) & "> " & vbCrLf
    Response.Write "           <td width=""150"" class=""tdbg"" align=""right"" ><strong> Ҫ�滻�Ľ������룺&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='TemplateReplaceEnd' ROWS='' COLS='' style='width:400px;height:100px'>" & TemplateReplaceEnd & "</TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_TemplateReplaceResult'>" & vbCrLf
    Response.Write "           <td width=""150"" class=""tdbg""  align=""right""><strong> Ҫ�滻��Ĵ��룺&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='TemplateReplaceResult' ROWS='' COLS='' style='width:400px;height:100px'>" & TemplateReplaceResult & "</TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg""> " & vbCrLf
    Response.Write "           <td colspan=""2"" align=""center"" class=""tdbg"" height=""50"">" & vbCrLf
    Response.Write "            <input name=""Action"" type=""hidden"" id=""Action"" value=""BatchReplace"">" & vbCrLf
    Response.Write "            <input name=""ChannelID"" type=""hidden"" id=""ChannelID"" value=" & ChannelID & ">" & vbCrLf
    Response.Write "            <input name=""Cancel"" type=""button"" id=""Cancel"" value=""������һ��"" onClick=""window.location.href='javascript:history.go(-1)'"" >&nbsp;&nbsp;" & vbCrLf
    Response.Write "            <input  type=""submit"" name=""Submit"" value="" ��ʼ�滻 "" onClick=""document.form1.Action.value='DoBatchReplace';"" >" & vbCrLf
    Response.Write "           </td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "       </table>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "     </tr>" & vbCrLf
    Response.Write " </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

'=================================================
'��������DoBatchReplace
'��  �ã������滻����
'=================================================
Sub DoBatchReplace()
    Dim rs, sql
    Dim TemplateType, TemplateID, TemplateChannelID, BatchTemplateID
    Dim BatchType, BatchContent, BatchDataType, TemplateReplace, TemplateReplaceStart, TemplateReplaceEnd, TemplateReplaceResult
    Dim FoundErr, ErrMsg
        
    FoundErr = False
    TemplateType = PE_CLng(Trim(Request.Form("TemplateType")))
    TemplateID = ReplaceBadChar(Trim(Request.Form("TemplateID")))
    BatchTemplateID = Trim(Request.Form("BatchTemplateID"))
    TemplateChannelID = Trim(Request.Form("TemplateChannelID"))
    BatchType = PE_CLng(Trim(Request.Form("BatchType")))
    BatchContent = PE_CLng(Trim(Request.Form("BatchContent")))
    TemplateReplace = Trim(Request.Form("TemplateReplace"))
    TemplateReplaceStart = Trim(Request.Form("TemplateReplaceStart"))
    TemplateReplaceEnd = Trim(Request.Form("TemplateReplaceEnd"))
    TemplateReplaceResult = Trim(Request.Form("TemplateReplaceResult"))
    If IsValidID(BatchTemplateID) = False Then
        BatchTemplateID = ""
    End If
    If IsValidID(TemplateChannelID) = False Then
        TemplateChannelID = ""
    End If

    If TemplateType = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>û��ѡ��Ҫ�滻��ģ������</li>"
    End If

    If BatchContent = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>û��ѡ��ģ����������</li>"
    End If

    If BatchType = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>û��ѡ��ģ���滻�ַ�����</li>"
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    If TemplateType = 1 Then
        If TemplateID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>û��ģ��ID��,�뷵������Ҫ�滻��ģ��ID</li>"
        End If

    ElseIf TemplateType = 2 Then

        If TemplateChannelID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>û��Ƶ��ID��,�뷵������Ҫ�滻��ģ��Ƶ��ID</li>"
        End If

    ElseIf TemplateType = 3 Then

        If BatchTemplateID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>û��ģ��ID��,�뷵������Ҫ�滻��ģ��ID</li>"
        End If

        TemplateID = BatchTemplateID
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ѡ���ģ�����Ͳ���</li>"
    End If

    If BatchType = 1 Then
        If TemplateReplace = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����Ҫ�滻�Ĵ��벻��Ϊ��</li>"
        End If

    ElseIf BatchType = 2 Then

        If TemplateReplaceStart = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����Ҫ�滻�Ŀ�ʼ���벻��Ϊ��</li>"
        End If

        If TemplateReplaceEnd = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����Ҫ�滻��Ľ������벻��Ϊ��</li>"
        End If

    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ѡ��ģ���滻�ַ����Ͳ���</li>"
    End If

    If TemplateReplaceResult = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����Ҫ�滻��Ĵ��벻��Ϊ��</li>"
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    Response.Write "<li>�����滻����</li>&nbsp;&nbsp;"
    Set rs = Server.CreateObject("ADODB.Recordset")

    sql = "select TemplateID,ChannelID,TemplateName,TemplateContent from PE_Template where "

    If TemplateType = 1 Or TemplateType = 3 Then
        If InStr(TemplateID, ",") > 0 Then
            sql = sql & " TemplateID in (" & TemplateID & ")"
        Else
            sql = sql & " TemplateID=" & TemplateID
        End If

    ElseIf TemplateType = 2 Then

        If InStr(TemplateChannelID, ",") > 0 Then
            sql = sql & " ChannelID in (" & TemplateChannelID & ")"
        Else
            sql = sql & " ChannelID=" & TemplateChannelID
        End If
    End If

    rs.Open sql, Conn, 1, 3

    If BatchContent = 1 Then
        BatchDataType = "Name"
    Else
        BatchDataType = "Content"
    End If

    Do While Not rs.EOF
        If BatchType = 1 Then
            If InStr(rs("Template" & BatchDataType & ""), TemplateReplace) <> 0 Then
                rs("Template" & BatchDataType & "") = Replace(rs("Template" & BatchDataType & ""), TemplateReplace, TemplateReplaceResult)
                Response.Write "<br>&nbsp;&nbsp;" & rs("TemplateName") & "..<font color='#009900'>ģ���滻�ɹ���</font>"
            Else
                Response.Write "<br>&nbsp;&nbsp;" & rs("TemplateName") & "..<font color='#FF0000'>ģ���滻���벻����,�����滻��</font>"
            End If

        ElseIf BatchType = 2 Then
            rs("Template" & BatchDataType & "") = BatchReplaceString(rs("Template" & BatchDataType & ""), TemplateReplaceStart, TemplateReplaceEnd, TemplateReplaceResult, rs("TemplateName"))
        End If

        rs.Update
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Response.Write "<br><center> <a href='Admin_Template.asp?Action=Main' >����ģ�����</a> </center>"
End Sub

'=================================================
'��������StrJS_Template
'��  �ã���ʾ��ǰƵ����ģ������
'=================================================
Sub StrJS_Template()
    Dim TrueSiteUrl
    TrueSiteUrl = Trim(Request.ServerVariables("HTTP_HOST"))
     
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "    addeditorcss=false;" & vbCrLf
    Response.Write "    addeditorcss2=false;" & vbCrLf
    Response.Write "    var strTemplateLabel;" & vbCrLf
    Response.Write "    var strTemplateLabel2;" & vbCrLf
    Response.Write "    function ResumeError() {" & vbCrLf
    Response.Write "        return true;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    window.onerror = ResumeError;" & vbCrLf
    Response.Write "    function sizeContent(num,objname){" & vbCrLf
    Response.Write "        var obj = document.getElementById(objname);" & vbCrLf
    Response.Write "        if (parseInt(obj.rows)+num>=1) {" & vbCrLf
    Response.Write "            obj.rows = parseInt(obj.rows) + num;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        if (num>0){" & vbCrLf
    Response.Write "            obj.width=""90%"";" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function copy(num) {" & vbCrLf
    Response.Write "        if (num==1) {" & vbCrLf
    Response.Write "            var content= document.form1.Content.value;" & vbCrLf
    Response.Write "            document.form1.Content.value=content;" & vbCrLf
    Response.Write "            document.form1.Content.focus();" & vbCrLf
    Response.Write "            document.form1.Content.select();" & vbCrLf
    Response.Write "            textRange = document.form1.Content.createTextRange();" & vbCrLf
    Response.Write "            textRange.execCommand(""Copy"");" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        else {" & vbCrLf
    Response.Write "            document.form1.Content2.focus();" & vbCrLf
    Response.Write "            document.form1.Content2.select();" & vbCrLf
    Response.Write "            textRange = document.form1.Content2.createTextRange();" & vbCrLf
    Response.Write "            textRange.execCommand(""Copy"");" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function LoadEditorAlgebra(num){" & vbCrLf
    Response.Write "        if (num==1){" & vbCrLf
    Response.Write "            document.form1.Content.rows=30;" & vbCrLf
    Response.Write "            document.form1.rollContent.rows=31;" & vbCrLf
    Response.Write "            showAlgebra.style.display="""";" & vbCrLf
    Response.Write "            showeditor.style.display=""none"";" & vbCrLf
    Response.Write "            showLabel.style.display=""none"";" & vbCrLf
    Response.Write "            Navigation1.style.display="""";" & vbCrLf
    Response.Write "            CommonLabel1.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation1.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation3.style.display="""";" & vbCrLf
    Response.Write "            document.form1.Editorfullscreen.disabled=true;" & vbCrLf
    Response.Write "            document.form1.Copy.disabled=false;" & vbCrLf
    Response.Write "            if (addeditorcss==false){" & vbCrLf
    Response.Write "                addeditorcss=true;" & vbCrLf
    Response.Write "                setContent('get',1)" & vbCrLf
    Response.Write "                editor.yToolbarsCss();" & vbCrLf
    Response.Write "                editor.showBorders();" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                setContent('get',1)" & vbCrLf
    Response.Write "                editor.yToolbarsCss()" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            document.form1.Content2.rows=30;" & vbCrLf
    Response.Write "            document.form1.rollContent2.rows=31;" & vbCrLf
    Response.Write "            showAlgebra2.style.display="""";" & vbCrLf
    Response.Write "            showeditor2.style.display=""none"";" & vbCrLf
    Response.Write "            showLabel2.style.display=""none"";" & vbCrLf
    Response.Write "            Navigation12.style.display="""";" & vbCrLf
    Response.Write "            CommonLabel2.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation2.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation2.style.display="""";" & vbCrLf
    Response.Write "            document.form1.Editorfullscreen2.disabled=true;" & vbCrLf
    Response.Write "            document.form1.Copy2.disabled=false;" & vbCrLf
    Response.Write "            if (addeditorcss2==false){" & vbCrLf
    Response.Write "                addeditorcss2=true;" & vbCrLf
    Response.Write "                setContent('get',2);" & vbCrLf
    Response.Write "                editor2.yToolbarsCss();" & vbCrLf
    Response.Write "                editor2.showBorders();" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                setContent('get',2)" & vbCrLf
    Response.Write "                editor.yToolbarsCss()" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write " function LoadEditorEdit(num){" & vbCrLf
    Response.Write "        if (num==1){" & vbCrLf
    Response.Write "            showAlgebra.style.display=""none"";" & vbCrLf
    Response.Write "            showeditor.style.display="""";" & vbCrLf
    Response.Write "            showLabel.style.display=""none"";" & vbCrLf
    Response.Write "            Navigation1.style.display=""none"";" & vbCrLf
    Response.Write "            CommonLabel1.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation1.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation3.style.display=""none"";" & vbCrLf
    Response.Write "            CloseNavigation1.style.display=""none"";" & vbCrLf
    Response.Write "            CloseNavigation3.style.display=""none"";" & vbCrLf
    Response.Write "            document.form1.Editorfullscreen.disabled=false;" & vbCrLf
    Response.Write "            document.form1.Copy.disabled=true;" & vbCrLf
    Response.Write "            if (addeditorcss==false){" & vbCrLf
    Response.Write "                addeditorcss=true;" & vbCrLf
    Response.Write "                setContent('set',1);" & vbCrLf
    Response.Write "                editor.yToolbarsCss();" & vbCrLf
    Response.Write "                editor.showBorders();" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                setContent('set',1)" & vbCrLf
    Response.Write "                editor.yToolbarsCss()" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            showAlgebra2.style.display=""none"";" & vbCrLf
    Response.Write "            showeditor2.style.display="""";" & vbCrLf
    Response.Write "            showLabel2.style.display=""none"";" & vbCrLf
    Response.Write "            Navigation12.style.display=""none"";" & vbCrLf
    Response.Write "            CommonLabel2.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation2.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation2.style.display=""none"";" & vbCrLf
    Response.Write "            CloseNavigation4.style.display=""none"";" & vbCrLf
    Response.Write "            CloseNavigation4.style.display=""none"";" & vbCrLf
    Response.Write "            document.form1.Editorfullscreen2.disabled=false;" & vbCrLf
    Response.Write "            document.form1.Copy2.disabled=true;" & vbCrLf
    Response.Write "            if (addeditorcss2==false){" & vbCrLf
    Response.Write "                addeditorcss2=true;" & vbCrLf
    Response.Write "                setContent('set',2);" & vbCrLf
    Response.Write "                editor2.yToolbarsCss();" & vbCrLf
    Response.Write "                editor2.showBorders();" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                setContent('set',2)" & vbCrLf
    Response.Write "                editor.yToolbarsCss()" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  function LoadEditorMix(num){" & vbCrLf
    Response.Write "        if (num==1){" & vbCrLf
    Response.Write "            document.form1.Content.rows=10;" & vbCrLf
    Response.Write "            document.form1.rollContent.rows=11;" & vbCrLf
    Response.Write "            showeditor.style.display="""";" & vbCrLf
    Response.Write "            showAlgebra.style.display="""";" & vbCrLf
    Response.Write "            showLabel.style.display=""none"";" & vbCrLf
    Response.Write "            Navigation1.style.display=""none"";" & vbCrLf
    Response.Write "            CommonLabel1.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation1.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation3.style.display=""none"";" & vbCrLf
    Response.Write "            document.form1.Editorfullscreen.disabled=false;" & vbCrLf
    Response.Write "            document.form1.Copy.disabled=false;" & vbCrLf
    Response.Write "            if (addeditorcss==false){" & vbCrLf
    Response.Write "                addeditorcss=true;" & vbCrLf
    Response.Write "                editor.yToolbarsCss()" & vbCrLf
    Response.Write "                editor.showBorders()" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                editor.yToolbarsCss()" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            document.form1.Content2.rows=10;" & vbCrLf
    Response.Write "            document.form1.rollContent2.rows=11;" & vbCrLf
    Response.Write "            showAlgebra2.style.display="""";" & vbCrLf
    Response.Write "            showeditor2.style.display="""";" & vbCrLf
    Response.Write "            showLabel2.style.display=""none"";" & vbCrLf
    Response.Write "            Navigation12.style.display=""none"";" & vbCrLf
    Response.Write "            CommonLabel2.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation2.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation4.style.display=""none"";" & vbCrLf
    Response.Write "            document.form1.Editorfullscreen2.disabled=false;" & vbCrLf
    Response.Write "            document.form1.Copy2.disabled=false;" & vbCrLf
    Response.Write "            if (addeditorcss2==false){" & vbCrLf
    Response.Write "                addeditorcss2=true;" & vbCrLf
    Response.Write "                editor2.yToolbarsCss();" & vbCrLf
    Response.Write "                editor2.showBorders();" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                editor.yToolbarsCss()" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  function OpenNavigation(TemplateType) {" & vbCrLf
    Response.Write "        if (TemplateType==1){" & vbCrLf
    Response.Write "            showLabel.style.display="""";" & vbCrLf
    Response.Write "            CloseNavigation1.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation1.style.display=""none"";" & vbCrLf
    Response.Write "            CloseNavigation3.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation3.style.display=""none"";" & vbCrLf
    Response.Write "            parent.parent.frame.cols='0,*';" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            showLabel2.style.display="""";" & vbCrLf
    Response.Write "            CloseNavigation2.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation2.style.display=""none"";" & vbCrLf
    Response.Write "            CloseNavigation4.style.display="""";" & vbCrLf
    Response.Write "            OpenNavigation4.style.display=""none"";" & vbCrLf
    Response.Write "            parent.parent.frame.cols='0,*';" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  function CloseNavigation(TemplateType) {" & vbCrLf
    Response.Write "        if (TemplateType==1){" & vbCrLf
    Response.Write "            showLabel.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation1.style.display="""";" & vbCrLf
    Response.Write "            CloseNavigation1.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation3.style.display="""";" & vbCrLf
    Response.Write "            CloseNavigation3.style.display=""none"";" & vbCrLf
    Response.Write "            parent.parent.frame.cols='200,*';" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            showLabel2.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation2.style.display="""";" & vbCrLf
    Response.Write "            CloseNavigation2.style.display=""none"";" & vbCrLf
    Response.Write "            OpenNavigation4.style.display="""";" & vbCrLf
    Response.Write "            CloseNavigation4.style.display=""none"";" & vbCrLf
    Response.Write "            parent.parent.frame.cols='200,*';" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  function setContent2(num) {" & vbCrLf
    Response.Write "    if (num==1){" & vbCrLf
    Response.Write "        form1.Content.focus();" & vbCrLf
    Response.Write "        strTemplateLabel = document.selection.createRange();" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        form1.Content2.focus();" & vbCrLf
    Response.Write "        strTemplateLabel2 = document.selection.createRange();" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  function insertTemplateLabel(strLabel,insertTemplateType) {" & vbCrLf
    Response.Write "    if (insertTemplateType==1){" & vbCrLf
    Response.Write "        strTemplateLabel.text = strLabel" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strTemplateLabel2.text = strLabel" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "function SuperFunctionLabel(url,label,title,ModuleType,ChannelShowType,iwidth,iheight,TemplateType){" & vbCrLf
    Response.Write "    if (TemplateType==1){" & vbCrLf
    Response.Write "        form1.Content.focus();" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        form1.Content2.focus();" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    var str = document.selection.createRange();" & vbCrLf
    Response.Write "    var arr = showModalDialog(url+""?ChannelID=" & ChannelID & "&Action=Add&LabelName=""+label+""&Title=""+title+""&ModuleType=""+ModuleType+""&ChannelShowType=""+ChannelShowType+""&InsertTemplate=1"", """", ""dialogWidth:""+iwidth+""px; dialogHeight:""+iheight+""px; help: no; scroll:yes; status: yes""); " & vbCrLf
    Response.Write "    if (arr != null){" & vbCrLf
    Response.Write "        str.text = arr;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function fullscreen(num) {" & vbCrLf
    Response.Write "    window.open (""../Editor/editor_fullscreen.asp?ChannelID=" & ChannelID & "&num=""+num+"""", """", ""toolbar=no, menubar=no, top=0,left=0,width=1024,height=768, scrollbars=no, resizable=no,location=no, status=no"");" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function Templateskin(){" & vbCrLf
    Response.Write "    if(confirm('��ȷ��Ҫת������ƣ������û�б��浱ǰ������ģ���뱣��ģ�塣')){" & vbCrLf
    Response.Write "        window.location.href='Admin_Skin.asp?Action=Modify&SkinID=1&IsDefault=-1';" & vbCrLf
    Response.Write "    }  " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function show_ln(txt_ln,txt_main){" & vbCrLf
    Response.Write "    var txt_ln  = document.getElementById(txt_ln);" & vbCrLf
    Response.Write "    var txt_main  = document.getElementById(txt_main);" & vbCrLf
    Response.Write "    txt_ln.scrollTop = txt_main.scrollTop;" & vbCrLf
    Response.Write "    while(txt_ln.scrollTop != txt_main.scrollTop)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        txt_ln.value += (i++) + '\n';" & vbCrLf
    Response.Write "        txt_ln.scrollTop = txt_main.scrollTop;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    return;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function editTab(){" & vbCrLf
    Response.Write "    var code, sel, tmp, r" & vbCrLf
    Response.Write "    var tabs=''" & vbCrLf
    Response.Write "    event.returnValue = false" & vbCrLf
    Response.Write "    sel =event.srcElement.document.selection.createRange()" & vbCrLf
    Response.Write "    r = event.srcElement.createTextRange()" & vbCrLf
    Response.Write "    switch (event.keyCode){" & vbCrLf
    Response.Write "        case (8) :" & vbCrLf
    Response.Write "        if (!(sel.getClientRects().length > 1)){" & vbCrLf
    Response.Write "            event.returnValue = true" & vbCrLf
    Response.Write "            return" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        code = sel.text" & vbCrLf
    Response.Write "        tmp = sel.duplicate()" & vbCrLf
    Response.Write "        tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)" & vbCrLf
    Response.Write "        sel.setEndPoint('startToStart', tmp)" & vbCrLf
    Response.Write "        sel.text = sel.text.replace(/\t/gm, '')" & vbCrLf
    Response.Write "        code = code.replace(/\t/gm, '').replace(/\r\n/g, '\r')" & vbCrLf
    Response.Write "        r.findText(code)" & vbCrLf
    Response.Write "        r.select()" & vbCrLf
    Response.Write "        break" & vbCrLf
    Response.Write "    case (9) :" & vbCrLf
    Response.Write "        if (sel.getClientRects().length > 1){" & vbCrLf
    Response.Write "            code = sel.text" & vbCrLf
    Response.Write "            tmp = sel.duplicate()" & vbCrLf
    Response.Write "            tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)" & vbCrLf
    Response.Write "            sel.setEndPoint('startToStart', tmp)" & vbCrLf
    Response.Write "            sel.text = '\t'+sel.text.replace(/\r\n/g, '\r\t')" & vbCrLf
    Response.Write "            code = code.replace(/\r\n/g, '\r\t')" & vbCrLf
    Response.Write "            r.findText(code)" & vbCrLf
    Response.Write "            r.select()" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            sel.text = '\t'" & vbCrLf
    Response.Write "            sel.select()" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        break" & vbCrLf
    Response.Write "    case (13) :" & vbCrLf
    Response.Write "        tmp = sel.duplicate()" & vbCrLf
   ' Response.write "        tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)" & vbCrLf
   ' Response.write "        tmp.setEndPoint('endToEnd', sel)" & vbCrLf
    Response.Write "        for (var i=0; tmp.text.match(/[\t]+/g) && i<tmp.text.match(/[\t]+/g)[0].length; i++) tabs += '\t'" & vbCrLf
    Response.Write "        sel.text = '\r\n'+tabs" & vbCrLf
    Response.Write "        sel.select()" & vbCrLf
    Response.Write "        break" & vbCrLf
    Response.Write "    default  :" & vbCrLf
    Response.Write "        event.returnValue = true" & vbCrLf
    Response.Write "        break" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</script>" & vbCrLf
    '=================================================
    '��  �ã����ǿͻ��˴���
    '=================================================
    Response.Write "<script language=""VBScript"">" & vbCrLf
    Response.Write "    Dim Strsave,Strsave2,addeditorcss3" & vbCrLf
    Response.Write "    Dim regEx, Match, Matches, StrBody,strTemp,strTemp2,strMatch,arrMatch,i" & vbCrLf
    Response.Write "    Dim Content1,Content2,Content3,Content4,TemplateContent,TemplateContent2,TemplateContent3,arrContent,EditorContent" & vbCrLf
    Response.Write "    Set regEx = New RegExp" & vbCrLf
    Response.Write "    regEx.IgnoreCase = True" & vbCrLf
    Response.Write "    regEx.Global = True" & vbCrLf
    Response.Write "    Strsave=""A""" & vbCrLf
    Response.Write "    Strsave2=""A""" & vbCrLf
    Response.Write "    Sub CheckForm(Num)" & vbCrLf
    Response.Write "        if document.form1.TemplateName.value="""" then" & vbCrLf
    Response.Write "            alert ""ģ�����Ʋ���Ϊ�գ�""" & vbCrLf
    Response.Write "            document.form1.TemplateName.focus()" & vbCrLf
    Response.Write "            exit sub" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        if document.form1.Content.value="""" then" & vbCrLf
    Response.Write "            alert ""ģ�������ݲ���Ϊ�գ�""" & vbCrLf
    Response.Write "            editor.HtmlEdit.focus()" & vbCrLf
    Response.Write "            exit sub" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        if Num=2 then" & vbCrLf
    Response.Write "            if document.form1.Content2.value="""" then " & vbCrLf
    Response.Write "                alert ""С��ģ�������ݲ���Ϊ�գ�""" & vbCrLf
    Response.Write "                exit sub" & vbCrLf
    Response.Write "            End if" & vbCrLf
    Response.Write "            if Strsave=""B"" then setContent ""get"",1" & vbCrLf
    Response.Write "            if Strsave2=""B"" then setContent ""get"",2" & vbCrLf
    Response.Write "            document.form1.EditorContent.value=""""" & vbCrLf
    Response.Write "            document.form1.EditorContent2.value=""""" & vbCrLf
    Response.Write "        Else" & vbCrLf
    Response.Write "            if Strsave=""B"" then setContent ""get"",1" & vbCrLf
    Response.Write "            document.form1.EditorContent.value=""""" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        form1.submit" & vbCrLf
    Response.Write "    End Sub" & vbCrLf

    Call Resumeblank
    '=================================================
    '��  �ã��ͻ��˽��洫ֵ
    '=================================================
    Response.Write "    function setContent(zhi,TemplateType)" & vbCrLf
    Response.Write "    if zhi=""get"" then" & vbCrLf
    Response.Write "        if TemplateType=1 then" & vbCrLf
    Response.Write "            if Strsave=""A"" then Exit Function" & vbCrLf
    Response.Write "            Strsave=""A""" & vbCrLf
    Response.Write "            TemplateContent= document.form1.Content.value" & vbCrLf
    Response.Write "            TemplateContent2= editor.HtmlEdit.document.body.innerHTML" & vbCrLf
    Response.Write "        else" & vbCrLf
    Response.Write "            if Strsave2=""A"" then Exit Function" & vbCrLf
    Response.Write "            Strsave2=""A""" & vbCrLf
    Response.Write "            TemplateContent= document.form1.Content2.value" & vbCrLf
    Response.Write "            TemplateContent2= editor2.HtmlEdit.document.body.innerHTML" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        if TemplateContent="""" then " & vbCrLf
    Response.Write "            alert ""��ɾ���˴������ҳ�����������д��ҳ ��""" & vbCrLf
    Response.Write "            Exit function" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        if Instr(TemplateContent,""<body>"")=0 then" & vbCrLf
    Response.Write "            regEx.Pattern = ""(\<body)(.[^\<]*)(\>)""" & vbCrLf
    Response.Write "            Set Matches = regEx.Execute(TemplateContent)" & vbCrLf
    Response.Write "            For Each Match In Matches" & vbCrLf
    Response.Write "                StrBody = Match.Value" & vbCrLf
    Response.Write "            Next" & vbCrLf
    Response.Write "            If StrBody = """"  Then" & vbCrLf
    Response.Write "                alert ""�����ص��ı���û�а��� <body> ����û�и�body �������ʹ��ҳ���ѿ�,�����ٸ��� <body> ��""" & vbCrLf
    Response.Write "                Exit function" & vbCrLf
    Response.Write "            End If " & vbCrLf
    Response.Write "        Else" & vbCrLf
    Response.Write "            StrBody=""<body>"" " & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        arrContent = Split(TemplateContent, StrBody)" & vbCrLf
    Response.Write "        if ubound(arrContent)=0 then " & vbCrLf
    Response.Write "           alert ""�����ص��ı���û�а��� <body> ����û�и�body �������ʹ��ҳ���ѿ�,�����ٸ��� <body> ��""" & vbCrLf
    Response.Write "           exit function" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        Content1 = arrContent(0) & StrBody" & vbCrLf
    Response.Write "        Content2 = arrContent(1)" & vbCrLf
    Response.Write "        regEx.Pattern = ""\<IMG(.[^\<]*?)\}['|""""]\>""" & vbCrLf
    Response.Write "        Set Matches = regEx.Execute(TemplateContent2)" & vbCrLf
    Response.Write "        For Each Match In Matches" & vbCrLf
    Response.Write "            regEx.Pattern = ""\{\$(.*?)\}""" & vbCrLf
    Response.Write "            Set strTemp = regEx.Execute(replace(Match.Value,"" "",""""))" & vbCrLf
    Response.Write "            For Each Match2 In strTemp" & vbCrLf
    Response.Write "                strTemp2 = Replace(Match2.Value, ""?"", """""""")" & vbCrLf
    Response.Write "                TemplateContent2 = Replace(TemplateContent2, Match.Value, ""<!--"" & strTemp2 & ""-->"")" & vbCrLf
    Response.Write "            Next" & vbCrLf
    Response.Write "        Next" & vbCrLf
    Response.Write "        regEx.Pattern = ""\<IMG(.[^\<]*?)\$\>""" & vbCrLf
    Response.Write "        Set Matches = regEx.Execute(TemplateContent2)" & vbCrLf
    Response.Write "        For Each Match In Matches" & vbCrLf
    Response.Write "            regEx.Pattern = ""\#\[(.*?)\]\#""" & vbCrLf
    Response.Write "            Set strTemp = regEx.Execute(Match.Value)" & vbCrLf
    Response.Write "            For Each Match2 In strTemp" & vbCrLf
    Response.Write "                strTemp2 = Replace(Match2.Value, ""&amp;"", ""&"")" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2, ""#"", """")" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2,""&13;&10;"",vbCrLf)" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2,""&9;"",vbTab)" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2,""��"",""'"")" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2, ""[!"", ""<"")" & vbCrLf
    Response.Write "                strTemp2 = Replace(strTemp2, ""!]"", "">"")" & vbCrLf
    Response.Write "                TemplateContent2 = Replace(TemplateContent2, Match.Value, strTemp2)" & vbCrLf
    Response.Write "            Next" & vbCrLf
    Response.Write "        Next" & vbCrLf
    Response.Write "        TemplateContent2=Replace(TemplateContent2, ""http://" & TrueSiteUrl & InstallDir & """,""{$InstallDir}"")" & vbCrLf
    Response.Write "        TemplateContent2=Replace(TemplateContent2, ""http://" & LCase(TrueSiteUrl) & LCase(InstallDir) & """,""{$InstallDir}"")" & vbCrLf
    Response.Write "        TemplateContent2=Resumeblank(TemplateContent2)" & vbCrLf
    Response.Write "        TemplateContent2=Replace(TemplateContent2,""{$InstallDir}{$rsClass_ClassUrl}"",""{$rsClass_ClassUrl}"")" & vbCrLf
    Response.Write "        regEx.Pattern = ""\{\$InstallDir\}editor.asp(.[^\<]*?)\#""" & vbCrLf
    Response.Write "        TemplateContent2 = regEx.Replace(TemplateContent2, ""#"")" & vbCrLf
    Response.Write "        if TemplateType =1 then" & vbCrLf
    Response.Write "            document.form1.Content.value=Content1& vbCrLf &TemplateContent2& vbCrLf &""</body>""& vbCrLf &""</html>""" & vbCrLf
    Response.Write "        else" & vbCrLf
    Response.Write "            document.form1.Content2.value=Content1 & vbCrLf &TemplateContent2& vbCrLf &""</body>""& vbCrLf &""</html>""" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "    Else" & vbCrLf
    Response.Write "        if TemplateType =1 then    " & vbCrLf
    Response.Write "            if Strsave=""B"" then Exit Function" & vbCrLf
    Response.Write "            Strsave=""B""" & vbCrLf
    Response.Write "            TemplateContent= document.form1.Content.value" & vbCrLf
    Response.Write "        else " & vbCrLf
    Response.Write "            if Strsave2=""B"" then Exit Function" & vbCrLf
    Response.Write "            Strsave2=""B""" & vbCrLf
    Response.Write "            TemplateContent= document.form1.Content2.value" & vbCrLf
    Response.Write "        End if    " & vbCrLf
    Response.Write "        if TemplateContent="""" then " & vbCrLf
    Response.Write "            alert ""��ɾ���˴������ҳ�����������д��ҳ ��""" & vbCrLf
    Response.Write "            Exit function" & vbCrLf
    Response.Write "           " & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        if Instr(TemplateContent,""<body>"")=0 then" & vbCrLf
    Response.Write "            regEx.Pattern = ""(\<body)(.[^\<]*)(\>)""" & vbCrLf
    Response.Write "            Set Matches = regEx.Execute(TemplateContent)" & vbCrLf
    Response.Write "            For Each Match In Matches" & vbCrLf
    Response.Write "                StrBody = Match.Value" & vbCrLf
    Response.Write "            Next" & vbCrLf
    Response.Write "            If StrBody = """"  Then" & vbCrLf
    Response.Write "                alert ""�����ص��ı���û�а��� <body> ����û�и�body �������ʹ��ҳ���ѿ�,�����ٸ��� <body> ��""" & vbCrLf
    Response.Write "                Exit function" & vbCrLf
    Response.Write "            End If " & vbCrLf
    Response.Write "        Else" & vbCrLf
    Response.Write "            StrBody=""<body>"" " & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        arrContent = Split(TemplateContent, StrBody)" & vbCrLf
    Response.Write "        if ubound(arrContent)=0 then " & vbCrLf
    Response.Write "           alert ""�����ص��ı���û�а��� <body> ����û�и�body �������ʹ��ҳ���ѿ�,�����ٸ��� <body> ��""" & vbCrLf
    Response.Write "           exit function" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        Content1 = arrContent(0) & StrBody" & vbCrLf
    Response.Write "        Content2 = arrContent(1)" & vbCrLf
    Response.Write "        Content2 = Replace(Content2, ""<!--$"", ""{$--"")" & vbCrLf
    Response.Write "        Content2 = Replace(Content2, ""$-->"", ""--}"")" & vbCrLf
    Response.Write "        Content2 = Replace(Content2, ""<!--{$"", ""{$"")" & vbCrLf
    Response.Write "        Content2 = Replace(Content2, ""}-->"", ""}"")" & vbCrLf
    Response.Write "        'ͼƬ�滻JS" & vbCrLf
    Response.Write "        regEx.Pattern = ""(\<Script)([\s\S]*?)(\<\/Script\>)""" & vbCrLf
    Response.Write "        Set Matches = regEx.Execute(Content2)" & vbCrLf
    Response.Write "        For Each Match In Matches" & vbCrLf
    Response.Write "            strTemp = Replace(Match.Value, ""<"", ""[!"")" & vbCrLf
    Response.Write "            strTemp = Replace(strTemp, "">"", ""!]"")" & vbCrLf
    Response.Write "            strTemp = Replace(strTemp, ""'"", ""��"")" & vbCrLf
    Response.Write "            strTemp = ""<IMG alt='#"" & strTemp & ""#' src=""""" & InstallDir & "editor/images/jscript.gif"""" border=0 $>""" & vbCrLf
    Response.Write "            Content2 = Replace(Content2, Match.Value, strTemp)" & vbCrLf
    Response.Write "        Next" & vbCrLf
    Response.Write "        'ͼƬ�滻������ǩ" & vbCrLf
    Response.Write "        regEx.Pattern = ""(\{\$GetPicArticle|\{\$GetArticleList|\{\$GetSlidePicArticle|\{\$GetPicSoft|\{\$GetSoftList|\{\$GetSlidePicSoft|\{\$GetPicPhoto|\{\$GetPhotoList|\{\$GetSlidePicPhoto|\{\$GetPicProduct|\{\$GetProductList|\{\$GetSlidePicProduct|\{\$GetPositionList|\{\$GetSearchResult)\((.*?)\)\}""" & vbCrLf
    Response.Write "        Content2 = regEx.Replace(Content2, ""<IMG src=""""" & InstallDir & "editor/images/label.gif"""" border=0 zzz='$1($2)}'>"")" & vbCrLf
    Response.Write "        regEx.Pattern = ""\{\$InstallDir\}""" & vbCrLf
    Response.Write "        Content2 = regEx.Replace(Content2,""http://" & TrueSiteUrl & InstallDir & """)" & vbCrLf
    Response.Write "        if TemplateType=1 then" & vbCrLf
    Response.Write "            editor.HtmlEdit.document.body.innerHTML=Content2" & vbCrLf
    Response.Write "            editor.showBorders()" & vbCrLf
    Response.Write "            editor.showBorders()" & vbCrLf
    Response.Write "        else" & vbCrLf
    Response.Write "            editor2.HtmlEdit.document.body.innerHTML=Content2" & vbCrLf
    Response.Write "            editor2.showBorders()" & vbCrLf
    Response.Write "            editor2.showBorders()" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "    End if" & vbCrLf
    Response.Write "    End function" & vbCrLf
    Response.Write "    function setstatus()" & vbCrLf 'Ϊ323 �����editor.asp ��Ч����
    Response.Write "    end function" & vbCrLf
    Response.Write "</script>" & vbCrLf
    
End Sub

'=================================================
'��������fullscreen
'��  �ã�ȫ��ģʽ
'=================================================
Sub fullscreen()
    Dim TrueSiteUrl

    If ChannelID = 0 Then
        Response.Write "Ƶ��������ʧ��"
        Response.End
    End If
            
    Response.Write "<HTML>" & vbCrLf
    Response.Write "<HEAD>" & vbCrLf
    Response.Write "<TITLE>HtmlEdit - ȫ���༭</TITLE>" & vbCrLf
    Response.Write "<META http-equiv=Content-Type content=""text/html; charset=gb2312"">" & vbCrLf
    Response.Write "</HEAD>" & vbCrLf
    Response.Write "<body leftmargin=0 topmargin=0 onunload=""Minimize()"">" & vbCrLf
    Response.Write "<input type=""hidden"" id=""ContentFullScreen"" name=""ContentFullScreen"" value="""">" & vbCrLf
    Response.Write "<script language=VBScript>" & vbCrLf
    Response.Write "   Dim Matches, Match, arrContent, Content1, Content2,Content3,Content5" & vbCrLf
    Response.Write "   Dim strTemp, strTemp2, StrBody,TemplateContent" & vbCrLf
    Response.Write "   Set regEx = New RegExp" & vbCrLf

    If Request.QueryString("num") = 1 Then
        Response.Write "ContentFullScreen.value=opener.editor.HtmlEdit.document.body.innerHTML" & vbCrLf
        Response.Write "TemplateContent= opener.document.form1.Content.value" & vbCrLf
    Else
        Response.Write "ContentFullScreen.value =opener.editor2.HtmlEdit.document.body.innerHTML" & vbCrLf
        Response.Write "TemplateContent= opener.document.form1.Content2.value" & vbCrLf
    End If

    Response.Write "   ContentFullScreen.value =""<html><head><META http-equiv=Content-Type content=text/html; charset=gb2312><link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'></head><body leftmargin=0 topmargin=0 >"" & ContentFullScreen.value" & vbCrLf
    Response.Write "   document.Write ""<iframe ID='EditorFullScreen' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=1&TemplateType=3&tContentid=ContentFullScreen' frameborder='0' scrolling=no width='100%' HEIGHT='100%'></iframe>""" & vbCrLf
    
    Call Resumeblank

    Response.Write "Function Minimize()" & vbCrLf
    Response.Write "       regEx.IgnoreCase = True" & vbCrLf
    Response.Write "       regEx.Global = True" & vbCrLf
    Response.Write "       regEx.Pattern = ""(\<body)(.[^\<]*)(\>)""" & vbCrLf
    Response.Write "       Set Matches = regEx.Execute(TemplateContent)" & vbCrLf
    Response.Write "         For Each Match In Matches" & vbCrLf
    Response.Write "            StrBody = Match.Value" & vbCrLf
    Response.Write "         Next" & vbCrLf
    Response.Write "         arrContent = Split(TemplateContent, StrBody)" & vbCrLf
    Response.Write "         Content1 = arrContent(0) & StrBody" & vbCrLf
    Response.Write "         Content2 = arrContent(1)" & vbCrLf
    Response.Write "         Content5 = EditorFullScreen.HtmlEdit.document.Body.innerHTML" & vbCrLf
    Response.Write "         regEx.Pattern = ""\<IMG(.[^\<]*)\}['|""""]\>""" & vbCrLf
    Response.Write "         Set Matches = regEx.Execute(Content5)" & vbCrLf
    Response.Write "         For Each Match In Matches" & vbCrLf
    Response.Write "             regEx.Pattern = ""\{\$(.*?)\}""" & vbCrLf
    Response.Write "             Set strTemp = regEx.Execute(Match.Value)" & vbCrLf
    Response.Write "             For Each Match2 In strTemp" & vbCrLf
    Response.Write "                strTemp2 = Replace(Match2.Value, ""?"", """""""")" & vbCrLf
    Response.Write "                Content5 = Replace(Content5, Match.Value, ""<!--""&strTemp2&""-->"")" & vbCrLf
    Response.Write "             Next" & vbCrLf
    Response.Write "         Next" & vbCrLf
    Response.Write "         regEx.Pattern = ""\<IMG(.[^\<]*)\$\>""" & vbCrLf
    Response.Write "         Set Matches = regEx.Execute(Content5)" & vbCrLf
    Response.Write "         For Each Match In Matches" & vbCrLf
    Response.Write "         regEx.Pattern = ""\#(.*?)\#""" & vbCrLf
    Response.Write "         Set strTemp = regEx.Execute(Match.Value)" & vbCrLf
    Response.Write "            For Each Match2 In strTemp" & vbCrLf
    Response.Write "               strTemp2 = Replace(strTemp2, ""#"", """")" & vbCrLf
    Response.Write "               strTemp2 = Replace(strTemp2, ""[!"", ""<"")" & vbCrLf
    Response.Write "               strTemp2 = Replace(strTemp2, ""!]"", "">"")" & vbCrLf
    Response.Write "               Content5 = Replace(Content5, Match.Value, strTemp2)" & vbCrLf
    Response.Write "            Next" & vbCrLf
    Response.Write "         Next" & vbCrLf
    Response.Write "        Content5=Replace(Content5, ""http://" & TrueSiteUrl & InstallDir & """,""{$InstallDir}"")" & vbCrLf
    Response.Write "        Content5=Replace(Content5, ""http://" & LCase(TrueSiteUrl) & LCase(InstallDir) & """,""{$InstallDir}"")" & vbCrLf
    
    If Request.QueryString("num") = 1 Then
        Response.Write "opener.editor.HtmlEdit.document.body.innerHTML=Resumeblank(EditorFullScreen.getHTML())" & vbCrLf
        Response.Write "opener.document.form1.Content.value=Content1& vbCrLf & Resumeblank(Content5) & vbCrLf &""</body>""& vbCrLf &""</html>""" & vbCrLf
        Response.Write "opener.editor.showBorders()" & vbCrLf
        Response.Write "opener.editor.showBorders()" & vbCrLf
    Else
        Response.Write "opener.editor2.HtmlEdit.document.body.innerHTML=Resumeblank(EditorFullScreen.getHTML())" & vbCrLf
        Response.Write "opener.document.form1.Content2.value=Content1& vbCrLf & Resumeblank(Content5) & vbCrLf &""</body>""& vbCrLf &""</html>""" & vbCrLf
        Response.Write "opener.editor2.showBorders()" & vbCrLf
        Response.Write "opener.editor2.showBorders()" & vbCrLf
        
    End If

    Response.Write "    Set regEx = Nothing" & vbCrLf
    Response.Write "End function" & vbCrLf
    Response.Write "function setstatus()" & vbCrLf '����������editor.asp���ܵ���
    Response.Write "End function" & vbCrLf
    Response.Write "function setContent(zhi,TemplateType)" & vbCrLf
    Response.Write "End function" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "   setTimeout(""EditorFullScreen.showBorders()"",2000);" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "</BODY>" & vbCrLf
    Response.Write "</HTML>" & vbCrLf
   
End Sub

'=================================================
'��������ShiftCharacter
'��  �ã��滻��ǩΪͼƬ��ʾ
'��  ����Ҫ�滻������    Content
'=================================================
Function ShiftCharacter(ByVal Content)

    Dim strTemp, StrBody, arrContent, ContentHead, arrMatch, strMatch, i, TrueSiteUrl
    
    '�滻�ļ���ע�⺯�������������ʾ����
    
    TrueSiteUrl = Trim(Request.ServerVariables("HTTP_HOST"))
    
    Content = Replace(Content, "<!--{$", "{$")
    Content = Replace(Content, "}-->", "}")


    regEx.Pattern = "(\<body\>)"
    Content = regEx.Replace(Content, "<body>")

    Set Matches = regEx.Execute(Content)
    For Each Match In Matches
        StrBody = Match.value
    Next

    If InStr(Content, "<body>") = 0 Then
        regEx.Pattern = "(\<body)(.[^\<]*)(\>)"
        Set Matches = regEx.Execute(Content)
        For Each Match In Matches
            StrBody = Match.value
        Next
    Else
        StrBody = "<body>"
    End If
    
    arrContent = Split(Content, StrBody)
    
    If UBound(arrContent) = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����ص�ģ��û�а��� <body> ����û�и�body �������ʹ��ҳ���ѿ�,�����ٸ��� <body> ��</li>"
        Exit Function
    End If
    
    ContentHead = arrContent(0) & StrBody
    Content = arrContent(1)
    
    'ͼƬ�滻JS
    regEx.Pattern = "(\<Script)([\s\S]*?)(\<\/Script\>)"
    Set Matches = regEx.Execute(Content)

    For Each Match In Matches
        strTemp = Replace(Match.value, "<", "[!")
        strTemp = Replace(strTemp, ">", "!]")
        strTemp = Replace(strTemp, "'", """")
        strTemp = "<IMG alt='#" & strTemp & "#' src=""" & InstallDir & "editor/images/jscript.gif"" border=0 $>"
        Content = Replace(Content, Match.value, strTemp)
    Next
    
    'ͼƬ�滻������ǩ
    regEx.Pattern = "(\{\$GetPicArticle|\{\$GetArticleList|\{\$GetSlidePicArticle|\{\$GetPicSoft|\{\$GetSoftList|\{\$GetSlidePicSoft|\{\$GetPicPhoto|\{\$GetPhotoList|\{\$GetSlidePicPhoto|\{\$GetPicProduct|\{\$GetProductList|\{\$GetSlidePicProduct|\{\$GetPositionList|\{\$GetSearchResult)\((.*?)\)\}"
    Content = regEx.Replace(Content, "<IMG src=""" & InstallDir & "editor/images/label.gif"" border=0 zzz='$1($2)}'>")
      
    Content = ContentHead & vbCrLf & Content
    '�滻�ļ���ǩ ת��Ϊ css�ļ�
    Content = Replace(Content, "{$Skin_CSS}", "<link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>")
    Content = Replace(Content, "{$MenuJS}", "<script language='JavaScript' type='text/JavaScript' src='" & InstallDir & "js/menu.js'></script>")
    '�滻�ļ��е�·��
    Content = Replace(Content, "[InstallDir_ChannelDir]", "http://" & TrueSiteUrl & InstallDir & ChannelDir & "/")
    Content = Replace(Content, "{$InstallDir}", "http://" & TrueSiteUrl & InstallDir)
    
    ShiftCharacter = Content
    
End Function

'=================================================
'��������ShiftCharacterSave
'��  �ã��滻��ǩΪͼƬ��ʾ
'��  ����Ҫ�滻������    Content
'=================================================
Function ShiftCharacterSave(Content)

    Dim NullBody
    Dim strTemp, strTemp2, Match2, strSiteUrl, strPhotoJs
    
    '�����Ե�ַת��Ϊ��Ե�ַ
    strSiteUrl = Request.ServerVariables("HTTP_REFERER")
    strSiteUrl = Left(strSiteUrl, InStrRev(strSiteUrl, "/") - 1)
    strSiteUrl = Left(strSiteUrl, InStrRev(strSiteUrl, "/"))
     
    'ʹ�������жϴ��࣬С���ļ����Ƿ��� <body> ����Щ����ɾ���ı����е�<body>
    regEx.Pattern = "(\<body\>)"
    Set Matches = regEx.Execute(Content)
    For Each Match In Matches
        NullBody = Match.value
    Next

    If NullBody = "" Then
        regEx.Pattern = "(\<body)(.[^\<]*)(\>)"
        Set Matches = regEx.Execute(Content)
        For Each Match In Matches
            NullBody = Match.value
        Next
    Else
        NullBody = "<body>"
    End If
    
    If NullBody = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����ص�ģ��û�а���  &lt;<font color=blue>body</font>&gt;  ������ҳ���ǲ�����ģ�</li>"
        ShiftCharacterSave = False
        Exit Function
    End If
      
    regEx.Pattern = "\<IMG(.[^\<]*)\}['|""]>"
    Set Matches = regEx.Execute(Content)

    For Each Match In Matches
        regEx.Pattern = "\{\$(.*?)\}"
        Set strTemp = regEx.Execute(Match.value)

        For Each Match2 In strTemp
            strTemp2 = Replace(Match2.value, "?", """")
            Content = Replace(Content, Match.value, "<!--" & strTemp2 & "-->")
        Next
    Next
    
    '����ͼƬJS��ǩ����
    regEx.Pattern = "\<IMG(.[^\<]*)\$\>"
    Set Matches = regEx.Execute(Content)

    For Each Match In Matches
        regEx.Pattern = "\#(.*?)\#"
        Set strTemp = regEx.Execute(Match.value)

        For Each Match2 In strTemp
            strTemp2 = Replace(Match2.value, "?", "?")
            strTemp2 = Replace(strTemp2, "&amp;", "&")
            strTemp2 = Replace(strTemp2, "&13;&10;", "vbCrLf")
            strTemp2 = Replace(strTemp2, "&9;", "vbTab")
            strTemp2 = Replace(strTemp2, "[!", "<")
            strTemp2 = Replace(strTemp2, "!]", ">")
            Content = Replace(Content, Match.value, strTemp2)
        Next
    Next

    '����༭������
    Content = Replace(Content, "{$InstallDir}{$rsClass_ClassUrl}", "{$rsClass_ClassUrl}")
    Content = Replace(Content, "{$InstallDir}{$ArticleUrl}", "{$ArticleUrl}")
    Content = Replace(Content, "{$InstallDir}{$SoftUrl}", "{$SoftUrl}")
    Content = Replace(Content, "{$InstallDir}{$PhotoUrl}", "{$PhotoUrl}")
    Content = Replace(Content, "{$InstallDir}{$ProductUrl}", "{$ProductUrl}")

    '����༭������Ϊ��ǩֵΪ������
    Content = Replace(Content, "{$InstallDir}", "[$InstallDir]")
    regEx.Pattern = "(\s)+(value|title|src|href)(\s)*\=(\s)*\{\$(.[^\<\{]*)\}"
    Set Matches = regEx.Execute(Content)

    For Each Match In Matches
        strTemp = Replace(Trim(Match.value), "{$", """{$")
        strTemp = Replace(strTemp, "}", "}""")
        Content = Replace(Content, Match.value, " " & strTemp)
    Next

    Content = Replace(Content, "[$InstallDir]", "{$InstallDir}")

    '�������ҳ�û�ɾ��ͼƬjs ����
    strPhotoJs = "<script language=""JavaScript"">" & vbCrLf
    strPhotoJs = strPhotoJs & "<!--" & vbCrLf
    strPhotoJs = strPhotoJs & "//�ı�ͼƬ��С" & vbCrLf
    strPhotoJs = strPhotoJs & "function resizepic(thispic)" & vbCrLf
    strPhotoJs = strPhotoJs & "{" & vbCrLf
    'strPhotoJs = strPhotoJs & "if(thispic.width>700) thispic.width=700;" & vbCrLf
    strPhotoJs = strPhotoJs & "  return true;" & vbCrLf
    strPhotoJs = strPhotoJs & "}" & vbCrLf
    strPhotoJs = strPhotoJs & "//�޼�����ͼƬ��С" & vbCrLf
    strPhotoJs = strPhotoJs & "function bbimg(o)" & vbCrLf
    strPhotoJs = strPhotoJs & "{" & vbCrLf
    'strPhotoJs = strPhotoJs & "  var zoom=parseInt(o.style.zoom, 10)||100;" & vbCrLf
    'strPhotoJs = strPhotoJs & "  zoom+=event.wheelDelta/12;" & vbCrLf
    'strPhotoJs = strPhotoJs & "  if (zoom>0) o.style.zoom=zoom+'%';" & vbCrLf
    strPhotoJs = strPhotoJs & "  return true;" & vbCrLf
    strPhotoJs = strPhotoJs & "}" & vbCrLf
    strPhotoJs = strPhotoJs & "-->" & vbCrLf
    strPhotoJs = strPhotoJs & "</script>" & vbCrLf
    strPhotoJs = strPhotoJs & "</head>" & vbCrLf

    If TemplateType = 3 Then
        If InStr(Content, "resizepic(thispic)") <= 0 Or InStr(Content, "bbimg(o)") <= 0 Then
            Content = Replace(Content, "</head>", strPhotoJs)
        End If
    End If
    
    ShiftCharacterSave = Content
End Function


'**************************************************
'��������BatchReplaceString
'��  �ã������滻������
'��  ����TemplateContent ----ģ������
'��  ����TemplateReplaceStart ----���Ҫ�滻�Ŀ�ͷ����
'��  ����TemplateReplaceEnd ----���Ҫ�滻�Ľ�������
'��  ����TemplateReplaceResult ----Ҫ�滻�Ĵ���
'��  ����TemplateName ----ģ������
'����ֵ��True  ----�Ѵ���
'**************************************************
Function BatchReplaceString(TemplateContent, _
                                    TemplateReplaceStart, _
                                    TemplateReplaceEnd, _
                                    TemplateReplaceResult, _
                                    TemplateName)

    If InStr(TemplateContent, TemplateReplaceStart) = 0 Or InStr(TemplateContent, TemplateReplaceEnd) = 0 Then
        BatchReplaceString = TemplateContent
        Response.Write "<br>&nbsp;&nbsp;" & TemplateName & "..<font color='#FF0000'>ģ���滻��ʼ���� �� ����������,�����滻��</font>"
        Exit Function
    End If

    If GetBody(TemplateContent, TemplateReplaceStart, TemplateReplaceEnd, True, True) = "" Then
        BatchReplaceString = TemplateContent
        Response.Write "<br>&nbsp;&nbsp;" & TemplateName & "..<font color='#FF0000'>ģ���滻��ʼ���� �� ������Ѱλ�ò���,�����滻��</font>"
        Exit Function
    End If

    BatchReplaceString = Replace(TemplateContent, GetBody(TemplateContent, TemplateReplaceStart, TemplateReplaceEnd, True, True), TemplateReplaceResult)
    Response.Write "<br>&nbsp;&nbsp;" & TemplateName & "..<font color='#009900'>ģ���滻�ɹ���</font>"
End Function

'=================================================
'��������Resumeblank
'��  �ã�����ͻ��� html
'=================================================
Sub Resumeblank()

    Response.Write "Function  Resumeblank(byval Content)" & vbCrLf
    Response.Write " Dim strHtml,strHtml2,Num,Numtemp,Strtemp" & vbCrLf
    Response.Write "   strHtml=Replace(Content, ""<DIV"", ""<div"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</DIV>"", vbCrLf & ""</div>""& vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<DD>"", vbCrLf & ""<dd>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<DT>"", vbCrLf & ""<dt>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<DL>"", vbCrLf & ""<dl>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</DD>"", vbCrLf & ""</dd>""& vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</DT>"", vbCrLf & ""</dt>""& vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</DL>"", vbCrLf & ""</dl>""& vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<TABLE"", ""<table"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</TABLE>"", vbCrLf & ""</table>""& vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<TBODY>"", """")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</TBODY>"","""" & vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<TR"", ""<tr"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</TR>"", vbCrLf & ""</tr>""& vbCrLf)" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<TD"", ""<td"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</TD>"", ""</td>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<!--"", vbCrLf & ""<!--"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<SELECT"",vbCrLf & ""<Select"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</SELECT>"",vbCrLf & ""</Select>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<OPTION"",vbCrLf & ""  <Option"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""</OPTION>"",""</Option>"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<INPUT"",vbCrLf & ""  <Input"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""<script"",vbCrLf & ""<script"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""&amp;"",""&"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""{$--"",vbCrLf & ""<!--$"")" & vbCrLf
    Response.Write "   strHtml=Replace(strHtml, ""--}"",""$-->"")" & vbCrLf
    Response.Write "   arrContent = Split(strHtml,vbCrLf)" & vbCrLf
    Response.Write "    For i = 0 To UBound(arrContent)" & vbCrLf
    Response.Write "        Numtemp=false" & vbCrLf
    Response.Write "        if Instr(arrContent(i),""<table"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""<table"" and Strtemp <>""</table>"" then" & vbCrLf
    Response.Write "              Num=Num+2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""<table""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""<tr"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""<tr"" and Strtemp<>""</tr>"" then" & vbCrLf
    Response.Write "              Num=Num+2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""<tr""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""<td"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""<td"" and Strtemp<>""</td>"" then" & vbCrLf
    Response.Write "              Num=Num+2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""<td""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""</table>"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""</table>"" and Strtemp<>""<table"" then" & vbCrLf
    Response.Write "              Num=Num-2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""</table>""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""</tr>"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""</tr>"" and Strtemp<>""<tr"" then" & vbCrLf
    Response.Write "              Num=Num-2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""</tr>""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""</td>"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "            if Strtemp<>""</td>"" and Strtemp<>""<td"" then" & vbCrLf
    Response.Write "              Num=Num-2" & vbCrLf
    Response.Write "            End if " & vbCrLf
    Response.Write "            Strtemp=""</td>""" & vbCrLf
    Response.Write "        elseif Instr(arrContent(i),""<!--"")>0 then" & vbCrLf
    Response.Write "            Numtemp=True" & vbCrLf
    Response.Write "        End if" & vbCrLf
    Response.Write "        if Num< 0 then Num = 0" & vbCrLf
    Response.Write "        if trim(arrContent(i))<>"""" then" & vbCrLf
    Response.Write "            if i=0 then" & vbCrLf
    Response.Write "                strHtml2= string(Num,"" "") & arrContent(i) " & vbCrLf
    Response.Write "            elseif Numtemp=True then" & vbCrLf
    Response.Write "                strHtml2= strHtml2 & vbCrLf & string(Num,"" "") & arrContent(i) " & vbCrLf
    Response.Write "            else" & vbCrLf
    Response.Write "                strHtml2= strHtml2 & vbCrLf & arrContent(i) " & vbCrLf
    Response.Write "            end if" & vbCrLf
    Response.Write "        end if" & vbCrLf
    Response.Write "    Next" & vbCrLf
    Response.Write "    Resumeblank=strHtml2" & vbCrLf
    Response.Write "End function" & vbCrLf
End Sub

'**************************************************
'��������IsFontChecked
'��  �ã��Ƿ���Ĭ��,Ĭ����ʾ��ɫ
'��  ����Compare1-----�Ƚ�ֵ1
'��  ����Compare2-----�Ƚ�ֵ2
'����ֵ���滻���ַ���
'**************************************************
Function IsFontChecked(ByVal Compare1, ByVal Compare2)
    If Compare1 = Compare2 Then
        IsFontChecked = " color='red'"
    Else
        IsFontChecked = ""
    End If
End Function

'**************************************************
'��������IsFontChecked2
'��  �ã��Ƿ���Ĭ��,Ĭ����ʾ��ɫ
'��  ����Compare1-----�Ƚ�ֵ1
'��  ����Compare2-----�Ƚ�ֵ2
'����ֵ���滻���ַ���
'**************************************************
Function IsFontChecked2(ByVal Compare1, ByVal Compare2, ByVal IsOnlinePayment1, ByVal IsOnlinePayment2)
    If Compare1 = Compare2 Then
        If IsOnlinePayment1 = IsOnlinePayment2 Then
            IsFontChecked2 = " color='red'"
        End If
    Else
        IsFontChecked2 = ""
    End If
End Function


'**************************************************
'��������IsRadioChecked
'��  �ã���ѡ,��ѡĬ��
'��  ����Compare1-----�Ƚ�ֵ1
'��  ����Compare2-----�Ƚ�ֵ2
'����ֵ���滻���ַ���
'**************************************************
Function IsRadioChecked(ByVal Compare1, _
                                ByVal Compare2)

    If Compare1 = Compare2 Then
        IsRadioChecked = " checked"
    Else
        IsRadioChecked = ""
    End If

End Function

'=================================================
'��������GetProject_Option
'��  �ã�������������
'��  ����iProjectName  ----��������
'=================================================
Function GetProject_Option(iProjectName)
    Dim sqlProject, rsProject, strProject

    sqlProject = "select * from PE_TemplateProject"
    Set rsProject = Conn.Execute(sqlProject)

    If rsProject.BOF And rsProject.EOF Then
    Else

        Do While Not rsProject.EOF
            strProject = strProject & "<option value='" & rsProject("TemplateProjectName") & "'"

            If rsProject("TemplateProjectName") = iProjectName Then
                strProject = strProject & " selected"
            End If

            strProject = strProject & ">" & rsProject("TemplateProjectName")

            If rsProject("IsDefault") = True Then
                strProject = strProject & "��Ĭ�ϣ�"
            End If

            strProject = strProject & "</option>"
            rsProject.MoveNext
        Loop

    End If

    rsProject.Close
    Set rsProject = Nothing
    GetProject_Option = strProject
End Function

'=================================================
'��������GetAllModule
'��  �ã���ʾ�����˵��ĵ���ģ��
'=================================================
Sub GetAllModule(SystemType, ModuleType)
    Response.Write "<option " & OptionValue(CLng(ModuleType), 0) & ">ͨ��ģ��</option>"
    Response.Write "<option " & OptionValue(CLng(ModuleType), 1) & ">����ģ��</option>" & vbCrLf
    Response.Write "<option " & OptionValue(CLng(ModuleType), 2) & ">����ģ��</option>" & vbCrLf
    Response.Write "<option " & OptionValue(CLng(ModuleType), 3) & ">ͼƬģ��</option>" & vbCrLf
    
    If SystemType = "4.03" Then
        Response.Write "<option " & OptionValue(CLng(ModuleType), -2) & ">�û��Զ���ģ��</option>" & vbCrLf
    Else
        Response.Write "<option " & OptionValue(CLng(ModuleType), 4) & ">����ģ��</option>" & vbCrLf
        Response.Write "<option " & OptionValue(CLng(ModuleType), 5) & ">�̳�ģ��</option>" & vbCrLf
        If FoundInArr(AllModules, "Supply", ",") Then
            Response.Write "<option " & OptionValue(CLng(ModuleType), 6) & ">����ģ��</option>" & vbCrLf
        End If
        If FoundInArr(AllModules, "House", ",") Then
            Response.Write "<option " & OptionValue(CLng(ModuleType), 7) & ">����ģ��</option>" & vbCrLf
        End If
        If FoundInArr(AllModules, "Job", ",") Then
            Response.Write "<option " & OptionValue(CLng(ModuleType), 8) & ">��Ƹģ��</option>" & vbCrLf
        End If
    End If

    Response.Write "<option " & OptionValue(CLng(ModuleType), 999999) & ">����ģ��</option>" & vbCrLf
End Sub

'=================================================
'��������GetTemplate_Option
'��  �ã�Ƶ��ģ����������ѡ��
'��  ����CurrentTemplateType --- �����ģ��ֵ
'=================================================
Function GetTemplate_Option(CurrentTemplateType)
    Dim strTemp
    If ChannelID > 0 Then
        If ModuleType = 4 Then
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 1) & ">������ҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 3) & ">���Է���ģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 4) & ">���Իظ�ģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 5) & ">��������ҳģ��</option>" & vbCrLf
        Else
            Select Case ModuleType
            Case 7 '����ģ��
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 1) & ">Ƶ����ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 2) & ">��Ŀģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 7) & ">�Ƽ�ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 8) & ">����ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 30) & ">��������ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 31) & ">��������ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 32) & ">������ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 33) & ">��������ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 34) & ">��������ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 5) & ">����ҳģ��</option>" & vbCrLf
            Case 8 '��Ƹģ��
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 1) & ">Ƶ����ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 5) & ">ְλ����ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 3) & ">ְλ����ҳģ��</option>" & vbCrLf
            Case Else
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 1) & ">Ƶ����ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 2) & ">��Ŀҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 3) & ">����ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 4) & ">ר��ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 22) & ">ר���б�ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 5) & ">����ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 6) & ">����" & ChannelShortName & "ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 7) & ">�Ƽ�" & ChannelShortName & "ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 8) & ">����" & ChannelShortName & "ҳģ��</option>" & vbCrLf
                strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 16) & ">����" & ChannelShortName & "ҳģ��</option>" & vbCrLf

                If ModuleType = 1 Or ModuleType = 2 Or ModuleType = 3 Then
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 23) & ">�������" & ChannelShortName & "ҳģ��</option>" & vbCrLf
                End If

                If ModuleType = 1 Then
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 17) & ">��ӡҳģ��</option>" & vbCrLf
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 20) & ">���ߺ���ҳģ��</option>" & vbCrLf
                ElseIf ModuleType = 5 Then
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 9) & ">���ﳵģ��</option>" & vbCrLf
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 10) & ">����̨ģ��</option>" & vbCrLf
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 11) & ">����Ԥ��ҳģ��</option>" & vbCrLf
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 12) & ">�����ɹ�ҳģ��</option>" & vbCrLf
                    'strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 13) & ">����֧����һ��ģ��</option>" & vbCrLf
                    'strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 14) & ">����֧���ڶ���ģ��</option>" & vbCrLf
                    'strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 15) & ">����֧��������ģ��</option>" & vbCrLf
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 19) & ">�ؼ���Ʒҳģ��</option>" & vbCrLf
                    strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 21) & ">�̳ǰ���ҳģ��</option>" & vbCrLf
                End If
            End Select
        End If

    Else

        If TempType = 1 Then
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 102) & ">��Ա����ͨ��ģ��</option>" & vbCrLf		
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 8) & ">��Ա��Ϣҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 9) & ">��Ա�б�ҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 18) & ">��Աע��ҳģ�壨���Э�飩</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 19) & ">��Աע��ҳģ�壨������Ŀ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 20) & ">��Աע��ҳģ�壨ѡ����Ŀ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 21) & ">��Աע��ҳģ�壨ע������</option>" & vbCrLf
        Else
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 1) & ">��վ��ҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 3) & ">��վ����ҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 4) & ">��վ����ģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 22) & ">��վ�����б�ҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 5) & ">��վ����ҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 6) & ">��վ����ҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 7) & ">��Ȩ����ҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 29) & ">ȫվר���б�ҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 30) & ">ȫվר��ҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 10) & ">������ʾҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 11) & ">�����б�ҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 12) & ">��Դ��ʾҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 13) & ">��Դ�б�ҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 103) & ">����Ͷ��ģ��</option>" & vbCrLf
			
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 14) & ">������ʾҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 15) & ">�����б�ҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 16) & ">Ʒ����ʾҳģ��</option>" & vbCrLf
            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 17) & ">Ʒ���б�ҳģ��</option>" & vbCrLf

            strTemp = strTemp & "<option " & OptionValue(CurrentTemplateType, 101) & ">�Զ����б�ģ��</option>" & vbCrLf
        End If
    End If

    GetTemplate_Option = strTemp
End Function

'=================================================
'��������GetTemplateTypeName
'��  �ã���ʾ��ǰƵ����ģ������
'��  ����iTemplateType --- �����ģ��ֵ
'=================================================
Function GetTemplateTypeName(iTemplateType, _
                                     ChannelID)

    If ChannelID > 0 Then
        If ModuleType = 4 Then

            Select Case iTemplateType

                Case 1
                    GetTemplateTypeName = "������ҳģ��"

                Case 3
                    GetTemplateTypeName = "���Է���ģ��"

                Case 4
                    GetTemplateTypeName = "���Իظ�ģ��"

                Case 5
                    GetTemplateTypeName = "��������ҳģ��"
            End Select

        Else

            Select Case iTemplateType
            Case 1
                GetTemplateTypeName = "Ƶ����ҳģ��"
            Case 2
                GetTemplateTypeName = "Ƶ����Ŀģ��"
            Case 3
                GetTemplateTypeName = "Ƶ������ҳģ��"
            Case 4
                GetTemplateTypeName = "Ƶ��ר��ҳģ��"
            Case 5
                GetTemplateTypeName = "Ƶ������ҳģ��"
            Case 6
                GetTemplateTypeName = "����" & ChannelShortName & "ҳģ��"
            Case 7
                GetTemplateTypeName = "�Ƽ�" & ChannelShortName & "ҳģ��"
            Case 8
                GetTemplateTypeName = "�ȵ�" & ChannelShortName & "ҳģ��"
            Case 16
                GetTemplateTypeName = "����" & ChannelShortName & "ҳģ��"
            Case 9
                GetTemplateTypeName = "���ﳵģ��"
            Case 10
                GetTemplateTypeName = "����̨ģ��"
            Case 11
                GetTemplateTypeName = "Ԥ������ģ��"
            Case 12
                GetTemplateTypeName = "�����ɹ�ҳģ��"
            'Case 13
            '   GetTemplateTypeName = "����֧����һ��ģ��"
            'Case 14
            '   GetTemplateTypeName = "����֧���ڶ���ģ��"
            'Case 15
            '   GetTemplateTypeName = "����֧��������ģ��"
            Case 17
                GetTemplateTypeName = "��ӡģ��"
            Case 101
                GetTemplateTypeName = "�Զ����б�ģ��"
            Case 19
                GetTemplateTypeName = "�ؼ���Ʒҳģ��"
            Case 20
                GetTemplateTypeName = "���ߺ���ҳģ��"
            Case 21
                GetTemplateTypeName = "�̳ǰ���ҳģ��"
            Case 22
                GetTemplateTypeName = "Ƶ��ר���б�ҳģ��"
            Case 23
                GetTemplateTypeName = "�������" & ChannelShortName & "ҳģ��"
            '���ӷ���ģ������ģ������
            Case 30
                GetTemplateTypeName = "��������ҳģ��"
            Case 31
                GetTemplateTypeName = "��������ҳģ��"
            Case 32
                GetTemplateTypeName = "������ҳģ��"
            Case 33
                GetTemplateTypeName = "��������ҳģ��"
            Case 34
                GetTemplateTypeName = "��������ҳģ��"
            '***************End********************
            End Select
        End If
    Else
        Select Case iTemplateType
        Case 1
            GetTemplateTypeName = "��վ��ҳģ��"
        Case 3
            GetTemplateTypeName = "��վ����ҳģ��"
        Case 4
            GetTemplateTypeName = "��վ����ҳģ��"
        Case 5
            GetTemplateTypeName = "��������ҳģ��"
        Case 6
            GetTemplateTypeName = "��վ����ҳģ��"
        Case 7
            GetTemplateTypeName = "��Ȩ����ҳģ��"
        Case 8
            GetTemplateTypeName = "��Ա��Ϣҳģ��"
        Case 102	
            GetTemplateTypeName = "��Ա����ͨ��ģ��"			
        Case 9
            GetTemplateTypeName = "��Ա�б�ҳģ��"
        Case 10
            GetTemplateTypeName = "������ʾҳģ��"
        Case 11
            GetTemplateTypeName = "�����б�ҳģ��"
        Case 12
            GetTemplateTypeName = "��Դ��ʾҳģ��"
        Case 13
            If IsOnlinePayment = 0 Then
                GetTemplateTypeName = "��Դ�б�ҳģ��"
            'Else
            '   GetTemplateTypeName = "����֧����һ��ģ��"
            End If
        Case 103
            GetTemplateTypeName = "����Ͷ��ģ��"			
        Case 14
            If IsOnlinePayment = 0 Then
                GetTemplateTypeName = "������ʾҳģ��"
            'Else
            '   GetTemplateTypeName = "����֧���ڶ���ģ��"
            End If
        Case 15
            If IsOnlinePayment = 0 Then
                GetTemplateTypeName = "�����б�ҳģ��"
            'Else
            '   GetTemplateTypeName = "����֧��������ģ��"
            End If
        Case 16
            GetTemplateTypeName = "Ʒ����ʾҳģ��"
        Case 17
            GetTemplateTypeName = "Ʒ���б�ҳģ��"
        Case 101
            GetTemplateTypeName = "�Զ����б�ģ��"
        Case 18
            GetTemplateTypeName = "��Աע��ҳģ�壨���Э�飩"
        Case 19
            GetTemplateTypeName = "��Աע��ҳģ�壨������Ŀ��"
        Case 20
            GetTemplateTypeName = "��Աע��ҳģ�壨ѡ����Ŀ��"
        Case 21
            GetTemplateTypeName = "��Աע��ҳģ�壨ע������"
        Case 22
            GetTemplateTypeName = "�����б�ҳģ��"
        Case 29
            GetTemplateTypeName = "ȫվר���б�ҳģ��"
        Case 30
            GetTemplateTypeName = "ȫվר��ҳģ��"
        End Select

    End If

    If iTemplateType = 0 Then
        GetTemplateTypeName = "��ǰ��������ģ��"
    End If

End Function
%>
