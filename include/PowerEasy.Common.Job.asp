<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Function GetEducation(Education)
    Dim strEducation
    Select Case Education
    Case 0
         strEducation = "û��ѡ��ѧ��"
    Case 1
         strEducation = "Сѧ"
    Case 2
         strEducation = "����"
    Case 3
         strEducation = "����"
    Case 4
         strEducation = "��ר"
    Case 5
         strEducation = "��ר"
    Case 6
         strEducation = "����"
    Case 7
         strEducation = "˶ʿ"
    Case 8
         strEducation = "��ʿ"
    Case 9
         strEducation = "��ʿ��"
    Case 10
         strEducation = "����"
    Case Else
         strEducation = "����"
    End Select
    GetEducation = strEducation
End Function



Function GetMarital(Marital)
    Dim strMarital
    Select Case Marital
    Case 0
         strMarital = "ûѡ�����״��"
    Case 1
         strMarital = "δ��"
    Case 2
         strMarital = "�ѻ�"
    Case 3
         strMarital = "����"
    Case 4
         strMarital = "����"
    Case Else
         strMarital = "����"
    End Select
    GetMarital = strMarital
End Function


Function GetSex(Sex)
    Dim strSex
    Select Case Sex
    Case 0
         strSex = "û��ѡ���Ա�"
    Case 1
         strSex = "��"
    Case 2
         strSex = "Ů"
    Case 3
         strSex = "����"
    Case Else
         strSex = "����"
    End Select
    GetSex = strSex
End Function

Function GetForeignLanguageKind(ForeignLanguageKind)
    Dim strForeignLanguageKind
    Select Case ForeignLanguageKind
    Case 0
         strForeignLanguageKind = "ûѡ����������"
    Case 1
         strForeignLanguageKind = "Ӣ��"
    Case 2
         strForeignLanguageKind = "����"
    Case 3
         strForeignLanguageKind = "����"
    Case 4
         strForeignLanguageKind = "����"
    Case 5
         strForeignLanguageKind = "��������"
    Case Else
         strForeignLanguageKind = "Ӣ��"
    End Select
    GetForeignLanguageKind = strForeignLanguageKind
End Function

Function GetCheckStatus(CheckStatus)
    Dim strCheckStatus
    Select Case CheckStatus
    Case 0
         strCheckStatus = "��δ�鿴"
    Case 1
         strCheckStatus = "<font color='#2C9EC9'>ѡΪ����</font>"
    Case 2
         strCheckStatus = "<font color='red'>ͨ������һ</font>"
    Case 3
         strCheckStatus = "<font color='#B51523'>ͨ�����Զ�</font>"
    Case 4
         strCheckStatus = "<font color='blue'>�Ѿ�¼��</font>"
    Case Else
         strCheckStatus = "<font color='#00000F'>�Ѿ��յ�</font>"

    End Select
    GetCheckStatus = strCheckStatus
End Function

%>
