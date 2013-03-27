<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_Content.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<!--#include file="../Include/PowerEasy.SendMail.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 3   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

Dim HtmlDir, PurviewChecked, AddType
Dim ManageType, Status, MyStatus, arrStatus
Dim ClassID, SpecialID, OnTop, IsElite, IsHot, Created

Dim tClass, ClassName, RootID, ParentID, Depth, ParentPath, Child, arrChildID, ParentDir, ClassDir, ClassPurview

Dim SoftID
Dim VoteID

Dim arrFields_Options, arrSoftType, arrSoftLanguage, arrCopyrightType, arrOperatingSystem
    
arrFields_Options = Split(",,,", ",")
arrSoftType = ""
arrSoftLanguage = ""
arrCopyrightType = ""
arrOperatingSystem = ""
If Fields_Options & "" <> "" Then
    arrFields_Options = Split(Fields_Options, "$$$")
    If UBound(arrFields_Options) = 3 Then
        arrSoftType = Split(arrFields_Options(0), vbCrLf)
        arrSoftLanguage = Split(arrFields_Options(1), vbCrLf)
        arrCopyrightType = Split(arrFields_Options(2), vbCrLf)
        arrOperatingSystem = Split(arrFields_Options(3), vbCrLf)
    End If
End If

If AdminPurview = 1 Then
    MyStatus = 3
Else
    Select Case CheckLevel
    Case 0, 1
        MyStatus = 3
    Case 2
        If AdminPurview_Channel <= 2 Then
            MyStatus = 3
        Else
            MyStatus = 2
        End If
    Case 3
        MyStatus = 4 - AdminPurview_Channel
    End Select
End If
arrStatus = Array("�����", "һ��ͨ��", "����ͨ��", "����ͨ��")

HtmlDir = InstallDir & ChannelDir

If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))

ManageType = Trim(Request("ManageType"))
Status = Trim(Request("Status"))
Created = Trim(Request("Created"))
OnTop = Trim(Request("OnTop"))
IsElite = Trim(Request("IsElite"))
IsHot = Trim(Request("IsHot"))
ClassID = PE_CLng(Trim(Request("ClassID")))
SpecialID = PE_CLng(Trim(Request("SpecialID")))
SoftID = Trim(Request("SoftID"))
AddType = Trim(Request("AddType"))

If Action = "" Then
    Action = "Manage"
End If
If Status = "" Then
    Status = 9
Else
    Status = PE_CLng(Status) '���״̬   9�������������-1�����ݸ壬0��������ˣ�1���������
End If

If IsValidID(SoftID) = False Then
    SoftID = ""
End If
If AddType = "" Then
    AddType = 1
Else
    AddType = PE_CLng(AddType)
End If

FileName = "Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=" & Action & "&ManageType=" & ManageType
strFileName = FileName & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&Field=" & strField & "&keyword=" & Keyword
If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))

    
Response.Write "<html><head><title>" & ChannelShortName & "����</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
If Action = "Add" Or Action = "Modify" Then
    Response.Write "<script src=""../JS/prototype.js""></script>" & vbCrLf
    Response.Write "<script src=""../JS/checklogin.js""></script>" & vbCrLf
End If
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
If Action = "Preview" Then
    Call Preview
    Call CloseConn
    Response.End
End If
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
Dim strTitle
strTitle = ChannelName & "����----"
Select Case Action
Case "Add"
    strTitle = strTitle & "���" & ChannelShortName
Case "Modify"
    strTitle = strTitle & "�޸�" & ChannelShortName
Case "Check"
    strTitle = strTitle & "���" & ChannelShortName
Case "SaveAdd", "SaveModify"
    strTitle = strTitle & "����" & ChannelShortName
Case "Move"
    strTitle = strTitle & "�ƶ�" & ChannelShortName
Case "Preview", "Show"
    strTitle = strTitle & "Ԥ��" & ChannelShortName
Case "Batch", "DoBatch"
    strTitle = strTitle & "�����޸�" & ChannelShortName & "����"
Case "MoveToClass"
    strTitle = strTitle & "�����ƶ�" & ChannelShortName
Case "BatchReplace"
    strTitle = strTitle & "�����滻" & ChannelShortName
Case "AddToSpecial"
    strTitle = strTitle & "���" & ChannelShortName & "��ר��"
Case "MoveToSpecial"
    strTitle = strTitle & "�ƶ�" & ChannelShortName & "��ר��"
Case "ShowReplace", "DoReplace"
    strTitle = strTitle & "�����޸�" & ChannelShortName & "��ַ"
Case "Other", "SaveOther"
    strTitle = strTitle & "��������"
Case "DownError", "ModifyDownError", "SaveModifyDownError", "DelDownError", "DelAllDownError"
    strTitle = strTitle & "������Ϣ����"
Case "Manage"
    Select Case ManageType
    Case "Check"
        strTitle = strTitle & ChannelShortName & "���"
    Case "HTML"
        strTitle = strTitle & ChannelShortName & "����"
    Case "Recyclebin"
        strTitle = strTitle & ChannelShortName & "����վ����"
    Case "Special"
        strTitle = strTitle & "ר��" & ChannelShortName & "����"
    Case Else
        strTitle = strTitle & ChannelShortName & "������ҳ"
    End Select
End Select
Call ShowPageTitle(strTitle, 10121)
Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30' ><strong>��������</strong></td><td colspan='5'>"
Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Status=9'>" & ChannelShortName & "������ҳ</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Add&AddType=1&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>���" & ChannelShortName & "</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Add&AddType=3&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>���" & ChannelShortName & "������</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&ManageType=Check&Status=0'>���" & ChannelShortName & "</a>"
If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&ManageType=Special'>ר��" & ChannelShortName & "����</a>"
End If
If AdminPurview = 1 Or AdminPurview_Channel = 1 Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&ManageType=Recyclebin&Status=9'>" & ChannelShortName & "����վ����</a>"
End If
If UseCreateHTML > 0 And ObjInstalled_FSO = True Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&ManageType=HTML&Status=1'><b>����HTML����</b></a>"
End If
Response.Write "</td></tr>" & vbCrLf
If Action = "Manage" Then
    Response.Write "<form name='form3' method='Post' action='" & strFileName & "'><tr class='tdbg'>"
    Response.Write "  <td width='70' height='30' ><strong>" & ChannelShortName & "ѡ�</strong></td><td>"
    If ManageType = "HTML" Then
        Response.Write "<input name='Created' type='radio' onclick='submit();' " & RadioValue(Created, "") & ">����" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='Created' type='radio' onclick='submit();' " & RadioValue(Created, "False") & ">δ���ɵ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp"
        Response.Write "<input name='Created' type='radio' onclick='submit();' " & RadioValue(Created, "True") & ">�����ɵ�" & ChannelShortName
    Else
        Response.Write "<input name='Status' type='radio' onclick='submit();' " & RadioValue(Status, 9) & ">����" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='Status' type='radio' onclick='submit();' " & RadioValue(Status, -1) & ">�ݸ�&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='Status' type='radio' onclick='submit();' " & RadioValue(Status, 0) & ">�����&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='Status' type='radio' onclick='submit();' " & RadioValue(Status, 1) & ">�����&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='Status' type='radio' onclick='submit();' " & RadioValue(Status, -2) & ">�˸�"
        Response.Write "</td><td>"
        Response.Write "<input name='OnTop' type='checkbox' onclick='submit()' " & RadioValue(OnTop, "True") & "> �̶�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='IsElite' type='checkbox' onclick='submit()' " & RadioValue(IsElite, "True") & "> �Ƽ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='IsHot' type='checkbox' onclick='submit()' " & RadioValue(IsHot, "True") & "> ����" & ChannelShortName
    End If
    Response.Write "</td></tr></form>" & vbCrLf
End If
Response.Write "</table>" & vbCrLf

strFileName = strFileName & "&Status=" & Status & "&Created=" & Created & "&OnTop=" & OnTop & "&IsElite=" & IsElite & "&IsHot=" & IsHot

Select Case Action
Case "Add"
    Call Add
Case "Modify", "Check"
    Call Modify
Case "SaveAdd", "SaveModify"
    Call SaveSoft
Case "SetOnTop", "CancelOnTop", "SetElite", "CancelElite", "SetPassed", "CancelPassed", "Reject"
    Call SetProperty
Case "Show"
    Call Show
Case "Del"
    Call Del
Case "ConfirmDel"
    Call ConfirmDel
Case "ClearRecyclebin"
    Call ClearRecyclebin
Case "Restore"
    Call Restore
Case "RestoreAll"
    Call RestoreAll
Case "DelFile"
    Call DelFile
Case "Batch"
    Call Batch
Case "DoBatch"
    Call DoBatch
Case "MoveToClass"
    Call ShowForm_MoveToClass
Case "MoveToSpecial"
    Call ShowForm_MoveToSpecial
Case "AddToSpecial"
    Call ShowForm_AddToSpecial
Case "DoMoveToClass"
    Call DoMoveToClass
Case "DoMoveToSpecial"
    Call DoMoveToSpecial
Case "DoAddToSpecial"
    Call DoAddToSpecial
Case "DelFromSpecial"
    Call DelFromSpecial
Case "ShowReplace"
    Call ShowReplace
Case "DoReplace"
    Call DoReplace
Case "Other"
    Call Other
Case "SaveOther"
    Call SaveOther
Case "Manage"
    Call main
Case "DownError"
    Call DownError
Case "ModifyDownError"
    Call ModifyDownError
Case "SaveModifyDownError"
    Call SaveModifyDownError
Case "DelDownError"
    Call DelDownError
Case "DelAllDownError"
    Call DelAllDownError
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub main()
    If ManageType = "HTML" And UseCreateHTML = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��Ƶ�������˲�����HTML�����Բ��ý������ɹ���</li>"
        Exit Sub
    End If
    Dim rsSoftList, sql, Querysql
    PurviewChecked = False
    If ClassID = 0 Then
        If strField = "" And AdminPurview = 2 And AdminPurview_Channel = 3 And ManageType <> "My" Then
            If ManageType = "Check" Then
                If arrClass_Check = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�Բ�����û���ڴ�Ƶ�����" & ChannelShortName & "��Ȩ�ޣ�</li>"
                    Exit Sub
                End If
                Set tClass = Conn.Execute("select top 1 ClassID,ClassName,RootID,ParentID,Depth,ParentPath,Child,arrChildID,ClassPurview,ParentDir,ClassDir from PE_Class where ChannelID=" & ChannelID & " and ClassID In (" & DelRightComma(arrClass_Check) & ")")
            Else
                If arrClass_Manage = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�Բ�����û���ڴ�Ƶ������" & ChannelShortName & "��Ȩ�ޣ�</li>"
                    Exit Sub
                End If
                Set tClass = Conn.Execute("select top 1 ClassID,ClassName,RootID,ParentID,Depth,ParentPath,Child,arrChildID,ClassPurview,ParentDir,ClassDir from PE_Class where ChannelID=" & ChannelID & " and ClassID In (" & DelRightComma(arrClass_Manage) & ")")
            End If
            If tClass.BOF And tClass.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�Բ�����û���ڴ�Ƶ���Ĺ���Ȩ�ޣ�</li>"
            Else
                ClassID = tClass(0)
                ClassName = tClass(1)
                RootID = tClass(2)
                ParentID = tClass(3)
                Depth = tClass(4)
                ParentPath = tClass(5)
                Child = tClass(6)
                arrChildID = tClass(7)
                PurviewChecked = True
                ClassPurview = tClass(8)
                ParentDir = tClass(9)
                ClassDir = tClass(10)
            End If
        End If
    ElseIf ClassID = -1 Then
        If AdminPurview = 1 Or (AdminPurview = 2 And AdminPurview_Channel < 3) Then PurviewChecked = True
    ElseIf ClassID > 0 Then
        Set tClass = Conn.Execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,arrChildID,ClassPurview,ParentDir,ClassDir from PE_Class where ClassID=" & ClassID)
        If tClass.BOF And tClass.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�Ҳ���ָ������Ŀ</li>"
        Else
            ClassName = tClass(0)
            RootID = tClass(1)
            ParentID = tClass(2)
            Depth = tClass(3)
            ParentPath = tClass(4)
            Child = tClass(5)
            arrChildID = tClass(6)
            ClassPurview = tClass(7)
            ParentDir = tClass(8)
            ClassDir = tClass(9)
        End If
        Set tClass = Nothing
    End If
    If FoundErr = True Then Exit Sub
    If PurviewChecked = False Then
        If ManageType = "Check" Then
            If ParentID > 0 Then
                PurviewChecked = CheckPurview_Class(arrClass_Check, ParentPath & "," & ClassID)
            Else
                PurviewChecked = CheckPurview_Class(arrClass_Check, ClassID)
            End If
        Else
            If ParentID > 0 Then
                PurviewChecked = CheckPurview_Class(arrClass_Manage, ParentPath & "," & ClassID)
            Else
                PurviewChecked = CheckPurview_Class(arrClass_Manage, ClassID)
            End If
        End If
    End If

    Call ShowJS_Manage(ChannelShortName)
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    If ManageType = "Special" Then
        Response.Write "<tr class='title'>"
        Response.Write "  <td height='22'>" & GetSpecialList() & "</td></tr>" & vbCrLf
    Else
        Response.Write "  <tr class='title'>"
        Response.Write "    <td height='22'>" & GetRootClass() & "</td>"
        Response.Write "  </tr>" & GetChild_Root() & ""
    End If
    Response.Write "</table><br>"


    Select Case ManageType
    Case "Check"
        Call ShowContentManagePath(ChannelShortName & "���")
    Case "HTML"
        Call ShowContentManagePath(ChannelShortName & "����")
    Case "Recyclebin"
        Call ShowContentManagePath(ChannelShortName & "����վ����")
    Case "Special"
        Call ShowContentManagePath("ר��" & ChannelShortName & "����")
    Case Else
        Call ShowContentManagePath(ChannelShortName & "����")
    End Select

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "    <form name='myform' method='Post' action='Admin_Soft.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "     <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "          <tr class='title' height='22'> "
    Response.Write "            <td height='22' width='30' align='center'><strong>ѡ��</strong></td>"
    Response.Write "            <td width='25' align='center'><strong>ID</strong></td>"
    If ManageType = "Special" Then
        Response.Write "        <td width='120' align='center'><strong>����ר��</strong></td>"
    End If
    Response.Write "            <td align='center' ><strong>" & ChannelShortName & "����</strong></td>"
    Response.Write "            <td width='60' align='center' ><strong>¼��</strong></td>"
    Response.Write "            <td width='40' align='center' ><strong>������</strong></td>"
    Response.Write "            <td width='80' align='center' ><strong>" & ChannelShortName & "����</strong></td>"
    Response.Write "            <td width='60' align='center' ><strong>���״̬</strong></td>"
    If UseCreateHTML > 0 And ObjInstalled_FSO = True And ManageType <> "Special" Then
        Response.Write "            <td width='40' align='center' ><strong>������</strong></td>"
    End If
    If ManageType = "Check" Then
        Response.Write "            <td width='80' align='center' ><strong>��˲���</strong></td>"
    ElseIf ManageType = "HTML" Then
        Response.Write "            <td width='180' align='center' ><strong>����HTML����</strong></td>"
    ElseIf ManageType = "Recyclebin" Then
        Response.Write "            <td width='100' align='center' ><strong>����վ����</strong></td>"
    ElseIf ManageType = "Special" Then
        Response.Write "            <td width='100' align='center' ><strong>ר��������</strong></td>"
    Else
        Response.Write "            <td width='180' align='center' ><strong>����������</strong></td>"
    End If
    Response.Write "          </tr>"

    If ManageType = "Special" Then
        sql = "select top " & MaxPerPage & " I.InfoID,I.SpecialID,S.SoftID,SP.SpecialName,S.SoftName,S.SoftVersion,S.Keyword,S.Author,S.UpdateTime,S.Inputer,"
        sql = sql & "S.SoftPicUrl,S.DownloadUrl,S.SoftSize,S.DecompressPassword,S.Hits,S.OnTop,S.Elite,S.Status,S.Stars,S.InfoPoint,S.VoteID"
        sql = sql & "  from PE_Soft S right join (PE_InfoS I left join PE_Special SP on I.SpecialID=SP.SpecialID) on S.SoftID=I.ItemID "
    Else
        If ClassID = -1 Or (ClassID > 0 And Child = 0) Then
            sql = "select top " & MaxPerPage & " S.ClassID,S.SoftID,S.SoftName,S.SoftVersion,S.Keyword,S.Author,S.UpdateTime,S.Inputer,"
            sql = sql & "S.SoftPicUrl,S.DownloadUrl,S.SoftSize,S.DecompressPassword,S.Hits,S.OnTop,S.Elite,S.Status,S.Stars,S.InfoPoint,S.VoteID"
            sql = sql & "  from PE_Soft S "
        Else
            sql = "select top " & MaxPerPage & " S.ClassID,S.SoftID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview,S.SoftName,S.SoftVersion,S.Keyword,S.Author,S.UpdateTime,S.Inputer,"
            sql = sql & "S.SoftPicUrl,S.DownloadUrl,S.SoftSize,S.DecompressPassword,S.Hits,S.OnTop,S.Elite,S.Status,S.Stars,S.InfoPoint,S.VoteID"
            sql = sql & " from PE_Soft S left join PE_Class C on S.ClassID=C.ClassID "
        End If
    End If
    
    Querysql = " where S.ChannelID=" & ChannelID
    If ManageType = "Special" Then
        Querysql = Querysql & " and I.ModuleType=" & ModuleType
    End If
    If ManageType = "Recyclebin" Then
        Querysql = Querysql & " and S.Deleted=" & PE_True & ""
    Else
        Querysql = Querysql & " and S.Deleted=" & PE_False & ""
    End If
    If ManageType = "HTML" Then
        If Created = "False" Then
            Querysql = Querysql & " and (S.CreateTime<=S.UpdateTime or S.CreateTime is Null)"
        ElseIf Created = "True" Then
            Querysql = Querysql & " and S.CreateTime>S.UpdateTime"
        End If
        Querysql = Querysql & " and S.Status=3" '�����Ϊ�����ʱ���������ɹ����г���
    Else
        Select Case Status
        Case -2 '�˸�
            Querysql = Querysql & " and S.Status=-2"
        Case -1 '�ݸ�
            Querysql = Querysql & " and S.Status=-1"
        Case 0  '�����
            Querysql = Querysql & " and S.Status>=0 and S.Status<" & MyStatus
        Case 1  '�����
            Querysql = Querysql & " and S.Status>=" & MyStatus
        Case Else
            Querysql = Querysql & " and S.Status>-1"
        End Select
        If OnTop = "True" Then
            Querysql = Querysql & " and S.OnTop=" & PE_True & ""
        End If
        If IsElite = "True" Then
            Querysql = Querysql & " and S.Elite=" & PE_True & ""
        End If
        If IsHot = "True" Then
            Querysql = Querysql & " and S.Hits>=" & HitsOfHot & ""
        End If
    End If
    
    If ClassID <> 0 Then
        If Child > 0 Then
            Querysql = Querysql & " and S.ClassID in (" & arrChildID & ")"
        Else
            Querysql = Querysql & " and S.ClassID=" & ClassID
        End If
    End If
    If SpecialID > 0 Then
        Querysql = Querysql & " and I.SpecialID=" & SpecialID
    End If
    If ManageType = "My" Then
        Querysql = Querysql & " and S.Inputer='" & UserName & "' "
    End If
    If Keyword <> "" Then
        Select Case strField
        Case "SoftName"
            Querysql = Querysql & " and S.SoftName like '%" & Keyword & "%' "
        Case "SoftIntro"
            Querysql = Querysql & " and S.SoftIntro like '%" & Keyword & "%' "
        Case "Author"
            Querysql = Querysql & " and S.Author like '%" & Keyword & "%' "
        Case "Inputer"
            Querysql = Querysql & " and S.Inputer='" & Keyword & "' "
        Case Else
            Querysql = Querysql & " and S.SoftName like '%" & Keyword & "%' "
        End Select
    End If
    
    If ManageType = "Special" Then
        totalPut = PE_CLng(Conn.Execute("select Count(*) from PE_InfoS I inner join PE_Soft S on I.ItemID=S.SoftID " & Querysql)(0))
    Else
        totalPut = PE_CLng(Conn.Execute("select Count(*) from PE_Soft S " & Querysql)(0))
    End If
    If CurrentPage < 1 Then
        CurrentPage = 1
    End If
    If (CurrentPage - 1) * MaxPerPage > totalPut Then
        If (totalPut Mod MaxPerPage) = 0 Then
            CurrentPage = totalPut \ MaxPerPage
        Else
            CurrentPage = totalPut \ MaxPerPage + 1
        End If
    End If
    If CurrentPage > 1 Then
        If ManageType = "Special" Then
            Querysql = Querysql & " and I.InfoID < (select min(InfoID) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " I.InfoID from PE_InfoS I inner join PE_Soft S on I.ItemID=S.SoftID " & Querysql & " order by I.InfoID desc) as QuerySoft)"
        Else
            Querysql = Querysql & " and S.SoftID < (select min(SoftID) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " S.SoftID from PE_Soft S " & Querysql & " order by S.SoftID desc) as QuerySoft)"
        End If
    End If
    If ManageType = "Special" Then
        sql = sql & Querysql & " order by I.InfoID desc"
    Else
        sql = sql & Querysql & " order by S.SoftID desc"
    End If

    Set rsSoftList = Server.CreateObject("ADODB.Recordset")
    rsSoftList.Open sql, Conn, 1, 1
    If rsSoftList.BOF And rsSoftList.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>"

        If ClassID > 0 Then
            Response.Write "����Ŀ��������Ŀ��û���κ�"
        Else
            Response.Write "û���κ�"
        End If

        Select Case Status
        Case -2
            Response.Write "�˸�"
        Case -1
            Response.Write "�ݸ�"
        Case 0
            Response.Write "<font color=blue>�����</font>��" & ChannelShortName & "��"
        Case 1
            Response.Write "<font color=green>�����</font>��" & ChannelShortName & "��"
        Case Else
            Response.Write ChannelShortName & "��"
        End Select
        Response.Write "<br><br></td></tr>"
    Else
        Dim SoftNum, SoftPath
        SoftNum = 0
        Do While Not rsSoftList.EOF
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            If ManageType = "Special" Then
                Response.Write "        <td width='30' align='center'><input name='InfoID' type='checkbox' onclick='CheckItem(this)' id='InfoID' value='" & rsSoftList("InfoID") & "'></td>"
                Response.Write "        <td width='25' align='center'>" & rsSoftList("InfoID") & "</td>"
                Response.Write "        <td width='120' align='center'>"
                If rsSoftList("SpecialID") > 0 Then
                    Response.Write "<a href='" & FileName & "&SpecialID=" & rsSoftList("SpecialID") & "'>" & rsSoftList("SpecialName") & "</a>"
                Else
                    Response.Write "&nbsp;"
                End If
                Response.Write "</td>"
            Else
                Response.Write "        <td width='30' align='center'><input name='SoftID' type='checkbox' onclick='CheckItem(this)' id='SoftID' value='" & rsSoftList("SoftID") & "'></td>"
                Response.Write "        <td width='25' align='center'>" & rsSoftList("SoftID") & "</td>"
            End If
            Response.Write "        <td>"
            If ManageType <> "Special" Then
                If rsSoftList("ClassID") <> ClassID And ClassID <> -1 Then
                    Response.Write "<a href='" & FileName & "&ClassID=" & rsSoftList("ClassID") & "'>["
                    If rsSoftList("ClassName") <> "" Then
                        Response.Write rsSoftList("ClassName")
                    Else
                        Response.Write "<font color='gray'>�������κ���Ŀ</font>"
                    End If
                    Response.Write "]</a>&nbsp;"
                End If
            End If
            
            Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Show&SoftID=" & rsSoftList("SoftID") & "'"
            Response.Write " title='��&nbsp;&nbsp;&nbsp;&nbsp;�ƣ�" & rsSoftList("SoftName") & vbCrLf & "��&nbsp;&nbsp;&nbsp;&nbsp;����" & rsSoftList("SoftVersion") & vbCrLf & "��&nbsp;&nbsp;&nbsp;&nbsp;�ߣ�" & rsSoftList("Author") & vbCrLf & "����ʱ�䣺" & rsSoftList("UpdateTime") & vbCrLf
            Response.Write "���ش�����" & rsSoftList("Hits") & vbCrLf & "�� �� �֣�" & Mid(rsSoftList("Keyword"), 2, Len(rsSoftList("Keyword")) - 2) & vbCrLf & "�Ƽ��ȼ���"
            If rsSoftList("Stars") = 0 Then
                Response.Write "��"
            Else
                Response.Write String(rsSoftList("Stars"), "��")
            End If
            Response.Write vbCrLf & "����" & PointName & "����" & rsSoftList("InfoPoint")
            Response.Write "'>" & rsSoftList("SoftName")
            If rsSoftList("SoftVersion") <> "" Then
                Response.Write "&nbsp;&nbsp;" & rsSoftList("SoftVersion")
            End If
            Response.Write "</a>"
            If CheckDownloadUrl(rsSoftList("DownloadUrl")) = False Then
                Response.Write " <font color='red'>��</font>"
            End If
            Response.Write "</td>"
            Response.Write "      <td width='60' align='center'><a href='" & FileName & "&field=Inputer&keyword=" & rsSoftList("Inputer") & "' title='������鿴���û�¼�������" & ChannelShortName & "'>" & rsSoftList("Inputer") & "</a></td>"
            Response.Write "      <td width='40' align='center'>" & rsSoftList("Hits") & "</td>"
            Response.Write "      <td width='80' align='center'>"
            If rsSoftList("OnTop") = True Then
                Response.Write "<font color=blue>��</font> "
            Else
                Response.Write "&nbsp;&nbsp;&nbsp;"
            End If
            If rsSoftList("Hits") >= HitsOfHot Then
                Response.Write "<font color=red>��</font> "
            Else
                Response.Write "&nbsp;&nbsp;&nbsp;"
            End If
            If rsSoftList("Elite") = True Then
                Response.Write "<font color=green>��</font> "
            Else
                Response.Write "&nbsp;&nbsp;&nbsp;"
            End If
            If Trim(rsSoftList("SoftPicUrl")) <> "" Then
                Response.Write "<font color=blue>ͼ</font>"
            Else
                Response.Write "&nbsp;&nbsp;"
            End If
            If rsSoftList("VoteID") > 0 Then
                Response.Write "<a href='" & InstallDir & "Vote.asp?ID=" & rsSoftList("VoteID") & "&Action=Show' target='_blank'>��</a>"
            Else
                Response.Write "&nbsp;&nbsp;"
            End If
            Response.Write "    </td>"
            Response.Write "    <td width='60' align='center'>"
            Select Case rsSoftList("Status")
            Case -2
                Response.Write "<font color=gray>�˸�</font>"
            Case -1
                Response.Write "<font color=gray>�ݸ�</font>"
            Case 0
                Response.Write "<font color=red>�����</font>"
            Case 1
                Response.Write "<font color=blue>һ��ͨ��</font>"
            Case 2
                Response.Write "<font color=green>����ͨ��</font>"
            Case 3
                Response.Write "<font color=black>����ͨ��</font>"
            End Select
            Response.Write "    </td>"

            Dim iClassPurview
            If UseCreateHTML > 0 And ObjInstalled_FSO = True And ManageType <> "Special" Then
                Response.Write "    <td width='40' align='center'>"
                If ClassID = -1 Or (ClassID > 0 And Child = 0) Then
                    iClassPurview = ClassPurview
                    SoftPath = HtmlDir & GetItemPath(StructureType, ParentDir, ClassDir, rsSoftList("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsSoftList("UpdateTime"), rsSoftList("SoftID")) & FileExt_Item
                Else
                    iClassPurview = PE_CLng(rsSoftList("ClassPurview"))
                    SoftPath = HtmlDir & GetItemPath(StructureType, rsSoftList("ParentDir"), rsSoftList("ClassDir"), rsSoftList("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsSoftList("UpdateTime"), rsSoftList("SoftID")) & FileExt_Item
                End If
                If fso.FileExists(Server.MapPath(SoftPath)) Then
                    Response.Write "<a href='#' title='�ļ�λ�ã�" & SoftPath & "'><b>��</b></a>"
                Else
                    Response.Write "<font color=red><b>��</b></font>"
                End If
                Response.Write "</td>"
            End If
            Select Case ManageType
            Case "Check"
                Response.Write "<td width='120' align='center'>"
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
                    If rsSoftList("Status") <= MyStatus Then
                        If rsSoftList("Status") > -1 Then
                            Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Reject&SoftID=" & rsSoftList("SoftID") & "'>ֱ���˸�</a>&nbsp;&nbsp;"
                        End If
                        If rsSoftList("Status") < MyStatus Then
                            Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Check&SoftID=" & rsSoftList("SoftID") & "'>���</a>&nbsp;&nbsp;"
                            Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=SetPassed&SoftID=" & rsSoftList("SoftID") & "'>ͨ��</a>&nbsp;&nbsp;"
                        Else
                            Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=CancelPassed&SoftID=" & rsSoftList("SoftID") & "'>ȡ�����</a>"
                        End If
                    End If
                End If
                Response.Write "</td>"
            Case "HTML"
                Response.Write "    <td width='180' align='left'>&nbsp;"
                Response.Write "<a href='Admin_CreateSoft.asp?ChannelID=" & ChannelID & "&Action=CreateSoft&SoftID=" & rsSoftList("SoftID") & "' title='���ɱ�" & ChannelShortName & "��HTMLҳ��'>�����ļ�</a>&nbsp;"
                If fso.FileExists(Server.MapPath(SoftPath)) Then
                    Response.Write "<a href='" & SoftPath & "' target='_blank' title='�鿴��" & ChannelShortName & "��HTMLҳ��'>�鿴�ļ�</a>&nbsp;"
                    Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=DelFile&SoftID=" & rsSoftList("SoftID") & "' title='ɾ����" & ChannelShortName & "��HTMLҳ��' onclick=""return confirm('ȷ��Ҫɾ����" & ChannelShortName & "��HTMLҳ����');"">ɾ���ļ�</a>&nbsp;"
                End If
                Response.Write "</td>"
            Case "Recyclebin"
                Response.Write "<td width='100' align='center'>"
                Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=ConfirmDel&SoftID=" & rsSoftList("SoftID") & "' onclick=""return confirm('ȷ��Ҫ����ɾ����" & ChannelShortName & "�𣿳���ɾ�����޷���ԭ��');"">����ɾ��</a> "
                Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Restore&SoftID=" & rsSoftList("SoftID") & "'>��ԭ</a>"
                Response.Write "</td>"
            Case "Special"
                Response.Write "<td width='100' align='center'>"
                If rsSoftList("SpecialID") > 0 Then
                    Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=DelFromSpecial&InfoID=" & rsSoftList("InfoID") & "' onclick=""return confirm('ȷ��Ҫ����" & ChannelShortName & "��������ר����ɾ����');"">������ר����ɾ��</a> "
                End If
                Response.Write "</td>"
            Case Else
                Response.Write "    <td width='150' align='left'>&nbsp;"
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or CheckPurview_Class(arrClass_Input, ParentPath & "," & ClassID) Or UserName = rsSoftList("Inputer") Then
                    Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Modify&SoftID=" & rsSoftList("SoftID") & "'>�޸�</a>&nbsp;"
                Else
                    Response.Write "<font color='gray'>�޸�&nbsp;</font>"
                End If
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Or UserName = rsSoftList("Inputer") Then
                    Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Del&SoftID=" & rsSoftList("SoftID") & "' onclick=""return confirm('ȷ��Ҫɾ����" & ChannelShortName & "��ɾ�����㻹���Դӻ���վ�л�ԭ��');"">ɾ��</a>&nbsp;"
                Else
                    Response.Write "<font color='gray'>ɾ��&nbsp;</font>"
                End If
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
                    If rsSoftList("OnTop") = False Then
                        Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=SetOnTop&SoftID=" & rsSoftList("SoftID") & "'>�̶�</a>&nbsp;"
                    Else
                        Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=CancelOnTop&SoftID=" & rsSoftList("SoftID") & "'>���</a>&nbsp;"
                    End If
                    If rsSoftList("Elite") = False Then
                        Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=SetElite&SoftID=" & rsSoftList("SoftID") & "'>��Ϊ�Ƽ�</a>"
                    Else
                        Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=CancelElite&SoftID=" & rsSoftList("SoftID") & "'>ȡ���Ƽ�</a>"
                    End If
                End If
                Response.Write "</td>"
            End Select
            Response.Write "</tr>"

            SoftNum = SoftNum + 1
            If SoftNum >= MaxPerPage Then Exit Do
            rsSoftList.MoveNext
        Loop
    End If
    rsSoftList.Close
    Set rsSoftList = Nothing
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>ѡ�б�ҳ��ʾ������" & ChannelShortName & "</td><td>"
    Select Case ManageType
    Case "Check"
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
            Response.Write "<input name='submit1' type='submit' value=' ���ͨ�� ' onClick=""document.myform.Action.value='SetPassed'"">&nbsp;&nbsp;"
            Response.Write "<input name='submit2' type='submit' value=' ȡ����� ' onClick=""document.myform.Action.value='CancelPassed'"">&nbsp;&nbsp;"
            If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
                Response.Write "<input name='submit3' type='submit' value=' ����ɾ�� ' onClick=""document.myform.Action.value='Del'"">"
            End If
        End If
    Case "HTML"
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
            Response.Write "<input name='CreateType' type='hidden' id='CreateType' value='1'>"
            Response.Write "<input name='ClassID' type='hidden' id='ClassID' value='" & ClassID & "'>"
            If ClassID > 0 Then
                If UseCreateHTML = 1 Or UseCreateHTML = 3 And ClassPurview < 2 Then
                    Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateClass';document.myform.CreateType.value='1';document.myform.action='Admin_CreateSoft.asp';"" value='���ɵ�ǰ��Ŀ�б�ҳ'>&nbsp;&nbsp;"
                End If
                If ClassPurview = 0 Then
                    Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateSoft';document.myform.CreateType.value='2';document.myform.action='Admin_CreateSoft.asp';"" value='���ɵ�ǰ��Ŀ��" & ChannelShortName & "'>&nbsp;&nbsp;"
                End If
            Else
                Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateIndex';document.myform.CreateType.value='1';document.myform.action='Admin_CreateSoft.asp';"" value='������ҳ'>&nbsp;&nbsp;"
                If UseCreateHTML = 1 Or UseCreateHTML = 3 Then
                    Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateClass';document.myform.CreateType.value='2';document.myform.action='Admin_CreateSoft.asp';"" value='����������Ŀ�б�ҳ'>&nbsp;&nbsp;"
                End If
                Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateSoft';document.myform.CreateType.value='3';document.myform.action='Admin_CreateSoft.asp';"" value='��������" & ChannelShortName & "'>&nbsp;&nbsp;"
            End If
            Response.Write "<input name='submit2' type='submit' id='submit2' onClick=""document.myform.Action.value='CreateSoft';document.myform.action='Admin_CreateSoft.asp';"" value='����ѡ����" & ChannelShortName & "'>&nbsp;&nbsp;&nbsp;&nbsp;"
            Response.Write "<input name='Submit3' type='submit' id='submit3' onClick=""document.myform.Action.value='DelFile';document.myform.action='Admin_Soft.asp'"" value='ɾ��ѡ��" & ChannelShortName & "��HTML�ļ�'>"
        End If
    Case "Recyclebin"
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
            Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='ConfirmDel'"" value=' ����ɾ�� '>&nbsp;"
            Response.Write "<input name='Submit2' type='submit' id='Submit2' onClick=""document.myform.Action.value='ClearRecyclebin'"" value='��ջ���վ'>&nbsp;&nbsp;&nbsp;&nbsp;"
            Response.Write "<input name='Submit3' type='submit' id='Submit3' onClick=""document.myform.Action.value='Restore'"" value='��ԭѡ����" & ChannelShortName & "'>&nbsp;"
            Response.Write "<input name='Submit4' type='submit' id='Submit4' onClick=""document.myform.Action.value='RestoreAll'"" value='��ԭ����" & ChannelShortName & "'>"
        End If
    Case "Special"
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
            Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='DelFromSpecial'"" value='������ר�����Ƴ�'> "
            Response.Write "<input name='Submit2' type='submit' id='Submit2' onClick=""document.myform.Action.value='AddToSpecial'"" value='��ӵ�����ר����'> "
            Response.Write "<input name='Submit2' type='submit' id='Submit2' onClick=""document.myform.Action.value='MoveToSpecial'"" value='�ƶ�����һר����'>"
        End If
    Case Else
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
            Response.Write "<input name='submit1' type='submit' value=' ����ɾ�� ' onClick=""document.myform.Action.value='Del'""> "
            If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
                Response.Write "<input type='submit' name='Submit4' value=' �����ƶ� ' onClick=""document.myform.Action.value='MoveToClass'""> "
                Response.Write "<input type='submit' name='Submit3' value=' �������� ' onClick=""document.myform.Action.value='Batch'""> "
                Response.Write "<input name='submit1' type='submit' value=' ���ͨ�� ' onClick=""document.myform.Action.value='SetPassed'""> "
                Response.Write "<input name='submit2' type='submit' value=' ȡ����� ' onClick=""document.myform.Action.value='CancelPassed'""> "
            End If
        End If
    End Select
    
    Response.Write "<input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName & "", True)
    End If

    Response.Write "<form method='Get' name='SearchForm' action='" & FileName & "'>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='80' align='right'><strong>" & ChannelShortName & "������</strong></td>"
    Response.Write "   <td>"
    Response.Write "<select name='Field' size='1'>"
    Response.Write "<option value='SoftName' selected>" & ChannelShortName & "����</option>"
    Response.Write "<option value='SoftIntro'>" & ChannelShortName & "���</option>"
    Response.Write "<option value='Author'>" & ChannelShortName & "����</option>"
    If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
        Response.Write "<option value='Inputer'>¼����</option>"
    End If
    Response.Write "<option value='ID'>" & ChannelShortName & "ID</option>"
    Response.Write "</select>"
    Response.Write "<select name='ClassID'><option value=''>������Ŀ</option>" & GetClass_Option(1, 0) & "</select>"
    Response.Write "<input type='text' name='keyword'  size='20' value='�ؼ���' maxlength='50' onFocus='this.select();'>"
    Response.Write "<input type='submit' name='Submit'  value='����'>"
    Response.Write "<input name='ManageType' type='hidden' id='ManageType' value='" & ManageType & "'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "</td></tr></table></form>"
    Response.Write "<br><b>˵����</b><br>&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "�����еĸ���壺<font color=blue>��</font>----�̶�" & ChannelShortName & "��<font color=red>��</font>----����" & ChannelShortName & "��<font color=green>��</font>----�Ƽ�" & ChannelShortName & "��<font color=blue>ͼ</font>----������ͼ��<font color=red>��</font>----" & ChannelShortName & "��ַ���д������ӣ�<font color=black>��</font>----�е���<br><br>"
End Sub

Sub ShowJS_Soft()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function SelectSoft(){" & vbCrLf
    Response.Write "  var arr=showModalDialog('Admin_SelectFile.asp?ChannelID=" & ChannelID & "&dialogtype=Softpic', '', 'dialogWidth:820px; dialogHeight:600px; help: no; scroll: yes; status: no');" & vbCrLf
    Response.Write "  if(arr!=null){" & vbCrLf
    Response.Write "    var ss=arr.split('|');" & vbCrLf
    Response.Write "    document.myform.SoftPicUrl.value=ss[0];" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function SelectFile(){" & vbCrLf
    Response.Write "  var arr=showModalDialog('Admin_SelectFile.asp?ChannelID=" & ChannelID & "&dialogtype=Soft', '', 'dialogWidth:820px; dialogHeight:600px; help: no; scroll: yes; status: no');" & vbCrLf
    Response.Write "  if(arr!=null){" & vbCrLf
    Response.Write "    var ss=arr.split('|');" & vbCrLf
    Response.Write "    strSoftUrl=ss[0];" & vbCrLf
    Response.Write "    var url='" & XmlText("Soft", "DownloadUrlTip", "���ص�ַ") & "'+(document.myform.DownloadUrl.length+1)+'|'+strSoftUrl;" & vbCrLf
    Response.Write "    document.myform.DownloadUrl.options[document.myform.DownloadUrl.length]=new Option(url,url);" & vbCrLf
    Response.Write "    document.myform.SoftSize.value=ss[1];" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function AddUrl(){" & vbCrLf
    Response.Write "  var thisurl='" & XmlText("Soft", "DownloadUrlTip", "���ص�ַ") & "'+(document.myform.DownloadUrl.length+1)+'|http://'; " & vbCrLf
    Response.Write "  var url=prompt('���������ص�ַ���ƺ����ӣ��м��á�|��������',thisurl);" & vbCrLf
    Response.Write "  if(url!=null&&url!=''){document.myform.DownloadUrl.options[document.myform.DownloadUrl.length]=new Option(url,url);}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function ModifyUrl(){" & vbCrLf
    Response.Write "  if(document.myform.DownloadUrl.length==0) return false;" & vbCrLf
    Response.Write "  var thisurl=document.myform.DownloadUrl.value; " & vbCrLf
    Response.Write "  if (thisurl=='') {alert('����ѡ��һ�����ص�ַ���ٵ��޸İ�ť��');return false;}" & vbCrLf
    Response.Write "  var url=prompt('���������ص�ַ���ƺ����ӣ��м��á�|��������',thisurl);" & vbCrLf
    Response.Write "  if(url!=thisurl&&url!=null&&url!=''){document.myform.DownloadUrl.options[document.myform.DownloadUrl.selectedIndex]=new Option(url,url);}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function DelUrl(){" & vbCrLf
    Response.Write "  if(document.myform.DownloadUrl.length==0) return false;" & vbCrLf
    Response.Write "  var thisurl=document.myform.DownloadUrl.value; " & vbCrLf
    Response.Write "  if (thisurl=='') {alert('����ѡ��һ�����ص�ַ���ٵ�ɾ����ť��');return false;}" & vbCrLf
    Response.Write "  document.myform.DownloadUrl.options[document.myform.DownloadUrl.selectedIndex]=null;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckForm()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  var CurrentMode=editor.CurrentMode;" & vbCrLf
    Response.Write "  if (CurrentMode==0){" & vbCrLf
    Response.Write "    document.myform.SoftIntro.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else if(CurrentMode==1){" & vbCrLf
    Response.Write "   document.myform.SoftIntro.value=editor.HtmlEdit.document.body.innerText;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    alert('Ԥ��״̬���ܱ��棡���Ȼص��༭״̬���ٱ���');" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.SoftName.value==''){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "���Ʋ���Ϊ�գ�');" & vbCrLf
    Response.Write "    document.myform.SoftName.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Keyword.value==''){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('�ؼ��ֲ���Ϊ�գ�');" & vbCrLf
    Response.Write "    document.myform.Keyword.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.SoftIntro.value==''){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "��鲻��Ϊ�գ�');" & vbCrLf
    Response.Write "    editor.HtmlEdit.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.DownloadUrl.length==0){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "���ص�ַ����Ϊ�գ�');" & vbCrLf
    Response.Write "    document.myform.DownloadUrl.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.SoftAddType.value=='3'&&document.myform.DownloadUrl.value==''){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "���ص�ַ����Ϊ�գ�');" & vbCrLf
    Response.Write "    document.myform.DownloadUrl.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf

    Response.Write "  var obj=document.myform.ClassID;" & vbCrLf
    Response.Write "  var iCount=0;" & vbCrLf
    Response.Write "  for(var i=0;i<obj.length;i++){" & vbCrLf
    Response.Write "    if(obj.options[i].selected==true){" & vbCrLf
    Response.Write "      iCount=iCount+1;" & vbCrLf
    Response.Write "      if(obj.options[i].value==''){" & vbCrLf
    Response.Write "        ShowTabs(0);" & vbCrLf
    Response.Write "        alert('" & ChannelShortName & "������Ŀ����ָ��Ϊ�ⲿ��Ŀ��');" & vbCrLf
    Response.Write "        document.myform.ClassID.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "      }" & vbCrLf
    Response.Write "      if(obj.options[i].selected==true&&obj.options[i].value=='0'){" & vbCrLf
    Response.Write "        ShowTabs(0);" & vbCrLf
    Response.Write "        alert('ָ������Ŀ���������" & ChannelShortName & "��ֻ������������Ŀ�����" & ChannelShortName & "��');" & vbCrLf
    Response.Write "        document.myform.ClassID.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "      }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (iCount==0){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('��ѡ��������Ŀ��');" & vbCrLf
    Response.Write "    document.myform.ClassID.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Action.value!='Preview'){" & vbCrLf
    Response.Write "    for(var i=0;i<document.myform.DownloadUrl.length;i++){" & vbCrLf
    Response.Write "      if (document.myform.DownloadUrls.value=='') document.myform.DownloadUrls.value=document.myform.DownloadUrl.options[i].value;" & vbCrLf
    Response.Write "      else document.myform.DownloadUrls.value+='$$$'+document.myform.DownloadUrl.options[i].value;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID>0){" & vbCrLf
    Response.Write "    Tabs_Bottom.style.display='none';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    Tabs_Bottom.style.display='';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "    if(ID==0){" & vbCrLf
    Response.Write "      editor.yToolbarsCss();" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function CopyTitle(){" & vbCrLf
    Response.Write "  if (document.myform.VoteTitle.value==''){" & vbCrLf
    Response.Write "     document.myform.VoteTitle.value = document.myform.SoftName.value;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function moreitem(inputname,listnum,ichannelid,inputype){" & vbCrLf
    Response.Write "    var chedkurl = '../inc/checklist.asp';" & vbCrLf
    Response.Write "    var CheckDOM = new ActiveXObject(""Microsoft.XMLDOM"");" & vbCrLf
    Response.Write "    CheckDOM.async=false;" & vbCrLf
    Response.Write "    var p = CheckDOM.createProcessingInstruction('xml','version=\""1.0\"" encoding=\""gb2312\""'); " & vbCrLf
    Response.Write "    CheckDOM.appendChild(p); " & vbCrLf

    Response.Write "    var CheckRoot = CheckDOM.createElement('root');" & vbCrLf
    Response.Write "    var CField = CheckDOM.createNode(1,'text',''); " & vbCrLf
    Response.Write "    CField.text = $F(inputname);" & vbCrLf
    Response.Write "    CheckRoot.appendChild(CField);" & vbCrLf
    Response.Write "    CField = CheckDOM.createNode(1,'lnum',''); " & vbCrLf
    Response.Write "    CField.text = listnum;" & vbCrLf
    Response.Write "    CheckRoot.appendChild(CField);" & vbCrLf
    Response.Write "    CField = CheckDOM.createNode(1,'channelid',''); " & vbCrLf
    Response.Write "    CField.text = ichannelid;" & vbCrLf
    Response.Write "    CheckRoot.appendChild(CField);" & vbCrLf
    Response.Write "    CField = CheckDOM.createNode(1,'type',''); " & vbCrLf
    Response.Write "    CField.text = inputype;" & vbCrLf
    Response.Write "    CheckRoot.appendChild(CField);" & vbCrLf
    Response.Write "    CField = CheckDOM.createNode(1,'inputname',''); " & vbCrLf
    Response.Write "    CField.text = inputname;" & vbCrLf
    Response.Write "    CheckRoot.appendChild(CField);" & vbCrLf
    Response.Write "    CheckDOM.appendChild(CheckRoot);" & vbCrLf

    Response.Write "    var CHttp = getHTTPObject();" & vbCrLf
    Response.Write "    CHttp.open('POST',chedkurl,true);" & vbCrLf
    Response.Write "    CHttp.onreadystatechange = function () " & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        if(CHttp.readyState == 4 && CHttp.status==200){" & vbCrLf
    Response.Write "            if(CHttp.responseText == ''){" & vbCrLf
    Response.Write "                Element.hide(inputype);" & vbCrLf
    Response.Write "            }else{" & vbCrLf
    Response.Write "                Element.show(inputype);" & vbCrLf
    Response.Write "                $(inputype).innerHTML=CHttp.responseText;" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    CHttp.send(CheckDOM);" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function addinput(iname,ivalue){" & vbCrLf
    Response.Write "  if(iname!='' && ivalue!=''){" & vbCrLf
    Response.Write "     $(iname).value=ivalue;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf

    'Response.Write "function getKey() {" & vbCrLf
    'Response.Write " if(window.event.keyCode==49) {" & vbCrLf
    'Response.Write "   ShowTabs(0);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==50) {" & vbCrLf
    'Response.Write "   ShowTabs(1);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==51) {" & vbCrLf
    'Response.Write "   ShowTabs(2);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==52) {" & vbCrLf
    'Response.Write "   ShowTabs(3);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==53) {" & vbCrLf
    'Response.Write "   ShowTabs(4);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==54) {" & vbCrLf
    'Response.Write "   ShowTabs(5);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==55) {" & vbCrLf
    'Response.Write "   ShowTabs(6);CopyTitle();" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write "}" & vbCrLf
    'Response.Write "document.onkeypress = getKey;" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub ShowTabs_Title()
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>������Ϣ</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>����ר��</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'"
    If FoundInArr(arrEnabledTabs, "SoftParameter", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">�������</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>��������</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(4)'"
    If FoundInArr(arrEnabledTabs, "Charge", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">�շ�ѡ��</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(5);CopyTitle()'"
    If FoundInArr(arrEnabledTabs, "Vote", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">��������</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(6)'>�Զ���ѡ��</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub

Sub ShowTabs_Bottom()
    Response.Write "<table id='Tabs_Bottom' width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center' height='24'>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title4' onclick='ShowTabs(0)'>������Ϣ</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(1)'>����ר��</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(2)'"
    If FoundInArr(arrEnabledTabs, "SoftParameter", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">�������</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(3)'>��������</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(4)'"
    If FoundInArr(arrEnabledTabs, "Charge", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">�շ�ѡ��</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(5);CopyTitle()'"
    If FoundInArr(arrEnabledTabs, "Vote", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">��������</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(6)'>�Զ���ѡ��</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub

Sub Add()
    Call ShowJS_Soft

    Response.Write "<br><table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "'>" & ChannelName & "����</a>&nbsp;&gt;&gt;&nbsp;���" & ChannelShortName & "</td></tr></table>"
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Soft.asp' target='_self'>"

    Call ShowTabs_Title

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf

    Call ShowTr_Class
    
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "���ƣ�</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='SoftName' type='text' value='' autocomplete='off' size='50' maxlength='255' onPropertyChange=""moreitem('SoftName',10," & ChannelID & ",'satitle3');"" onBlur=""setTimeout('Element.hide(satitle3)',500);""> <font color='#FF0000'>*</font><input type='button' name='checksame' value='����Ƿ������ͬ��" & ChannelShortName & "��' onclick=""showModalDialog('Admin_CheckSameTitle.asp?ModuleType=" & ModuleType & "&Title='+document.myform.SoftName.value,'checksame','dialogWidth:350px; dialogHeight:250px; help: no; scroll: no; status: no');"">"
    Response.Write "              </div><div id=""satitle3"" style='display:none'></div>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ؼ��֣�</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='Keyword' type='text' id='Keyword' value='" & Trim(Session("Keyword")) & "' autocomplete='off' size='50' maxlength='255' onPropertyChange=""moreitem('Keyword',10," & ChannelID & ",'skey');"" onBlur=""setTimeout('Element.hide(skey)',500);""> <font color='#FF0000'>*</font> " & GetKeywordList("Admin", ChannelID)
    Response.Write "              </div><div id=""skey"" style='display:none'></div>"
    Response.Write "              <font color='#0000FF'>�����������" & ChannelShortName & "�����������ؼ��֣��м���<font color='#FF0000'>��|��</font>���������ܳ���&quot;'&?;:()���ַ���</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>����/�����̣�</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Author' type='text' id='Author' value='" & Trim(Session("Author")) & "' size='50' maxlength='100'>" & GetAuthorList("Admin", ChannelID, UserName)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>����ͼ��</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='SoftPicUrl' type='text' id='SoftPicUrl' size='80' maxlength='200'>"
    Response.Write "              <input type='button' name='Button2' value='�����ϴ�ͼƬ��ѡ��' onclick='SelectSoft()'>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'></td>"
    Response.Write "            <td><table><tr><td>�ϴ�" & ChannelShortName & "ͼƬ��</td><td><iframe style='top:2px' id='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=Softpic' frameborder=0 scrolling=no width='450' height='25'></iframe></td></tr></table></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��飺</td>"
    Response.Write "            <td><textarea name='SoftIntro' cols='80' rows='10' id='SoftIntro' style='display:none'></textarea>"
    Response.Write "                <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=3&tContentid=SoftIntro' frameborder='1' scrolling='no' width='700' height='200' ></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    '����ģʽ��������ַ
    If AddType = 3 Then
        Response.Write "           <tr class='tdbg' id='UrlType1'>"
        Response.Write "             <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��ַ�� </strong></td>"
        Response.Write "            <td>"
        Response.Write "              <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
        Response.Write "                <tr>"
        Response.Write "                  <td>"
        Response.Write "                  ������ĸ�ʽ��Soft/UploadSoft/20060331/200603311712.rar<br>"
        Response.Write "                  <input name='DownloadUrl' type='text' value='' size='65' maxlength='255'> <font color='#FF0000'>*</font>"
        Response.Write "                    <input type='hidden' name='SoftAddType' value='3'>"
        Response.Write "                  </td>"
        Response.Write "                </tr>"
        Response.Write "              </table>"
        Response.Write "            </td>"
        Response.Write "          </tr>"
    Else
        Response.Write "          <tr class='tdbg' id='UrlType2'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'>�ϴ�" & ChannelShortName & "��</td>"
        Response.Write "            <td>"

        Response.Write "              <iframe style='top:2px' ID='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=Soft' frameborder=0 scrolling=no width='450' height='25'></iframe>"
        Response.Write "                    <input type='hidden' name='SoftAddType' value='1'>"
        Response.Write "            </td>"
        Response.Write "          </tr>"	
        Response.Write "          <tr class='tdbg' id='UrlType2'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��ַ��</td>"
        Response.Write "            <td>"
        Response.Write "              <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
        Response.Write "                <tr>"
        Response.Write "                  <td width='450'>"
        Response.Write "                    <input type='hidden' name='DownloadUrls' value=''>"
        Response.Write "                    <select name='DownloadUrl' style='width:500;height:100' size='2' ondblclick='return ModifyUrl();'></select>"
        Response.Write "                  </td>"
        Response.Write "                  <td>"
        Response.Write "                    <input type='button' name='Softselect' value='�����ϴ�" & ChannelShortName & "��ѡ��' onclick='SelectFile()'><br><br>"
        Response.Write "                    <input type='button' name='addurl' value='����ⲿ��ַ' onclick='AddUrl();'><br>"
        Response.Write "                    <input type='button' name='modifyurl' value='�޸ĵ�ǰ��ַ' onclick='return ModifyUrl();'><br>"
        Response.Write "                    <input type='button' name='delurl' value='ɾ����ǰ��ַ' onclick='DelUrl();'>"
        Response.Write "                  </td>"
        Response.Write "                </tr>"
        Response.Write "              <tr><td  colspan='3'>ϵͳ�ṩ���ϴ�����ֻ�ʺ��ϴ��Ƚ�С��" & ChannelShortName & "����ASPԴ����ѹ�����������" & ChannelShortName & "�Ƚϴ�" & MaxFileSize \ 1024 & "M���ϣ�������ʹ��FTP�ϴ�������Ҫʹ��ϵͳ�ṩ���ϴ����ܣ������ϴ���������ռ�÷�������CPU��Դ��FTP�ϴ����뽫��ַ���Ƶ�����ĵ�ַ���С�</td></tr>"		
        Response.Write "              </table>"
        Response.Write "            </td>"
        Response.Write "          </tr>"

    End If

    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��С��</td>"
    Response.Write "            <td><input name='SoftSize' type='text' id='SoftSize' size='10' maxlength='10'> K</strong></td>"
    Response.Write "          </tr>"

    Call ShowTabs_Status_Add

    Response.Write "        </tbody>" & vbCrLf
    
    Call ShowTabs_Special(SpecialID, "")
    
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "�汾��</td>"
    Response.Write "            <td><input name='SoftVersion' type='text' size='15' maxlength='100'></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��Դ��</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='CopyFrom' type='text' id='CopyFrom' value='" & Trim(Session("CopyFrom")) & "' size='50' maxlength='100'>" & GetCopyFromList("Admin", ChannelID)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "���</td>"
    Response.Write "            <td><select name='SoftType' id='SoftType'>" & GetSoftType(0) & "</select></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "���ԣ�</td>"
    Response.Write "            <td><select name='SoftLanguage' id='SoftLanguage'>" & GetSoftLanguage(0) & "</select></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>��Ȩ��ʽ��</td>"
    Response.Write "            <td><select name='CopyrightType' id='CopyrightType'>" & GetCopyrightType(0) & "</select></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "ƽ̨��</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='OperatingSystem' type='text' value='" & XmlText("Soft", "OperatingSystem", "Win9x/NT/2000/XP/") & "' size='80' maxlength='200'> <br>" & GetOperatingSystemList
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��ʾ��ַ��</td>"
    Response.Write "            <td><input name='DemoUrl' type='text' value='http://' size='80' maxlength='200'></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "ע���ַ��</td>"
    Response.Write "            <td><input name='RegUrl' type='text' value='http://' size='80' maxlength='200'></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>��ѹ���룺</td>"
    Response.Write "            <td><input name='DecompressPassword' type='text' id='DecompressPassword' size='30' maxlength='100'></td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf
    
    Call ShowTabs_Property_Add
    
    Call ShowTabs_Purview_Add("����")
    
    Call ShowTabs_Vote_Add
    
    Call ShowTabs_MyField_Add
    
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"

    Call ShowTabs_Bottom

    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "   <input name='AddType' type='hidden' id='AddType' value='" & AddType & "'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='add' type='submit'  id='Add' value=' �� �� ' onClick=""document.myform.Action.value='SaveAdd';document.myform.target='_self';"" style='cursor:hand;'>&nbsp; "
    Response.Write "   <input name='Preview' type='submit'  id='Preview' value=' Ԥ �� ' onClick=""document.myform.Action.value='Preview';document.myform.target='_blank';"" style='cursor:hand;'>"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
End Sub

Sub Modify()
    Dim rsSoft, tmpAuthor, tmpCopyFrom
    Dim AddType, imageDownloadUrl
    
    If SoftID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�" & ChannelShortName & "ID</li>"
        Exit Sub
    Else
        SoftID = PE_CLng(SoftID)
    End If
    Set rsSoft = Conn.Execute("select * from PE_Soft where SoftID=" & SoftID & "")
    If rsSoft.BOF And rsSoft.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���" & ChannelShortName & "</li>"
        rsSoft.Close
        Set rsSoft = Nothing
        Exit Sub
    End If

    ClassID = rsSoft("ClassID")
    Dim arrSpecialID
    arrSpecialID = GetSpecialIDArr(ModuleType, SoftID)
    
    If rsSoft("Inputer") <> UserName Then
        Call CheckClassPurview(Action, ClassID)
    End If

    If FoundErr = True Then
        rsSoft.Close
        Set rsSoft = Nothing
        Exit Sub
    End If
     
    imageDownloadUrl = Trim(Replace(rsSoft("DownloadUrl"), "@@@", ""))
    If InStr(rsSoft("DownloadUrl"), "@@@") > 0 Then
        AddType = 3
    Else
        AddType = 1 '��ģʽ�͸߼�ģʽ��һ�����޸Ľ���
    End If

    tmpAuthor = rsSoft("Author")
    tmpCopyFrom = rsSoft("CopyFrom")
    EmailOfReject = Replace(EmailOfReject, "{$Title}", rsSoft("SoftName"))
    EmailOfPassed = Replace(EmailOfPassed, "{$Title}", rsSoft("SoftName"))

    Call ShowJS_Soft

    Response.Write "<br><table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "'>" & ChannelName & "����</a>&nbsp;&gt;&gt;&nbsp;�޸�" & ChannelShortName & "</td></tr></table>"
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Soft.asp' target='_self'>"


    Call ShowTabs_Title

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf

    Call ShowTr_Class
    
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "���ƣ�</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='SoftName' type='text' value='" & rsSoft("SoftName") & "' autocomplete='off' size='50' maxlength='255' onPropertyChange=""moreitem('SoftName',10," & ChannelID & ",'satitle3');"" onBlur=""setTimeout('Element.hide(satitle3)',500);""> <font color='#FF0000'>*</font>"
    Response.Write "              </div><div id=""satitle3"" style='display:none'></div>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ؼ��֣�</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='Keyword' type='text' id='Keyword' value='" & Mid(rsSoft("Keyword"), 2, Len(rsSoft("Keyword")) - 2) & "' autocomplete='off' size='50' maxlength='255' onPropertyChange=""moreitem('Keyword',10," & ChannelID & ",'skey');"" onBlur=""setTimeout('Element.hide(skey)',500);""> <font color='#FF0000'>*</font> " & GetKeywordList("Admin", ChannelID)
    Response.Write "              </div><div id=""skey"" style='display:none'></div>"
    Response.Write "              <font color='#0000FF'>�����������" & ChannelShortName & "�����������ؼ��֣��м���<font color='#FF0000'>��|��</font>���������ܳ���&quot;'&?;:()���ַ���</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>����/�����̣�</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Author' type='text' id='Author' value='" & tmpAuthor & "' size='50' maxlength='100'>" & GetAuthorList("Admin", ChannelID, UserName)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>����ͼ��</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='SoftPicUrl' type='text' id='SoftPicUrl' value='" & rsSoft("SoftPicUrl") & "' size='80' maxlength='200'>"
    Response.Write "              <input type='button' name='Button2' value='�����ϴ�ͼƬ��ѡ��' onclick='SelectSoft()'>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'></td>"
    Response.Write "            <td><table><tr><td>�ϴ�" & ChannelShortName & "ͼƬ��</td><td><iframe style='top:2px' id='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=Softpic' frameborder=0 scrolling=no width='450' height='25'></iframe></td></tr></table></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��飺</td>"
    Response.Write "            <td><textarea name='SoftIntro' cols='80' rows='10' id='SoftIntro' style='display:none'>" & Server.HTMLEncode(FilterBadTag(rsSoft("SoftIntro"), rsSoft("Inputer"))) & "</textarea>"
    Response.Write "                <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=3&tContentid=SoftIntro' frameborder='1' scrolling='no' width='700' height='200' ></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    
    '����ģʽ��ַ
    If AddType = 3 Then
        Response.Write "           <tr class='tdbg'>"
        Response.Write "             <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��ַ�� </strong></td>"
        Response.Write "            <td>"
        Response.Write "              <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
        Response.Write "                <tr>"
        Response.Write "                  <td>"
        Response.Write "                   ������ĸ�ʽ��Soft/UploadSoft/20060331/200603311712.rar<br>"
        Response.Write "              <input name='DownloadUrl' type='text' value='" & imageDownloadUrl & "' size='65' maxlength='255'> <font color='#FF0000'>*</font>"
        Response.Write "                    <input type='hidden' name='SoftAddType' value='3'>"
        Response.Write "                  </td>"
        Response.Write "                </tr>"
        Response.Write "              </table>"
        Response.Write "            </td>"
        Response.Write "          </tr>"
    Else
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'>�ϴ�" & ChannelShortName & "��</td>"
        Response.Write "            <td>"
        Response.Write "              <iframe style='top:2px' ID='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=Soft' frameborder=0 scrolling=no width='450' height='25'></iframe>"
        Response.Write "                    <input type='hidden' name='SoftAddType' value='1'>"
        Response.Write "            </td>"
        Response.Write "          </tr>"	
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��ַ��</td>"
        Response.Write "            <td>"
        Response.Write "              <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
        Response.Write "                <tr>"
        Response.Write "                  <td>"
        Response.Write "                    <input type='hidden' name='DownloadUrls' value=''>"
        Response.Write "                    <select name='DownloadUrl' style='width:400;height:100' size='2' ondblclick='return ModifyUrl();'>"
        Dim DownloadUrls, arrDownloadUrls, iTemp
        DownloadUrls = rsSoft("DownloadUrl")
        If InStr(DownloadUrls, "$$$") > 1 Then
            arrDownloadUrls = Split(DownloadUrls, "$$$")
            For iTemp = 0 To UBound(arrDownloadUrls)
                Response.Write "<option value='" & arrDownloadUrls(iTemp) & "'>" & arrDownloadUrls(iTemp) & "</option>"
            Next
        Else
            Response.Write "<option value='" & DownloadUrls & "'>" & DownloadUrls & "</option>"
        End If
        Response.Write "                    </select>"
        Response.Write "                  </td>"
        Response.Write "                  <td>"
        Response.Write "                    <input type='button' name='Softselect' value='�����ϴ�" & ChannelShortName & "��ѡ��' onclick='SelectFile()'><br><br>"
        Response.Write "                    <input type='button' name='addurl' value='����ⲿ��ַ' onclick='AddUrl();'><br>"
        Response.Write "                    <input type='button' name='modifyurl' value='�޸ĵ�ǰ��ַ' onclick='return ModifyUrl();'><br>"
        Response.Write "                    <input type='button' name='delurl' value='ɾ����ǰ��ַ' onclick='DelUrl();'>"
        Response.Write "                  </td>"
        Response.Write "                </tr>"
        Response.Write "              <tr><td  colspan='3'>�ʺ��ϴ��Ƚ�С��" & ChannelShortName & "����ASPԴ����ѹ�����������" & ChannelShortName & "�Ƚϴ�" & MaxFileSize \ 1024 & "M���ϣ�������ʹ��FTP�ϴ�������Ҫʹ��ϵͳ�ṩ���ϴ����ܣ������ϴ���������ռ�÷�������CPU��Դ��FTP�ϴ����뽫��ַ���Ƶ�����ĵ�ַ���С�</td></tr>"			
        Response.Write "              </table>"
        Response.Write "            </td>"
        Response.Write "          </tr>"		

    End If
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��С��</td>"
    Response.Write "            <td><input name='SoftSize' type='text' id='SoftSize' value='" & rsSoft("SoftSize") & "' size='10' maxlength='10'> K</td>"
    Response.Write "          </tr>"

    Call ShowTabs_Status_Modify(rsSoft)
    Response.Write "        </tbody>" & vbCrLf
    
    
    
    Call ShowTabs_Special(arrSpecialID, "")
    
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "�汾��</td>"
    Response.Write "            <td><input name='SoftVersion' type='text' size='15' maxlength='100' value='" & rsSoft("SoftVersion") & "'></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��Դ��</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='CopyFrom' type='text' id='CopyFrom' value='" & tmpCopyFrom & "' size='50' maxlength='100'>" & GetCopyFromList("Admin", ChannelID)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "���</td>"
    Response.Write "            <td>"
    Response.Write "              <select name='SoftType' id='SoftType'>" & GetSoftType(rsSoft("SoftType")) & "</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "              <strong>" & ChannelShortName & "���ԣ�</strong> <select name='SoftLanguage' id='SoftLanguage'>" & GetSoftLanguage(rsSoft("SoftLanguage")) & "</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "              <strong>��Ȩ��ʽ��</strong> <select name='CopyrightType' id='CopyrightType'>" & GetCopyrightType(rsSoft("CopyrightType")) & "</select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "ƽ̨��</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='OperatingSystem' type='text' value='" & rsSoft("OperatingSystem") & "' size='80' maxlength='200'> <br>" & GetOperatingSystemList
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "��ʾ��ַ��</td>"
    Response.Write "            <td><input name='DemoUrl' type='text' value='" & rsSoft("DemoUrl") & "' size='80' maxlength='200'></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "ע���ַ��</td>"
    Response.Write "            <td><input name='RegUrl' type='text' value='" & rsSoft("RegUrl") & "' size='80' maxlength='200'></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>��ѹ���룺</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='DecompressPassword' type='text' id='DecompressPassword' value='" & rsSoft("DecompressPassword") & "' size='30' maxlength='100'>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf
    
    
    Call ShowTabs_Property_Modify(rsSoft)
    
    Call ShowTabs_Purview_Modify("����", rsSoft, "")
    
    Call ShowTabs_Vote_Modify(rsSoft)

    Call ShowTabs_MyField_Modify(rsSoft)
    
    
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>" & vbCrLf


    Call ShowTabs_Bottom

    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "   <input name='SoftID' type='hidden' id='SoftID' value='" & rsSoft("SoftID") & "'>"
    Response.Write "   <input name='AddType' type='hidden' id='AddType' value='" & AddType & "'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='Save' type='submit' value='�����޸Ľ��' onClick=""document.myform.Action.value='SaveModify';document.myform.target='_self';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Save' type='submit' value='���Ϊ��" & ChannelShortName & "' onClick=""document.myform.Action.value='SaveAdd';document.myform.target='_self';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Preview' type='submit'  id='Preview' value=' Ԥ �� ' onClick=""document.myform.Action.value='Preview';document.myform.target='_blank';"" style='cursor:hand;'>"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
    rsSoft.Close
    Set rsSoft = Nothing
End Sub

Sub SaveSoft()
    Dim rsSoft, sql, trs, i
    Dim SoftID, ClassID, SpecialID, SoftName, Author, tAuthor, CopyFrom
    Dim Inputer, Editor, DownloadUrls, UpdateTime
    Dim AddType
    Dim arrSpecialID

    AddType = PE_CLng(Request.Form("AddType"))
    If AddType = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������" & ChannelShortName & "���ģʽΪ��</li>"
    End If
    
    SoftID = Trim(Request.Form("SoftID"))
    ClassID = Trim(Request.Form("ClassID"))
    SpecialID = Trim(Request.Form("SpecialID"))
    SoftName = Trim(Request.Form("SoftName"))
    Keyword = Trim(Request.Form("Keyword"))
    Author = PE_HTMLEncode(Trim(Request.Form("Author")))
    CopyFrom = PE_HTMLEncode(Trim(Request.Form("CopyFrom")))
    If AddType = 3 Then
        DownloadUrls = "@@@" & PE_HTMLEncode(Trim(Request.Form("DownloadUrl"))) '�����ַ
    Else
        DownloadUrls = PE_HTMLEncode(Trim(Request.Form("DownloadUrls")))
    End If
    UpdateTime = PE_CDate(Trim(Request.Form("UpdateTime")))
    Status = PE_CLng(Trim(Request.Form("Status")))

    Inputer = UserName
    Editor = AdminName
    
    Call CheckClassPurview(Action, ClassID)
    If FoundErr = True Then Exit Sub
    
    If SoftName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "���Ʋ���Ϊ��</li>"
    End If
    Keyword = ReplaceBadChar(Keyword)
    If Keyword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������" & ChannelShortName & "�ؼ���</li>"
    Else
        Call SaveKeyword(Keyword)
    End If

    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-2 or ChannelID=" & ChannelID & "")
    Do While Not rsField.EOF
        If rsField("EnableNull") = False Then
            If Trim(Request(rsField("FieldName"))) = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>������" & rsField("Title") & "��</li>"
            End If
        End If
        rsField.MoveNext
    Loop
    
    If FoundErr = True Then
        Exit Sub
    End If

    SoftName = PE_HTMLEncode(SoftName)
    Keyword = "|" & Keyword & "|"
    If Author = "" Then Author = XmlText("BaseText", "DefAuthor", "����")
    If CopyFrom = "" Then CopyFrom = XmlText("BaseText", "DefCopyFrom", "��վԭ��")
    
    If Status = 1 Then
        Status = MyStatus
    End If

    Call GetUser(UserName)

    Call SaveVote


    If SpecialID = "" Then
        arrSpecialID = Split("0", ",")
    Else
        arrSpecialID = Split(SpecialID, ",")
    End If
    PresentExp = Int(PresentExp * PresentExpTimes)

    Set rsSoft = Server.CreateObject("adodb.recordset")
    If Action = "SaveAdd" Then
        If Session("SoftName") = SoftName And DateDiff("S", Session("AddTime"), Now()) < 100 Then
            FoundErr = True
            ErrMsg = "<li>�벻Ҫ�ظ����ͬһ" & ChannelItemUnit & ChannelShortName & "</li>"
            Exit Sub
        Else
            Session("SoftName") = SoftName
            Session("AddTime") = Now()
            SoftID = GetNewID("PE_Soft", "SoftID")

            For i = 0 To UBound(arrSpecialID)
                Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (" & ModuleType & "," & SoftID & "," & PE_CLng(arrSpecialID(i)) & ")")
            Next
            
            sql = "select top 1 * from PE_Soft"
            rsSoft.Open sql, Conn, 1, 3
            rsSoft.addnew
            rsSoft("SoftID") = SoftID
            rsSoft("ChannelID") = ChannelID
            rsSoft("Inputer") = Inputer

            Dim blogid
            If UserID <> "" And UserID > 0 Then
                Set blogid = Conn.Execute("select top 1 ID from PE_Space where Type=1 and UserID=" & UserID)
                If blogid.BOF And blogid.EOF Then
                    rsSoft("BlogID") = 0
                Else
                    rsSoft("BlogID") = blogid("ID")
                End If
                Set blogid = Nothing
            End If

            If Status = 3 Then
                Conn.Execute ("update PE_User set UserExp=UserExp+" & PresentExp & " where UserName='" & Inputer & "'")
            End If
        End If
    ElseIf Action = "SaveModify" Then
        If SoftID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����ȷ��" & ChannelShortName & "ID��ֵ</li>"
        Else
            SoftID = PE_CLng(SoftID)
            sql = "select * from PE_Soft where SoftID=" & SoftID
            rsSoft.Open sql, Conn, 1, 3
            If rsSoft.BOF And rsSoft.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�Ҳ�����" & ChannelShortName & "�������Ѿ���������ɾ����</li>"
            Else
            
                'ɾ�����ɵ��ļ�����Ϊ���ɵ��ļ����ܻ����Ÿ���ʱ�䣬����Ȩ�޵ȷ����仯������������ļ�
                If UseCreateHTML > 0 And ObjInstalled_FSO = True Then
                    Dim tClass, SoftPath
                    Set tClass = Conn.Execute("select ParentDir,ClassDir from PE_Class where ClassID=" & rsSoft("ClassID") & "")
                    If tClass.BOF And tClass.EOF Then
                        ParentDir = "/"
                        ClassDir = ""
                    Else
                        ParentDir = tClass("ParentDir")
                        ClassDir = tClass("ClassDir")
                    End If
                    SoftPath = HtmlDir & GetItemPath(StructureType, ParentDir, ClassDir, rsSoft("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsSoft("UpdateTime"), rsSoft("SoftID"))
                    If fso.FileExists(Server.MapPath(SoftPath & FileExt_Item)) Then
                        DelSerialFiles Server.MapPath(SoftPath & FileExt_Item)
                    End If
                End If
                If rsSoft("Inputer") <> UserName And rsSoft("Status") <> Status And (Status = -2 Or Status = 3) Then
                    Call SendEmailOfCheck(rsSoft("Inputer"), rsSoft)
                End If

                Call UpdateUserData(0, rsSoft("Inputer"), 0, 0)
            
                If rsSoft("Status") < 3 And Status = 3 Then
                    Conn.Execute ("update PE_User set UserExp=UserExp+" & rsSoft("PresentExp") & " where UserName='" & rsSoft("Inputer") & "'")
                End If
                If rsSoft("Status") = 3 And Status < 3 Then
                    Conn.Execute ("update PE_User set UserExp=UserExp-" & rsSoft("PresentExp") & " where UserName='" & rsSoft("Inputer") & "'")
                End If


                Dim rsInfo, sqlInfo, j
                i = 0
                sqlInfo = "select * from PE_InfoS where ModuleType=" & ModuleType & " and ItemID=" & SoftID & " order by InfoID"
                Set rsInfo = Server.CreateObject("adodb.recordset")
                rsInfo.Open sqlInfo, Conn, 1, 3
                Do While Not rsInfo.EOF
                    If i > UBound(arrSpecialID) Then
                        rsInfo.Delete
                    Else
                        rsInfo("SpecialID") = arrSpecialID(i)
                    End If
                    rsInfo.Update
                    rsInfo.MoveNext
                    i = i + 1
                Loop
                rsInfo.Close
                Set rsInfo = Nothing
                If (i - 1) < UBound(arrSpecialID) Then
                    For j = i To UBound(arrSpecialID)
                        If PE_CLng(arrSpecialID(j)) <> 0 Then
                            Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (" & ModuleType & "," & SoftID & "," & PE_CLng(arrSpecialID(j)) & ")")
                        End If
                    Next
                End If
                
                
                
            End If
        End If
    End If
    rsSoft("ClassID") = ClassID
    rsSoft("SoftName") = SoftName
    rsSoft("Keyword") = Keyword
    rsSoft("SoftVersion") = Trim(Request.Form("SoftVersion"))
    rsSoft("SoftType") = PE_HTMLEncode(Trim(Request.Form("SoftType")))
    rsSoft("SoftLanguage") = PE_HTMLEncode(Trim(Request.Form("SoftLanguage")))
    rsSoft("CopyrightType") = PE_HTMLEncode(Trim(Request.Form("CopyrightType")))
    rsSoft("OperatingSystem") = PE_HTMLEncode(Trim(Request.Form("OperatingSystem")))
    rsSoft("Author") = Author
    rsSoft("CopyFrom") = CopyFrom
    
    rsSoft("DemoUrl") = PE_HTMLEncode(Trim(Request.Form("DemoUrl")))
    rsSoft("RegUrl") = PE_HTMLEncode(Trim(Request.Form("RegUrl")))
    rsSoft("SoftPicUrl") = PE_HTMLEncode(Trim(Request.Form("SoftPicUrl")))
    rsSoft("SoftIntro") = Trim(Request.Form("SoftIntro"))
    rsSoft("Hits") = PE_CLng(Trim(Request.Form("Hits")))
    rsSoft("DayHits") = PE_CLng(Trim(Request.Form("DayHits")))
    rsSoft("WeekHits") = PE_CLng(Trim(Request.Form("WeekHits")))
    rsSoft("MonthHits") = PE_CLng(Trim(Request.Form("MonthHits")))
    rsSoft("Stars") = PE_CLng(Trim(Request.Form("Stars")))
    rsSoft("UpdateTime") = UpdateTime
    rsSoft("CreateTime") = UpdateTime
    rsSoft("Status") = Status
    rsSoft("Deleted") = False
    rsSoft("PresentExp") = PresentExp
    'rsSoft("Inputer") = Inputer
    rsSoft("Editor") = Editor
    rsSoft("OnTop") = PE_CBool(Trim(Request.Form("OnTop")))
    rsSoft("Elite") = PE_CBool(Trim(Request.Form("Elite")))
    rsSoft("DecompressPassword") = PE_HTMLEncode(Trim(Request.Form("DecompressPassword")))
    rsSoft("SoftSize") = PE_CLng(Trim(Request.Form("SoftSize")))
    rsSoft("DownloadUrl") = DownloadUrls
    rsSoft("SkinID") = PE_CLng(Trim(Request.Form("SkinID")))
    rsSoft("TemplateID") = PE_CLng(Trim(Request.Form("TemplateID")))

    rsSoft("InfoPurview") = PE_CLng(Trim(Request.Form("InfoPurview")))
    rsSoft("arrGroupID") = ReplaceBadChar(Trim(Request.Form("GroupID")))
    rsSoft("InfoPoint") = PE_CLng(Trim(Request.Form("InfoPoint")))
    rsSoft("ChargeType") = PE_CLng(Trim(Request.Form("ChargeType")))
    rsSoft("PitchTime") = PE_CLng(Trim(Request.Form("PitchTime")))
    rsSoft("ReadTimes") = PE_CLng(Trim(Request.Form("ReadTimes")))
    rsSoft("DividePercent") = PE_CLng(Trim(Request.Form("DividePercent")))

        rsSoft("VoteID") = VoteID
        If Not (rsField.BOF And rsField.EOF) Then
            rsField.MoveFirst
            Do While Not rsField.EOF
                If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                    rsSoft(Trim(rsField("FieldName"))) = Trim(Request(rsField("FieldName")))
                End If
                rsField.MoveNext
            Loop
        End If
        Set rsField = Nothing
    
    rsSoft.Update
    rsSoft.Close
    Set rsSoft = Nothing
    Call UpdateChannelData(ChannelID)
    If Action = "SaveAdd" Then
        Call UpdateUserData(0, Inputer, 0, 0)
    End If

    Response.Write "<br><br>"
    Response.Write "<table class='border' align=center width='500' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "  <tr align=center> "
    Response.Write "    <td  height='22' align='center' class='title' colspan='2'> "
    If Action = "SaveAdd" Then
        Response.Write "<b>���" & ChannelShortName & "�ɹ�</b>"
    Else
        Response.Write "<b>�޸�" & ChannelShortName & "�ɹ�</b>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='100' align='right' class='tdbg5'><strong>������Ŀ��</td>"
    Response.Write "    <td width='400'>" & ShowClassPath() & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='100' align='right' class='tdbg5'><strong>" & ChannelShortName & "���ƣ�</td>"
    Response.Write "    <td width='400'>" & SoftName & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='100' align='right' class='tdbg5'><strong>" & ChannelShortName & "�汾��</td>"
    Response.Write "    <td width='400'>" & Trim(Request("SoftVersion")) & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='100' align='right' class='tdbg5'><strong>" & ChannelShortName & "���ߣ�</td>"
    Response.Write "    <td width='400'>" & Author & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='100' align='right' class='tdbg5'><strong>�� �� �֣�</td>"
    Response.Write "    <td width='400'>" & Mid(Keyword, 2, Len(Keyword) - 2) & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='100' align='right' class='tdbg5'><strong>" & ChannelShortName & "״̬��</strong></td>"
    Response.Write "    <td width='400'>"
    If Status = -1 Then
        Response.Write "�ݸ�"
    ElseIf Status = -2 Then
        Response.Write "�˸�"
    Else
        Response.Write arrStatus(Status)
    End If
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='40' colspan='2' align='center'>"
    Response.Write "��<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Modify&SoftID=" & SoftID & "'>�޸Ĵ�" & ChannelShortName & "</a>��&nbsp;"
    Response.Write "��<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Add&AddType=" & AddType & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>�������" & ChannelShortName & "</a>��&nbsp;"
    Response.Write "��<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Manage&ClassID=" & ClassID & "'>" & ChannelShortName & "����</a>��&nbsp;"
    Response.Write "��<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Show&SoftID=" & SoftID & "'>Ԥ��" & ChannelShortName & "����</a>��"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf

    Session("Keyword") = Trim(Request("Keyword"))
    Session("Author") = Author
    Session("CopyFrom") = CopyFrom
    Session("SkinID") = PE_CLng(Trim(Request("SkinID")))
    Session("TemplateID") = PE_CLng(Trim(Request("TemplateID")))

    Call ClearSiteCache(0)
    Call CreateAllJS

    If Status = 3 And UseCreateHTML > 0 And ObjInstalled_FSO = True And Trim(Request("CreateImmediate")) = "Yes" Then
        Response.Write "<br><iframe id='CreateSoft' width='100%' height='210' frameborder='0' src='Admin_CreateSoft.asp?ChannelID=" & ChannelID & "&Action=CreateSoft2&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&SoftID=" & SoftID & "&ShowBack=No'></iframe>"
    End If

End Sub

Sub Show()
    If SoftID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��" & ChannelShortName & "ID��</li>"
        Exit Sub
    End If
    
    Dim rsSoft, PurviewChecked, PurviewChecked2
    PurviewChecked = False
    PurviewChecked2 = False
    Set rsSoft = Conn.Execute("select * from PE_Soft where SoftID=" & SoftID & "")
    If rsSoft.BOF And rsSoft.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���" & ChannelShortName & "��</li>"
        rsSoft.Close
        Set rsSoft = Nothing
        Exit Sub
    End If
    ClassID = rsSoft("ClassID")

    Call CheckClassPurview(Action, ClassID)
    If FoundErr = True Then
        rsSoft.Close
        Set rsSoft = Nothing
        Exit Sub
    End If

    Dim arrSpecialID
    arrSpecialID = GetSpecialIDArr(ModuleType, SoftID)

    Response.Write "<SCRIPT language='javascript'>" & vbCrLf
    Response.Write "function resizepic(thispic){" & vbCrLf
    Response.Write "  if(thispic.width>600) thispic.width=600;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function bbimg(o){" & vbCrLf
    Response.Write "  var zoom=parseInt(o.style.zoom, 10)||100;" & vbCrLf
    Response.Write "  zoom+=event.wheelDelta/12;" & vbCrLf
    Response.Write "  if (zoom>0) o.style.zoom=zoom+'%';" & vbCrLf
    Response.Write "  if (zoom>0) o.style.zoom=zoom+'%';" & vbCrLf
    Response.Write "  return false;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf

    Response.Write "<br>�����ڵ�λ�ã�&nbsp;<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "'>" & ChannelShortName & "����</a>&nbsp;&gt;&gt;&nbsp;"
    If ParentID > 0 Then
        Dim sqlPath, rsPath
        sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
        Set rsPath = Conn.Execute(sqlPath)
        Do While Not rsPath.EOF
            Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    Response.Write "<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;�鿴" & ChannelShortName & "��Ϣ��"
    Response.Write rsSoft("SoftName") & "<br><br>"
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>�����Ϣ</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>����ר��</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>�շ���Ϣ</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>�Զ���ѡ��</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "���ƣ�</td>"
    Response.Write "  <td colspan='3'><strong>" & PE_HTMLEncode(rsSoft("SoftName")) & "&nbsp;&nbsp;" & PE_HTMLEncode(rsSoft("SoftVersion")) & "</strong></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>�ļ���С��</td>"
    Response.Write "  <td width='300'>" & rsSoft("SoftSize") & " K" & "</td>"
    Response.Write "  <td colspan='2' rowspan='5' align=center valign='middle'>"
    If rsSoft("SoftPicUrl") = "" Then
        Response.Write "<img src='" & InstallDir & "images/nopic.gif'>"
    Else
        Response.Write "<img src='" & GetSoftPicUrl(rsSoft("SoftPicUrl")) & "' width='150'>"
    End If
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>�� �� �̣�</td>"
    Response.Write "  <td width='300'>" & rsSoft("Author") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "��Դ��</td>"
    Response.Write "  <td width='300'>" & rsSoft("CopyFrom") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "ƽ̨��</td>"
    Response.Write "  <td width='300'>" & rsSoft("OperatingSystem") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "���</td>"
    Response.Write "  <td width='300'>" & rsSoft("SoftType") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "���ԣ�</td>"
    Response.Write "  <td width='300'>" & rsSoft("SoftLanguage") & "</td>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>��Ȩ��ʽ��</td>"
    Response.Write "  <td width='300'>" & rsSoft("CopyrightType") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>��ʾ��ַ��</td>"
    Response.Write "  <td width='300'><a href='" & rsSoft("DemoUrl") & "' target='_blank'>" & rsSoft("DemoUrl") & "</a></td>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>ע���ַ��</td>"
    Response.Write "  <td width='300'><a href='" & rsSoft("RegUrl") & "' target='_blank'>" & rsSoft("RegUrl") & "</a></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>��ѹ���룺</td>"
    Response.Write "  <td width='300'>" & rsSoft("DecompressPassword") & "</td>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>���ֵȼ���</td>"
    Response.Write "  <td>" & String(rsSoft("Stars"), "��") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "��ӣ�</td>"
    Response.Write "  <td width='300'>" & rsSoft("Inputer") & "</td>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>���α༭��</td>"
    Response.Write "  <td>"
    If rsSoft("Status") = 3 Then
        Response.Write rsSoft("Editor")
    Else
        Response.Write "��"
    End If
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>���ʱ�䣺</td>"
    Response.Write "  <td width='300'>" & rsSoft("UpdateTime") & "</td>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>���ش�����</td>"
    Response.Write "  <td colspan='3'>���գ�" & rsSoft("DayHits") & "&nbsp;&nbsp;&nbsp;&nbsp;���ܣ�" & rsSoft("WeekHits") & "&nbsp;&nbsp;&nbsp;&nbsp;���£�" & rsSoft("MonthHits") & "&nbsp;&nbsp;&nbsp;&nbsp;�ܼƣ�" & rsSoft("Hits")
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>���ص�ַ��</td>"
    Response.Write "  <td colspan='3'>"
    Call ShowDownloadUrls(rsSoft("DownloadUrl"))
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "��飺</td>"
    Response.Write "  <td height='100' colspan='3'>" & FilterBadTag(rsSoft("SoftIntro"), rsSoft("Inputer")) & "</td>"
    Response.Write "</tr>"
    Response.Write "        </tbody>" & vbCrLf
    
    Call ShowTabs_Special(arrSpecialID, " disabled")

    Call ShowTabs_Purview_Modify("����", rsSoft, " disabled")

    Call ShowTabs_MyField_View(rsSoft)

    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf



    Response.Write "<form name='formA' method='get' action='Admin_Soft.asp'><p align='center'>"
    Response.Write "<input type='hidden' name='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input type='hidden' name='SoftID' value='" & SoftID & "'>"
    Response.Write "<input type='hidden' name='Action' value=''>" & vbCrLf
    If rsSoft("Deleted") = False Then
        PurviewChecked = CheckClassPurview("Manage", ClassID)
        PurviewChecked2 = CheckClassPurview("Check", ClassID)
        If (rsSoft("Inputer") = UserName And rsSoft("Status") = 0) Or PurviewChecked = True Then
            Response.Write "<input type='submit' name='submit' value='�޸�/���' onclick=""document.formA.Action.value='Modify'"">&nbsp;&nbsp;"
            Response.Write "<input type='submit' name='submit' value=' ɾ �� ' onclick=""document.formA.Action.value='Del'"">&nbsp;&nbsp;"
        End If
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
            Response.Write "<input type='submit' name='submit' value=' �� �� ' onclick=""document.formA.Action.value='MoveToClass'"">&nbsp;&nbsp;"
        End If
        If PurviewChecked2 = True Then
            If rsSoft("Status") > -1 Then
                Response.Write "<input type='submit' name='submit' value='ֱ���˸�' onclick=""document.formA.Action.value='Reject'"">&nbsp;&nbsp;"
            End If
            If rsSoft("Status") < MyStatus Then
                Response.Write "<input type='submit' name='submit' value='" & arrStatus(MyStatus) & "' onclick=""document.formA.Action.value='SetPassed'"">&nbsp;&nbsp;"
            End If
            If rsSoft("Status") >= MyStatus Then
                Response.Write "<input type='submit' name='submit' value='ȡ�����' onclick=""document.formA.Action.value='CancelPassed'"">&nbsp;&nbsp;"
            End If
        End If
        If PurviewChecked = True Then
            If rsSoft("OnTop") = False Then
                Response.Write "<input type='submit' name='submit' value='��Ϊ�̶�' onclick=""document.formA.Action.value='SetOnTop'"">&nbsp;&nbsp;"
            Else
                Response.Write "<input type='submit' name='submit' value='ȡ���̶�' onclick=""document.formA.Action.value='CancelOnTop'"">&nbsp;&nbsp;"
            End If
            If rsSoft("Elite") = False Then
                Response.Write "<input type='submit' name='submit' value='��Ϊ�Ƽ�' onclick=""document.formA.Action.value='SetElite'"">"
            Else
                Response.Write "<input type='submit' name='submit' value='ȡ���Ƽ�' onclick=""document.formA.Action.value='CancelElite'"">"
            End If
        End If
    Else
        If AdminPurview = 1 Or AdminPurview_Channel = 1 Then
            Response.Write "<input type='submit' name='submit' value='����ɾ��' onclick=""if(confirm('ȷ��Ҫ����ɾ����" & ChannelShortName & "�𣿳���ɾ�����޷���ԭ��')==true){document.formA.Action.value='ConfirmDel';}"">&nbsp;&nbsp;"
            Response.Write "<input type='submit' name='submit' value=' �� ԭ ' onclick=""document.formA.Action.value='Restore'"">"
        End If
    End If
    Response.Write "</p></form>"

    rsSoft.Close
    Set rsSoft = Nothing

    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='0'><tr class='tdbg'><td>"
    Response.Write "<li>��һ" & ChannelItemUnit & ChannelShortName & "��"
    Dim rsPrev
    Set rsPrev = Conn.Execute("Select Top 1 S.SoftID,S.SoftName,C.ClassID,C.ClassName from PE_Soft S left join PE_Class C On S.ClassID=C.ClassID Where S.ChannelID=" & ChannelID & " and S.Deleted=" & PE_False & " and S.SoftID<" & SoftID & " order by S.SoftID desc")
    If rsPrev.EOF Then
        Response.Write "û����"
    Else
        Response.Write "[<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&ClassID=" & rsPrev("ClassID") & "'>" & rsPrev("ClassName") & "</a>] <a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Show&SoftID=" & rsPrev("SoftID") & "'>" & rsPrev("SoftName") & "</a>"
    End If
    rsPrev.Close
    Set rsPrev = Nothing
    Response.Write "</li></td</tr><tr class='tdbg'><td><li>��һ" & ChannelItemUnit & ChannelShortName & "��"
    Dim rsNext
    Set rsNext = Conn.Execute("Select Top 1 S.SoftID,S.SoftName,C.ClassID,C.ClassName from PE_Soft S left join PE_Class C On S.ClassID=C.ClassID Where S.ChannelID=" & ChannelID & " and S.Deleted=" & PE_False & " and S.SoftID>" & SoftID & " order by S.SoftID asc")
    If rsNext.EOF Then
        Response.Write "û����"
    Else
        Response.Write "[<a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&ClassID=" & rsNext("ClassID") & "'>" & rsNext("ClassName") & "</a>] <a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Show&SoftID=" & rsNext("SoftID") & "'>" & rsNext("SoftName") & "</a>"
    End If
    rsNext.Close
    Set rsNext = Nothing
    Response.Write "</li></td></tr></table><br>" & vbCrLf

    Dim InfoType
    InfoType = PE_CLng(Trim(Request("InfoType")))

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr align='center' height='24'><td"
    If InfoType = 0 Then
        Response.Write " class='title6'"
    Else
        Response.Write " class='title5' onclick=""window.location.href='Admin_Soft.asp?Action=Show&ChannelID=" & ChannelID & "&SoftID=" & SoftID & "&InfoType=0'"""
    End If
    Response.Write ">�������</td><td"
    If InfoType = 1 Then
        Response.Write " class='title6'"
    Else
        Response.Write " class='title5' onclick=""window.location.href='Admin_Soft.asp?Action=Show&ChannelID=" & ChannelID & "&SoftID=" & SoftID & "&InfoType=1'"""
    End If
    Response.Write ">����շ�</td>"
    Response.Write "<td>&nbsp;</td></tr></table>"
    
    strFileName = "Admin_Soft.asp?Action=Show&ChannelID=" & ChannelID & "&SoftID=" & SoftID & "&InfoType=" & InfoType
    
    Select Case InfoType
    Case 0
        Call ShowComment(SoftID)
    Case 1
        Call ShowConsumeLog(SoftID)
    End Select
End Sub


Sub Preview()
    Response.Write "<br><table width='100%' border=0 align=center cellPadding=2 cellSpacing=1 bgcolor='#FFFFFF' class='border' style='WORD-BREAK: break-all'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='4'>"

    If ClassID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��������Ŀ</li>"
        Exit Sub
    ElseIf ClassID > 0 Then
        Set tClass = Conn.Execute("select ClassName,ParentID,ParentPath from PE_Class where ClassID=" & ClassID)
        If tClass.BOF And tClass.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�Ҳ���ָ������Ŀ</li>"
            Set tClass = Nothing
            Exit Sub
        Else
            ClassName = tClass(0)
            ParentID = tClass(1)
            ParentPath = tClass(2)
        End If
        Set tClass = Nothing
        If ParentID > 0 Then
            Dim sqlPath, rsPath
            sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
            Set rsPath = Conn.Execute(sqlPath)
            Do While Not rsPath.EOF
                Response.Write rsPath(1) & "&nbsp;&gt;&gt;&nbsp;"
                rsPath.MoveNext
            Loop
            rsPath.Close
            Set rsPath = Nothing
        End If
        Response.Write ClassName & "&nbsp;&gt;&gt;&nbsp;"
    End If

    Response.Write PE_HTMLEncode(Request("SoftName"))
    Response.Write " </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "���ƣ�</td>"
    Response.Write "  <td colspan='3'><strong>" & PE_HTMLEncode(Request("SoftName")) & "&nbsp;&nbsp;" & PE_HTMLEncode(Request("SoftVersion")) & "</strong></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>�ļ���С��</td>"
    Response.Write "  <td width='300'>" & Request("SoftSize") & " K" & "</td>"
    Response.Write "  <td colspan='2' rowspan='7' align=center valign='middle'>"
    If Request("SoftPicUrl") = "" Then
        Response.Write "������"
    Else
        Response.Write "<img src='" & GetSoftPicUrl(Request("SoftPicUrl")) & "' width='150'>"
    End If
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>�� �� �̣�</td>"
    Response.Write "  <td width='300'>" & PE_HTMLEncode(Request("Author")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "ƽ̨��</td>"
    Response.Write "  <td width='300'>" & PE_HTMLEncode(Request("OperatingSystem")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "���</td>"
    Response.Write "  <td width='300'>" & PE_HTMLEncode(Request("SoftType")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "���ԣ�</td>"
    Response.Write "  <td width='300'>" & PE_HTMLEncode(Request("SoftLanguage")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>��Ȩ��ʽ��</td>"
    Response.Write "  <td width='300'>" & PE_HTMLEncode(Request("CopyrightType")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>������ӣ�</td>"
    Response.Write "  <td width='300'><a href='" & Request("DemoUrl") & "' target='_blank'>" & ChannelShortName & "��ʾ��ַ</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='" & Request("RegUrl") & "' target='_blank'>" & ChannelShortName & "ע���ַ</a></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>��ѹ���룺</td>"
    Response.Write "  <td width='300'>" & Request("DecompressPassword") & "</td>"
    Response.Write "  <td width='100' align='right'>���ֵȼ���</td>"
    Response.Write "  <td>" & String(Request("Stars"), "��") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>���ʱ�䣺</td>"
    Response.Write "  <td width='300'>" & Now() & "</td>"
    Response.Write "  <td width='100' align='right'>����" & PointName & "����</td>"
    Response.Write "  <td><font color=red> " & Request("InfoPoint") & "</font> ��</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>���ش�����</td>"
    Response.Write "  <td colspan='3'>���գ�" & Request("DayHits") & "&nbsp;&nbsp;&nbsp;&nbsp;���ܣ�" & Request("WeekHits") & "&nbsp;&nbsp;&nbsp;&nbsp;���£�" & Request("MonthHits") & "&nbsp;&nbsp;&nbsp;&nbsp;�ܼƣ�" & Request("Hits")
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>���ص�ַ��</td>"
    Response.Write "  <td colspan='3'>"
    Call ShowDownloadUrls(Request("DownloadUrl"))
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "��飺</td>"
    Response.Write "  <td height='100' colspan='3'>" & Request("SoftIntro") & "</td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    Response.Write "<p align='center'>��<a href='javascript:window.close();'>�رմ���</a>��</p>"
End Sub

Sub ShowReplace()
    Dim i
    Response.Write "<form name='myform' method='post' action='Admin_Soft.asp'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>�� �� �� �� �� ַ �� �� �� ��</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='40' align='right'>��ѡ��</td>"
    Response.Write "    <td height='40'><input name='UrlType' type='radio' value='0' checked>�������ص�ַ<br><input name='UrlType' type='radio' value='1'>ָ�����ص�ַ<br>"
    
    For i = 0 To 9
        Response.Write "<input name='UrlID' type='checkbox' value='" & i & "'>���ص�ַ" & i + 1
        If i = 4 Then
            Response.Write "<br>"
        Else
            Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
        End If
    Next
    Response.Write "</td></tr>"
    
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='40' align='right'>���ַ���</td>"
    Response.Write "    <td height='40'><input name='strSource' type='text' id='strSource' size='60' maxlength='200'><font color='#FF0000'>* ע���Сд</font></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='40' align='right'>�滻�ɣ�</td>"
    Response.Write "    <td height='40'><input name='strTarget' type='text' id='strTarget' size='60' maxlength='200'></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='40' colspan='2' align='center'><input name='Action' type='hidden' id='Action' value='DoReplace'>"
    Response.Write "        <input type='submit' name='Submit' value=' ��ʼ�滻 '>&nbsp; <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Soft.asp?Action=Manage&ChannelID=" & ChannelID & "&Status=9'"" style='cursor:hand;'></td>"
    Response.Write "        <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
End Sub

Sub DoReplace()
    Dim strSource, strTarget, UrlID, UrlType
    Dim arrUrlID, DownloadUrls, arrDownloadUrls, iTemp, strUrls, isArr
    Dim sqlSoft, rsSoft, i, IsReplaceSoftItem

    strSource = Trim(Request("strSource"))
    strTarget = Trim(Request("strTarget"))
    UrlID = Trim(Request("UrlID"))
    UrlType = PE_CLng(Trim(Request("UrlType")))
    IsReplaceSoftItem = False

    If UrlType = 1 Then
        If UrlID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ѡ��Ҫ�滻�����ص�ַ��</li>"
        Else
            UrlID = Replace(UrlID, "'", "")
        End If
    End If

    If strSource = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���������滻���ַ�����</li>"
    End If

    If strTarget = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������滻����ַ�����</li>"
    End If

    If InStr(strSource, "|") > 0 Or InStr(strSource, "$") > 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>Ҫ�滻�����ص�ַ�����С�|���ź͡�$����</li>"
    End If

    If InStr(strTarget, "|") > 0 Or InStr(strTarget, "$") > 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�滻Ϊ���ص�ַ�����С�|���ź͡�$����</li>"
    End If

    If FoundErr = True Then Exit Sub
    i = 0
    
    If UrlType = 1 Then
        If InStr(UrlID, ",") > 0 Then
            isArr = True
            arrUrlID = Split(UrlID, ",")
        Else
            isArr = False
            UrlID = PE_CLng(UrlID)
        End If
    End If
    
    sqlSoft = "select DownloadUrl from PE_Soft where ChannelID=" & ChannelID & " And Deleted = " & PE_False & " order by SoftID"
    Set rsSoft = Server.CreateObject("adodb.recordset")
    rsSoft.Open sqlSoft, Conn, 1, 3

    If UrlType = 0 Then
        Do While Not rsSoft.EOF
            If InStr(rsSoft(0), strSource) > 0 Then
                rsSoft(0) = Replace(rsSoft(0), strSource, strTarget)
                i = i + 1
            End If
            rsSoft.Update
            rsSoft.MoveNext
        Loop
    Else
        Do While Not rsSoft.EOF
            DownloadUrls = rsSoft(0)
            If InStr(DownloadUrls, "$$$") > 1 Then
                arrDownloadUrls = Split(DownloadUrls, "$$$")
                If isArr = True Then
                    For iTemp = 0 To UBound(arrUrlID)
                        If PE_CLng(arrUrlID(iTemp)) <= UBound(arrDownloadUrls) Then
                            strUrls = arrDownloadUrls(arrUrlID(iTemp))
                            If InStr(strUrls, strSource) > 0 Then
                                strUrls = Replace(strUrls, strSource, strTarget)
                                IsReplaceSoftItem = True
                            End If
                            arrDownloadUrls(arrUrlID(iTemp)) = strUrls
                        End If
                    Next
                Else
                    If UrlID <= UBound(arrDownloadUrls) Then
                        strUrls = arrDownloadUrls(UrlID)
                        If InStr(strUrls, strSource) > 0 Then
                            strUrls = Replace(strUrls, strSource, strTarget)
                            IsReplaceSoftItem = True
                        End If
                        arrDownloadUrls(UrlID) = strUrls
                    End If
                End If
                strUrls = ""
                For iTemp = 0 To UBound(arrDownloadUrls)
                    If strUrls = "" Then
                        strUrls = arrDownloadUrls(iTemp)
                    Else
                        strUrls = strUrls & "$$$" & arrDownloadUrls(iTemp)
                    End If
                Next
            Else
                If isArr = True Then
                    If PE_CLng(arrUrlID(0)) = 0 Then
                        If InStr(DownloadUrls, strSource) > 0 Then
                            strUrls = Replace(DownloadUrls, strSource, strTarget)
                            IsReplaceSoftItem = True
                        End If
                    End If
                Else
                    If UrlID = 0 Then
                        If InStr(DownloadUrls, strSource) > 0 Then
                            strUrls = Replace(DownloadUrls, strSource, strTarget)
                            IsReplaceSoftItem = True
                        End If
                    End If
                End If
            End If
            If Trim(strUrls) <> "" Then
                rsSoft(0) = strUrls
                rsSoft.Update
            End If
            If IsReplaceSoftItem = True Then
                i = i + 1
            End If
            rsSoft.MoveNext
        Loop
    End If
    rsSoft.Close
    Set rsSoft = Nothing
    Call WriteSuccessMsg("�����滻���ص�ַ�ɹ������滻�� <font color=red><b>" & i & "</b></font> ��" & ChannelShortName & "�����ص�ַ��", ComeUrl)
End Sub

Sub Other()
    If AdminPurview > 1 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    If UBound(arrFields_Options) <> 3 Then
        arrFields_Options = Split(",,,", ",")
    End If
    Response.Write "<form name='myform' method='post' action='Admin_Soft.asp'>"
    Response.Write "<table width='100%'  border='0' cellpadding='5' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' valign='top' class='title'>"
    Response.Write "    <td colspan='4'><strong>��������</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr align='center' valign='top' class='tdbg'>"
    Response.Write "    <td>" & ChannelShortName & "������"
    Response.Write "      <br><textarea name='SoftTypes' cols='20' rows='10' id='SoftTypes'>" & arrFields_Options(0) & "</textarea>"
    Response.Write "      <br><div align='left'>˵����ÿһ�����Ϊһ��</div>"
    Response.Write "    </td>"
    Response.Write "    <td>" & ChannelShortName & "���Թ���"
    Response.Write "      <br><textarea name='SoftLanguages' cols='20' rows='10' id='SoftLanguages'>" & arrFields_Options(1) & "</textarea>"
    Response.Write "      <br><div align='left'>˵����ÿһ������Ϊһ��</div>"
    Response.Write "      </td>"
    Response.Write "    <td>��Ȩ��ʽ����"
    Response.Write "      <br><textarea name='CopyrightTypes' cols='20' rows='10' id='CopyrightTypes'>" & arrFields_Options(2) & "</textarea>"
    Response.Write "      <br><div align='left'>˵����ÿһ����Ȩ��ʽΪһ��</div>"
    Response.Write "    </td>"
    Response.Write "    <td>" & ChannelShortName & "ƽ̨����"
    Response.Write "      <br><textarea name='OperatingSystems' cols='20' rows='10' id='OperatingSystems'>" & arrFields_Options(3) & "</textarea>"
    Response.Write "      <br><div align='left'>˵����ÿһ������ƽ̨Ϊһ��</div>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr align='center' valign='top' class='tdbg'>"
    Response.Write "    <td colspan='4'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'><input name='Action' type='hidden' id='Action' value='SaveOther'><input type='submit' name='Submit' value=' �������� '></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
End Sub

Sub SaveOther()
    If AdminPurview > 1 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    
    Dim SoftTypes, SoftLanguages, CopyrightTypes, OperatingSystems
    Dim rsChannel, sqlChannel
    SoftTypes = Trim(Request("SoftTypes"))
    SoftLanguages = Trim(Request("SoftLanguages"))
    CopyrightTypes = Trim(Request("CopyrightTypes"))
    OperatingSystems = Trim(Request("OperatingSystems"))
    
    sqlChannel = "select Fields_Options from PE_Channel where ChannelID=" & ChannelID & ""
    Set rsChannel = Server.CreateObject("adodb.recordset")
    rsChannel.Open sqlChannel, Conn, 1, 3
    rsChannel(0) = SoftTypes & "$$$" & SoftLanguages & "$$$" & CopyrightTypes & "$$$" & OperatingSystems
    rsChannel.Update
    rsChannel.Close
    Set rsChannel = Nothing
    Call WriteSuccessMsg("����ɹ���", ComeUrl)
End Sub


'******************************************************************************************
'����Ϊ������������ʹ�õĺ�������ģ��ʵ�ֹ������ƣ��޸�ʱע��ͬʱ�޸ĸ�ģ�����ݡ�
'******************************************************************************************

Sub Batch()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If

    Response.Write "<SCRIPT language='javascript'>" & vbCrLf
    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.BatchClassID.length;i++){" & vbCrLf
    Response.Write "    document.myform.BatchClassID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.BatchClassID.length;i++){" & vbCrLf
    Response.Write "    document.myform.BatchClassID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

    
    SoftID = Replace(SoftID, " ", "")
    Response.Write "<form method='POST' name='myform' action='Admin_Soft.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><b>�����޸�" & ChannelShortName & "����</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='center'>"
    Response.Write "      <td class='tdbg' valign='top'>"
    Response.Write "        <table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <tr>"
    Response.Write "            <td align='center'><b>" & ChannelShortName & "��Χ</b></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <input type='radio' name='BatchType' value='1' checked>ָ��" & ChannelShortName & "ID��<br>"
    Response.Write "              <input type='text' name='BatchSoftID' value='" & SoftID & "' size='28'><br>"
    Response.Write "              <input type='radio' name='BatchType' value='2'>ָ����Ŀ��" & ChannelShortName & "��<br>"
    Response.Write "              <select name='BatchClassID' size='2' multiple style='height:280px;width:180px;'>" & GetClass_Option(0, 0) & "</select><br><div align='center'>"
    Response.Write "      <input type='button' name='Submit' value='  ѡ��������Ŀ  ' onclick='SelectAll()'><br>"
    Response.Write "      <input type='button' name='Submit' value='ȡ��ѡ��������Ŀ' onclick='UnSelectAll()'></div></td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "      <td valign='top' align='left'><br>"
    
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>��������</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>�շ�ѡ��</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>�Զ���ѡ��</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyAuthor' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>����/�����̣�</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Author' type='text' id='Author' value='" & Trim(Session("Author")) & "' size='15' maxlength='30'> " & GetAuthorList("Admin", ChannelID, UserName)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyCopyFrom' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "��Դ��</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='CopyFrom' type='text' id='CopyFrom' value='" & Trim(Session("CopyFrom")) & "' size='15' maxlength='50'> " & GetCopyFromList("Admin", ChannelID)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifySoftType' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "���</td>"
    Response.Write "            <td><select name='SoftType' id='SoftType'>" & GetSoftType(1) & "</select></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifySoftLanguage' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "���ԣ�</td>"
    Response.Write "            <td><select name='SoftLanguage' id='SoftLanguage'>" & GetSoftLanguage(2) & "</select></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyCopyrightType' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>��Ȩ��ʽ��</td>"
    Response.Write "            <td><select name='CopyrightType' id='CopyrightType'>" & GetCopyrightType(2) & "</select></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyOperatingSystem' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "ƽ̨��</td>"
    Response.Write "            <td><input name='OperatingSystem' type='text' value='" & XmlText("Soft", "OperatingSystem", "Win9x/NT/2000/XP/") & "' size='50' maxlength='100'><br>" & GetOperatingSystemList & "</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyDecompressPassword' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>��ѹ���룺</td>"
    Response.Write "            <td><input name='DecompressPassword' type='text' value='' size='50' maxlength='100'></td>"
    Response.Write "          </tr>"
    Call ShowBatchCommon
    Response.Write "        </tbody>" & vbCrLf

    Call ShowTabs_Purview_Batch("����")
    Call ShowTabs_MyField_Batch

    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "  <br><b>˵����</b><br>1����Ҫ�����޸�ĳ�����Ե�ֵ������ѡ�������ĸ�ѡ��Ȼ�����趨����ֵ��<br>2��������ʾ������ֵ����ϵͳĬ��ֵ������ѡ" & ChannelShortName & "�����������޹�<br>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "  <p align='center'>"
    Response.Write "    <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "    <input name='Action' type='hidden' id='Action' value='DoBatch'>"
    Response.Write "    <input name='add' type='submit'  id='Add' value=' ִ�������� ' style='cursor:hand;'>&nbsp; "
    Response.Write "    <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p>"
    Response.Write "  <br>"
    Response.Write "</form>"
End Sub

Sub DoBatch()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    
    Dim rs, sql, BatchType, BatchSoftID, BatchClassID, rsField
    Dim Author, CopyFrom, SoftType, SoftLanguage, CopyrightType, OperatingSystem, DecompressPassword
    Dim Keyword, OnTop, Elite, Stars, Hits, UpdateTime, SkinID, TemplateID
    Dim InfoPurview, arrGroupID, InfoPoint, ChargeType, PitchTime, ReadTimes, DividePercent
    
    BatchType = PE_CLng(Trim(Request("BatchType")))
    BatchSoftID = Trim(Request.Form("BatchSoftID"))
    BatchClassID = FilterArrNull(Trim(Request.Form("BatchClassID")), ",")
    Author = Trim(Request.Form("Author"))
    CopyFrom = Trim(Request.Form("CopyFrom"))
    SoftType = Trim(Request.Form("SoftType"))
    SoftLanguage = Trim(Request.Form("SoftLanguage"))
    CopyrightType = Trim(Request.Form("CopyrightType"))
    OperatingSystem = Trim(Request.Form("OperatingSystem"))
    DecompressPassword = Trim(Request.Form("DecompressPassword"))
    Keyword = Trim(Request.Form("Keyword"))
    OnTop = Trim(Request.Form("OnTop"))
    Elite = Trim(Request.Form("Elite"))
    Stars = PE_CLng(Trim(Request.Form("Stars")))
    Hits = PE_CLng(Trim(Request.Form("Hits")))
    UpdateTime = PE_CDate(Trim(Request.Form("UpdateTime")))
    SkinID = PE_CLng(Trim(Request.Form("SkinID")))
    TemplateID = PE_CLng(Trim(Request.Form("TemplateID")))
    
    InfoPurview = PE_CLng(Trim(Request.Form("InfoPurview")))
    arrGroupID = ReplaceBadChar(Trim(Request.Form("GroupID")))
    InfoPoint = PE_CLng(Trim(Request.Form("InfoPoint")))
    ChargeType = PE_CLng(Trim(Request.Form("ChargeType")))
    PitchTime = PE_CLng(Trim(Request.Form("PitchTime")))
    ReadTimes = PE_CLng(Trim(Request.Form("ReadTimes")))
    DividePercent = PE_CLng(Trim(Request.Form("DividePercent")))
    If BatchType = 1 Then
        If IsValidID(BatchSoftID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����޸ĵ�" & ChannelShortName & "��ID</li>"
        End If
    Else
        If IsValidID(BatchClassID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����޸ĵ�" & ChannelShortName & "����Ŀ</li>"
        End If
    End If
    If Trim(Request("ModifyKeyword")) = "Yes" And Keyword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������" & ChannelShortName & "�ؼ���</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    If Author = "" Then Author = XmlText("BaseText", "DefAuthor", "����")
    If CopyFrom = "" Then CopyFrom = XmlText("BaseText", "DefCopyFrom", "��վԭ��")
    Keyword = "|" & ReplaceBadChar(Keyword) & "|"
    If OnTop = "Yes" Then
        OnTop = True
    Else
        OnTop = False
    End If
    If Elite = "Yes" Then
        Elite = True
    Else
        Elite = False
    End If
    
    Set rs = Server.CreateObject("adodb.recordset")
    If BatchType = 1 Then
        sql = "select * from PE_Soft where ChannelID=" & ChannelID & " and SoftID in (" & BatchSoftID & ")"
    Else
        sql = "select * from PE_Soft where ChannelID=" & ChannelID & " and ClassID in (" & BatchClassID & ")"
    End If
    rs.Open sql, Conn, 1, 3
    Do While Not rs.EOF
        If Trim(Request("ModifyAuthor")) = "Yes" Then rs("Author") = Author
        If Trim(Request("ModifyCopyFrom")) = "Yes" Then rs("CopyFrom") = CopyFrom
        If Trim(Request("ModifySoftType")) = "Yes" Then rs("SoftType") = SoftType
        If Trim(Request("ModifySoftLanguage")) = "Yes" Then rs("SoftLanguage") = SoftLanguage
        If Trim(Request("ModifyCopyrightType")) = "Yes" Then rs("CopyrightType") = CopyrightType
        If Trim(Request("ModifyOperatingSystem")) = "Yes" Then rs("OperatingSystem") = OperatingSystem
        If Trim(Request("ModifyDecompressPassword")) = "Yes" Then rs("DecompressPassword") = DecompressPassword
        If Trim(Request("ModifyInfoPoint")) = "Yes" Then rs("InfoPoint") = InfoPoint
        If Trim(Request("ModifyKeyword")) = "Yes" Then rs("Keyword") = Keyword
        If Trim(Request("ModifyOnTop")) = "Yes" Then rs("OnTop") = OnTop
        If Trim(Request("ModifyElite")) = "Yes" Then rs("Elite") = Elite
        If Trim(Request("ModifyStars")) = "Yes" Then rs("Stars") = Stars
        If Trim(Request("ModifyHits")) = "Yes" Then rs("Hits") = Hits
        If Trim(Request("ModifyUpdateTime")) = "Yes" Then rs("UpdateTime") = DateAdd("d", DateDiff("d", rs("UpdateTime"), UpdateTime), rs("UpdateTime"))
        If Trim(Request("ModifySkin")) = "Yes" Then rs("SkinID") = SkinID
        If Trim(Request("ModifyTemplate")) = "Yes" Then rs("TemplateID") = TemplateID
        If Trim(Request("ModifyInfoPurview")) = "Yes" Then
            rs("InfoPurview") = InfoPurview
            rs("arrGroupID") = arrGroupID
        End If
        If Trim(Request("ModifyInfoPoint")) = "Yes" Then rs("InfoPoint") = InfoPoint
        If Trim(Request("ModifyChargeType")) = "Yes" Then
            rs("ChargeType") = ChargeType
            rs("PitchTime") = PitchTime
            rs("ReadTimes") = ReadTimes
        End If
        If Trim(Request("ModifyDividePercent")) = "Yes" Then rs("DividePercent") = DividePercent

        Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-2 or ChannelID=" & ChannelID & "")
        If Not (rsField.BOF And rsField.EOF) Then
            Do While Not rsField.EOF
                If Trim(Request("Modify" & Trim(rsField("FieldName")))) = "Yes" Then
                    If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                        rs(Trim(rsField("FieldName"))) = Trim(Request(rsField("FieldName")))
                    End If
                End If
                rsField.MoveNext
            Loop
        End If
        Set rsField = Nothing

        rs("CreateTime") = rs("UpdateTime")
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Call ClearSiteCache(0)

    Call WriteSuccessMsg("�����޸�" & ChannelShortName & "���Գɹ���", "Admin_Soft.asp?ChannelID=" & ChannelID)
    Call CreateAllJS
End Sub


'******************************************************************************************
'����Ϊ���ù̶����Ƽ�������ʹ�õĺ�������ģ��ʵ�ֹ������ƣ��޸�ʱע��ͬʱ�޸ĸ�ģ�����ݡ�
'******************************************************************************************

Sub SetProperty()
    If SoftID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ѡ��" & ChannelShortName & "��</li>"
        Exit Sub
    End If
    If Action = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
        Exit Sub
    End If
    
    Dim sqlProperty, rsProperty, arrUser
    arrUser = ""
    If InStr(SoftID, ",") > 0 Then
        sqlProperty = "select * from PE_Soft where SoftID in (" & SoftID & ")"
    Else
        sqlProperty = "select * from PE_Soft where SoftID=" & SoftID
    End If
    Set rsProperty = Server.CreateObject("ADODB.Recordset")
    rsProperty.Open sqlProperty, Conn, 1, 3
    Do While Not rsProperty.EOF
        If CheckClassPurview(Action, rsProperty("ClassID")) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�� " & rsProperty("SoftName") & " û�в���Ȩ�ޣ�</li>"
        Else
            If FoundInArr(arrUser, rsProperty("Inputer"), ",") = False Then
                If arrUser = "" Then
                    arrUser = rsProperty("Inputer")
                Else
                    arrUser = arrUser & "," & rsProperty("Inputer")
                End If
            End If
            Select Case Action
            Case "SetOnTop"
                rsProperty("OnTop") = True
            Case "CancelOnTop"
                rsProperty("OnTop") = False
            Case "SetElite"
                rsProperty("Elite") = True
            Case "CancelElite"
                rsProperty("Elite") = False
            Case "SetPassed"
                If rsProperty("Status") < MyStatus And MyStatus = 3 Then
                    Conn.Execute ("update PE_User set UserExp=UserExp+" & rsProperty("PresentExp") & " where UserName='" & rsProperty("Inputer") & "'")
                End If
                rsProperty("Status") = MyStatus
                If MyStatus < 3 Or CheckLevel = 1 Then
                    rsProperty("Editor") = AdminName
                End If
            Case "CancelPassed", "Reject"
                If rsProperty("Status") = 3 Then
                    Conn.Execute ("update PE_User set UserExp=UserExp-" & rsProperty("PresentExp") & " where UserName='" & rsProperty("Inputer") & "'")
                End If
                If Action = "CancelPassed" Then
                    rsProperty("Status") = 0
                Else
                    rsProperty("Status") = -2
                End If
            End Select
            rsProperty("CreateTime") = rsProperty("UpdateTime")
            rsProperty.Update
        End If
        rsProperty.MoveNext
    Loop
    rsProperty.Close
    Set rsProperty = Nothing

    Call UpdateChannelData(ChannelID)
    Call UpdateUserData(0, arrUser, 0, 0)

    Call ClearSiteCache(0)
    Call WriteSuccessMsg("�����ɹ���", "Admin_Soft.asp?ChannelID=" & ChannelID)
    Call CreateAllJS
End Sub


'******************************************************************************************
'����Ϊ�ƶ�����Ŀ��ר��Ȳ���ʹ�õĺ�������ģ��ʵ�ֹ������ƣ��޸�ʱע��ͬʱ�޸ĸ�ģ�����ݡ�
'******************************************************************************************

Sub DoMoveToClass()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    
    Dim SoftType, BatchSoftID, BatchClassID
    Dim tChannelID, tClassID, tChannelDir, tUploadDir
    
    SoftType = PE_CLng(Trim(Request("SoftType")))
    BatchSoftID = Trim(Request.Form("BatchSoftID"))
    BatchClassID = FilterArrNull(Trim(Request.Form("BatchClassID")), ",")
    tChannelID = Trim(Request("tChannelID"))
    tClassID = Trim(Request("tClassID"))
    
    If SoftType = 1 Then
        If IsValidID(BatchSoftID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����ƶ���" & ChannelShortName & "��ID</li>"
        End If
    Else
        If IsValidID(BatchClassID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����ƶ���" & ChannelShortName & "����Ŀ</li>"
        End If
    End If
    If tChannelID = "" Then
        tChannelID = ChannelID
    Else
        tChannelID = PE_CLng(tChannelID)
        If tChannelID <> ChannelID Then
            If AdminPurview > 1 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<br><li>�Բ������Ȩ�޲�����</li>"
            Else
                Dim rsChannel
                Set rsChannel = Conn.Execute("select ChannelDir,UploadDir from PE_Channel where ChannelID=" & tChannelID & "")
                If rsChannel.BOF And rsChannel.EOF Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�Ҳ���Ŀ��Ƶ����</li>"
                Else
                    tChannelDir = rsChannel("ChannelDir")
                    tUploadDir = rsChannel("UploadDir")
                End If
                Set rsChannel = Nothing
            End If
        End If
    End If
    If tClassID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ŀ����Ŀ������ָ��Ϊ�ⲿ��Ŀ��</li>"
    Else
        tClassID = PE_CLng(tClassID)
        If tClassID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>Ŀ����Ŀ���������" & ChannelShortName & "</li>"
        End If
    End If
    If FoundErr = True Then Exit Sub
    
    Dim rsBatchMove, sqlBatchMove, SoftPath
    sqlBatchMove = "select S.SoftID,S.SoftPicUrl,S.DownloadUrl,S.UpdateTime,C.ParentDir,C.ClassDir from PE_Soft S left join PE_Class C on S.ClassID=C.ClassID"
    If SoftType = 1 Then
        sqlBatchMove = sqlBatchMove & " where S.ChannelID=" & ChannelID & " and S.SoftID in (" & BatchSoftID & ")"
    Else
        sqlBatchMove = sqlBatchMove & " where S.ChannelID=" & ChannelID & " and S.ClassID in (" & BatchClassID & ")"
    End If
    BatchSoftID = ""
    Set rsBatchMove = Conn.Execute(sqlBatchMove)
    Do While Not rsBatchMove.EOF
        SoftPath = HtmlDir & GetItemPath(StructureType, rsBatchMove("ParentDir"), rsBatchMove("ClassDir"), rsBatchMove("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsBatchMove("UpdateTime"), rsBatchMove("SoftID"))
        If fso.FileExists(Server.MapPath(SoftPath & FileExt_Item)) Then
            DelSerialFiles Server.MapPath(SoftPath & FileExt_Item)
        End If
        If tChannelID <> ChannelID Then
            Call MoveUpPic(rsBatchMove("SoftPicUrl"), tChannelDir)
            Call MoveUpFiles(rsBatchMove("DownloadUrl"), tChannelDir & "/" & tUploadDir)    '�ƶ��ϴ��ļ�
        End If
        If BatchSoftID = "" Then
            BatchSoftID = rsBatchMove("SoftID")
        Else
            BatchSoftID = BatchSoftID & "," & rsBatchMove("SoftID")
        End If
        rsBatchMove.MoveNext
    Loop
    rsBatchMove.Close
    Set rsBatchMove = Nothing
    If BatchSoftID <> "" Then
        Conn.Execute ("update PE_Soft set ChannelID=" & tChannelID & ",ClassID=" & tClassID & ",TemplateID=0,CreateTime=UpdateTime where SoftID in (" & BatchSoftID & ")")
    End If

    Call WriteSuccessMsg("�ɹ���ѡ����" & ChannelShortName & "�ƶ���Ŀ��Ƶ����Ŀ����Ŀ�У�", "Admin_Soft.asp?ChannelID=" & ChannelID & "")
    Call ClearSiteCache(0)
End Sub


Sub MoveUpPic(strFile, strTargetDir)
    On Error Resume Next
    Dim strTrueFile, strTrueDir
    If strFile = "" Or strTargetDir = "" Then Exit Sub
    
    If Not fso.FolderExists(Server.MapPath(InstallDir & strTargetDir)) Then fso.CreateFolder Server.MapPath(InstallDir & strTargetDir)
    
    If Left(strFile, 1) <> "/" And InStr(strFile, "://") <= 0 Then
        strTrueDir = Server.MapPath(InstallDir & strTargetDir & "/" & Left(strFile, InStrRev(strFile, "/")))
        If Not fso.FolderExists(strTrueDir) Then fso.CreateFolder strTrueDir
        strTrueFile = Server.MapPath(InstallDir & ChannelDir & "/" & strFile)
        If fso.FileExists(strTrueFile) Then
            fso.MoveFile strTrueFile, Server.MapPath(InstallDir & strTargetDir & "/" & strFile)
        End If
    End If
End Sub

Sub MoveUpFiles(strFiles, strTargetDir)
    On Error Resume Next
    Dim arrSoftUrls, strTrueFile, arrUrls, strTrueDir, iTemp
    If strFiles = "" Or strTargetDir = "" Then Exit Sub
    
    If Not fso.FolderExists(Server.MapPath(InstallDir & strTargetDir)) Then fso.CreateFolder Server.MapPath(InstallDir & strTargetDir)
    
    arrSoftUrls = Split(strFiles, "$$$")
    For iTemp = 0 To UBound(arrSoftUrls)
        arrUrls = Split(arrSoftUrls(iTemp), "|")
        If UBound(arrUrls) = 1 Then
            If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                strTrueDir = Server.MapPath(InstallDir & strTargetDir & "/" & Left(arrUrls(1), InStr(arrUrls(1), "/")))
                If Not fso.FolderExists(strTrueDir) Then fso.CreateFolder strTrueDir
                strTrueFile = Server.MapPath(InstallDir & ChannelDir & "/" & UploadDir & "/" & arrUrls(1))
                If fso.FileExists(strTrueFile) Then
                    fso.MoveFile strTrueFile, Server.MapPath(InstallDir & strTargetDir & "/" & arrUrls(1))
                End If
            End If
        End If
    Next
    
End Sub

'******************************************************************************************
'����Ϊɾ������ա���ԭ�Ȳ���ʹ�õĺ�������ģ��ʵ�ֹ������ƣ��޸�ʱע��ͬʱ�޸ĸ�ģ�����ݡ�
'******************************************************************************************

Sub Del()
    If SoftID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ѡ��" & ChannelShortName & "��</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel, SoftPath, arrUser
    arrUser = ""
    sqlDel = "select S.SoftID,S.SoftName,S.UpdateTime,S.CreateTime,S.Inputer,S.Status,S.Deleted,S.PresentExp,S.ClassID,C.ParentDir,C.ClassDir from PE_Soft S left join PE_Class C on S.ClassID=C.ClassID"
    If InStr(SoftID, ",") > 0 Then
        sqlDel = sqlDel & " where S.SoftID in (" & SoftID & ") order by S.SoftID"
    Else
        sqlDel = sqlDel & " where S.SoftID=" & SoftID
    End If
    Set rsDel = Server.CreateObject("ADODB.Recordset")
    rsDel.Open sqlDel, Conn, 1, 3
    Do While Not rsDel.EOF
        PurviewChecked = False
        ClassID = rsDel("ClassID")
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or (rsDel("Inputer") = UserName And rsDel("Status") = 0) Then
            PurviewChecked = True
        Else
            PurviewChecked = CheckClassPurview(Action, ClassID)
        End If
        
        If PurviewChecked = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ɾ�� <font color='red'>" & rsDel("SoftName") & "</font> ʧ�ܣ�ԭ��û�в���Ȩ�ޣ�</li>"
        Else
            If FoundInArr(arrUser, rsDel("Inputer"), ",") = True Then
                If arrUser = "" Then
                    arrUser = rsDel("Inputer")
                Else
                    arrUser = arrUser & "," & rsDel("Inputer")
                End If
            End If

            SoftPath = HtmlDir & GetItemPath(StructureType, rsDel("ParentDir"), rsDel("ClassDir"), rsDel("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsDel("UpdateTime"), rsDel("SoftID"))
            If fso.FileExists(Server.MapPath(SoftPath & FileExt_Item)) Then
                DelSerialFiles Server.MapPath(SoftPath & FileExt_Item)
            End If

            If rsDel("Status") = 3 Then
                Conn.Execute ("update PE_User set UserExp=UserExp-" & rsDel("PresentExp") & " where UserName='" & rsDel("Inputer") & "'")
            End If
            rsDel("Deleted") = True
            rsDel("CreateTime") = rsDel("UpdateTime")
            rsDel.Update
        End If
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing

    Call UpdateChannelData(ChannelID)
    Call UpdateUserData(0, arrUser, 0, 0)

    Call ClearSiteCache(0)
    Call WriteSuccessMsg("�����ɹ���", "Admin_Soft.asp?ChannelID=" & ChannelID)
    Call CreateAllJS
End Sub

Sub DelFile()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    If SoftID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ѡ��" & ChannelShortName & "��</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel, SoftPath
    SoftID = ReplaceBadChar(SoftID)
    sqlDel = "select S.SoftID,S.UpdateTime,C.ParentDir,C.ClassDir from PE_Soft S left join PE_Class C on S.ClassID=C.ClassID where S.SoftID in (" & SoftID & ") order by S.SoftID"
    Set rsDel = Conn.Execute(sqlDel)
    Do While Not rsDel.EOF
        SoftPath = HtmlDir & GetItemPath(StructureType, rsDel("ParentDir"), rsDel("ClassDir"), rsDel("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsDel("UpdateTime"), rsDel("SoftID"))
        If fso.FileExists(Server.MapPath(SoftPath & FileExt_Item)) Then
            DelSerialFiles Server.MapPath(SoftPath & FileExt_Item)
        End If
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    Conn.Execute ("update PE_Soft set CreateTime=UpdateTime where SoftID in (" & SoftID & ")")
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub ConfirmDel()
    If AdminPurview = 2 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    If SoftID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ѡ��" & ChannelShortName & "��</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel
    SoftID = ReplaceBadChar(SoftID)
    sqlDel = "select SoftPicUrl,DownloadUrl,VoteID from PE_Soft where SoftID in (" & SoftID & ")"
    Set rsDel = Conn.Execute(sqlDel)
    Do While Not rsDel.EOF
        Call DelUploadFiles(GetUploadFiles(rsDel("DownloadUrl"), rsDel("SoftPicUrl")))
        Conn.Execute ("delete from PE_Vote where ID=" & rsDel("VoteID") & "")
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    Conn.Execute ("delete from PE_Soft where SoftID in (" & SoftID & ")")
    Conn.Execute ("delete from PE_Comment where ModuleType=" & ModuleType & " and InfoID in (" & SoftID & ")")
    Conn.Execute ("delete from PE_ConsumeLog where ModuleType=" & ModuleType & " and InfoID in (" & SoftID & ")")
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub ClearRecyclebin()
    If AdminPurview = 2 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel
    SoftID = ""
    sqlDel = "select SoftID,SoftPicUrl,DownloadUrl,VoteID from PE_Soft where Deleted=" & PE_True & " and ChannelID=" & ChannelID
    Set rsDel = Conn.Execute(sqlDel)
    Do While Not rsDel.EOF
        If SoftID = "" Then
            SoftID = rsDel(0)
        Else
            SoftID = SoftID & "," & rsDel(0)
        End If
        Call DelUploadFiles(GetUploadFiles(rsDel("DownloadUrl"), rsDel("SoftPicUrl")))
        Conn.Execute ("delete from PE_Vote where ID=" & rsDel("VoteID") & "")
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    If SoftID <> "" Then
        Conn.Execute ("delete from PE_Soft where Deleted=" & PE_True & " and ChannelID=" & ChannelID & "")
        Conn.Execute ("delete from PE_Comment where ModuleType=" & ModuleType & " and InfoID in (" & SoftID & ")")
        Conn.Execute ("delete from PE_ConsumeLog where ModuleType=" & ModuleType & " and InfoID in (" & SoftID & ")")
    End If
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub Restore()
    If AdminPurview = 2 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    If SoftID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ѡ��" & ChannelShortName & "��</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel, arrUser
    arrUser = ""
    If InStr(SoftID, ",") > 0 Then
        sqlDel = "select * from PE_Soft where SoftID in (" & SoftID & ")"
    Else
        sqlDel = "select * from PE_Soft where SoftID=" & SoftID
    End If
    Set rsDel = Server.CreateObject("ADODB.Recordset")
    rsDel.Open sqlDel, Conn, 1, 3
    Do While Not rsDel.EOF
        If FoundInArr(arrUser, rsDel("Inputer"), ",") = True Then
            If arrUser = "" Then
                arrUser = rsDel("Inputer")
            Else
                arrUser = arrUser & "," & rsDel("Inputer")
            End If
        End If
        If rsDel("Status") = 3 Then
            Conn.Execute ("update PE_User set UserExp=UserExp+" & rsDel("PresentExp") & " where UserName='" & rsDel("Inputer") & "'")
        End If
        rsDel("Deleted") = False
        rsDel.Update
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    Call UpdateChannelData(ChannelID)
    Call UpdateUserData(0, arrUser, 0, 0)
    Call ClearSiteCache(0)
    Call WriteSuccessMsg("�����ɹ���", "Admin_Soft.asp?ChannelID=" & ChannelID)
    Call CreateAllJS
End Sub

Sub RestoreAll()
    If AdminPurview = 2 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel, arrUser
    arrUser = ""
    sqlDel = "select * from PE_Soft where Deleted=" & PE_True & " and ChannelID=" & ChannelID
    Set rsDel = Server.CreateObject("ADODB.Recordset")
    rsDel.Open sqlDel, Conn, 1, 3
    Do While Not rsDel.EOF
        If FoundInArr(arrUser, rsDel("Inputer"), ",") = True Then
            If arrUser = "" Then
                arrUser = rsDel("Inputer")
            Else
                arrUser = arrUser & "," & rsDel("Inputer")
            End If
        End If
        If rsDel("Status") = 3 Then
            Conn.Execute ("update PE_User set UserExp=UserExp+" & rsDel("PresentExp") & " where UserName='" & rsDel("Inputer") & "'")
        End If
        rsDel("Deleted") = False
        rsDel.Update
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    Call UpdateChannelData(ChannelID)
    Call UpdateUserData(0, arrUser, 0, 0)
    Call ClearSiteCache(0)
    Call WriteSuccessMsg("�����ɹ���", "Admin_Soft.asp?ChannelID=" & ChannelID)
    Call CreateAllJS
End Sub


Sub DelUploadFiles(strUploadFiles)
    On Error Resume Next
    If Trim(strUploadFiles) = "" Or ObjInstalled_FSO <> True Then Exit Sub
    Dim arrUploadFiles, strFileName, i
    arrUploadFiles = Split(strUploadFiles, "|")
    For i = 0 To UBound(arrUploadFiles)
        If Trim(arrUploadFiles(i)) <> "" Then
            strFileName = InstallDir & ChannelDir & "/" & UploadDir & "/" & arrUploadFiles(i)
            If fso.FileExists(Server.MapPath(strFileName)) Then
                fso.DeleteFile (Server.MapPath(strFileName))
            End If
        End If
    Next
End Sub

Function GetUploadFiles(DownloadUrls, SoftPicUrl)
    Dim arrDownloadUrls, arrUrls, iTemp, strUrls
    strUrls = ""
    If LCase(Left(SoftPicUrl, 13)) = "uploadSoftpic" Then
        strUrls = strUrls & SoftPicUrl
    End If
    arrDownloadUrls = Split(DownloadUrls, "$$$")
    For iTemp = 0 To UBound(arrDownloadUrls)
        arrUrls = Split(arrDownloadUrls(iTemp), "|")
        If UBound(arrUrls) = 1 Then
            If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                If strUrls <> "" Then
                    strUrls = strUrls & "|" & UploadDir & "/" & arrUrls(1)
                Else
                    strUrls = UploadDir & "/" & arrUrls(1)
                End If
            End If
        End If
    Next
    GetUploadFiles = strUrls
End Function

Function GetOperatingSystemList()
    Dim strOperatingSystemList, i
    
    strOperatingSystemList = "<script language = 'JavaScript'>" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "function ToSystem(addTitle){" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "    var str=document.myform.OperatingSystem.value;" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "    if (document.myform.OperatingSystem.value=="""") {" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "        document.myform.OperatingSystem.value=document.myform.OperatingSystem.value+addTitle;" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "    }else{" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "        if (str.substr(str.length-1,1)==""" & XmlText("Soft", "OperatingSystemEmblem", "/") & """){" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "            document.myform.OperatingSystem.value=document.myform.OperatingSystem.value+addTitle;" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "        }else{" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "            document.myform.OperatingSystem.value=document.myform.OperatingSystem.value+""" & XmlText("Soft", "OperatingSystemEmblem", "/") & """+addTitle;" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "        }" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "    }" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "    document.myform.OperatingSystem.focus();" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "}" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "</script>" & vbCrLf

    strOperatingSystemList = strOperatingSystemList & "<font color='#808080'>ƽ̨ѡ��"
    If IsArray(arrOperatingSystem) Then
        For i = 0 To UBound(arrOperatingSystem)
            If Trim(arrOperatingSystem(i)) <> "" Then
                strOperatingSystemList = strOperatingSystemList & "<a href=""javascript:ToSystem('" & arrOperatingSystem(i) & "')"">" & arrOperatingSystem(i) & "</a>" & XmlText("Soft", "OperatingSystemEmblem", "/")
            End If
        Next
    End If
    strOperatingSystemList = strOperatingSystemList & "</font>"
    GetOperatingSystemList = strOperatingSystemList
End Function

Function GetSoftPicUrl(SoftPicUrl)
    If LCase(Left(SoftPicUrl, Len("UploadSoftPic"))) = "uploadsoftpic" Then
        GetSoftPicUrl = InstallDir & ChannelDir & "/" & SoftPicUrl
    Else
        GetSoftPicUrl = SoftPicUrl
    End If
End Function


Function CheckDownloadUrl(ByVal strUrl)
    On Error Resume Next
    Dim arrDownloadUrls, arrUrls, iTemp, DownloadUrl
    CheckDownloadUrl = True
    If InStr(strUrl, "@@@") > 0 Then
       CheckDownloadUrl = True
        Exit Function
    End If
    If Trim(strUrl) = "" Or IsNull(strUrl) Then
        CheckDownloadUrl = False
        Exit Function
    End If
    arrDownloadUrls = Split(strUrl, "$$$")
    For iTemp = 0 To UBound(arrDownloadUrls)
        arrUrls = Split(arrDownloadUrls(iTemp), "|")
        If UBound(arrUrls) >= 1 Then
            DownloadUrl = arrUrls(1)
            If DownloadUrl = "" Then
                CheckDownloadUrl = False
                Exit For
            End If
            If Left(DownloadUrl, 1) <> "/" And InStr(DownloadUrl, "://") <= 0 Then
            DownloadUrl = Replace(DownloadUrl,  "&nbsp;", " ")
                If Not fso.FileExists(Server.MapPath(InstallDir & ChannelDir & "/" & UploadDir & "/" & DownloadUrl)) Then
                    CheckDownloadUrl = False
                    Exit For
                End If
            End If
        Else
            CheckDownloadUrl = False
            Exit For
        End If
    Next
End Function

Function GetSoftType(SoftType)
    If IsArray(arrSoftType) = False Then Exit Function
    
    Dim strTemp, i
    For i = 0 To UBound(arrSoftType)
        If Trim(arrSoftType(i)) <> "" Then
            strTemp = strTemp & "<option value='" & arrSoftType(i) & "'"
            If Trim(SoftType) = arrSoftType(i) Then strTemp = strTemp & " selected"
            strTemp = strTemp & ">" & arrSoftType(i) & "</option>"
        End If
    Next
    GetSoftType = strTemp
End Function

Function GetSoftLanguage(SoftLanguage)
    If IsArray(arrSoftLanguage) = False Then Exit Function
    
    Dim strTemp, i
    For i = 0 To UBound(arrSoftLanguage)
        If Trim(arrSoftLanguage(i)) <> "" Then
            strTemp = strTemp & "<option value='" & arrSoftLanguage(i) & "'"
            If Trim(SoftLanguage) = arrSoftLanguage(i) Then strTemp = strTemp & " selected"
            strTemp = strTemp & ">" & arrSoftLanguage(i) & "</option>"
        End If
    Next
    GetSoftLanguage = strTemp
End Function

Function GetCopyrightType(CopyrightType)
    If IsArray(arrCopyrightType) = False Then Exit Function
    
    Dim strTemp, i
    For i = 0 To UBound(arrCopyrightType)
        If Trim(arrCopyrightType(i)) <> "" Then
            strTemp = strTemp & "<option value='" & arrCopyrightType(i) & "'"
            If Trim(CopyrightType) = arrCopyrightType(i) Then strTemp = strTemp & " selected"
            strTemp = strTemp & ">" & arrCopyrightType(i) & "</option>"
        End If
    Next
    GetCopyrightType = strTemp
End Function





'=================================================
'��������SaveModifyDownError
'��  �ã������޸ĺ�����ش�����Ϣ
'=================================================
Sub SaveModifyDownError()
    If AdminPurview > 1 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    Dim rsDownError, sqlDownError
    Dim strUrlName, iChannelID, iUrlID, iInfoID, strDownloadUrl
    
    iChannelID = Request("ChannelID")
    strUrlName = Request.Form("UrlName")
    iUrlID = PE_CLng(Request.Form("UrlID"))
    iInfoID = PE_CLng(Request.Form("InfoID"))

    rsDownError = Server.CreateObject("ADODB.RecordSet")
    sqlDownError = "select DownloadUrl from PE_Soft where ChannelID=" & iChannelID & " And SoftID = " & iInfoID
    Set rsDownError = Server.CreateObject("ADODB.RecordSet")
   
    rsDownError.Open sqlDownError, Conn, 1, 3
    If rsDownError.BOF And rsDownError.EOF Then
    Response.Write "δ�ҵ�������ƥ��Ĳ��������ݡ�"
    Else
    strDownloadUrl = UpdateDownloadUrl(rsDownError("DownloadUrl"), iUrlID, strUrlName)
   'Response.Write "strDownloadUrl=" & strDownloadUrl
    'Response.Write "strUrlName=" & strUrlName
    'Response.Write "iUrlID=" & iUrlID
    'Exit Sub
    
     Conn.Execute ("update PE_Soft set DownloadUrl='" & strDownloadUrl & "' where SoftID=" & iInfoID & "")
    'Conn.Execute ("update PE_Soft set DownloadUrl=" & strDownloadUrl & " where SoftID=" & iInfoID & "")
   
     Response.Redirect "Admin_Soft.asp?ChannelID=" & ChannelID & "&action=DownError"
    rsDownError.Close
    Set rsDownError = Nothing
    End If
End Sub

'=================================================
'��������GetUrlName
'��  �ã�ȡ�����ص�ַ�о���ĳ��URLID��������ַ
'��  ����DownloadUrls  ----���ص�ַ
'        UrlID ----�����ı��ID
'=================================================
Function GetUrlName(DownloadUrls, ByVal UrlID)
     Dim DownloadUrl, arrDownloadUrls, arrUrls, iTemp
     
     If DownloadUrls = "" Or UrlID = "" Or IsNull(DownloadUrls) Then
        GetUrlName = ""
        Exit Function
     End If
    
    iTemp = UrlID - 1
    arrDownloadUrls = Split(DownloadUrls, "$$$")
    If UBound(arrDownloadUrls) >= iTemp Then
        arrUrls = Split(arrDownloadUrls(iTemp), "|")
        If UBound(arrUrls) >= 1 Then
            If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                DownloadUrl = InstallDir & ChannelDir & "/" & UploadDir & "/" & arrUrls(1)
            Else
                DownloadUrl = GetFirstSeparatorToEnd(arrDownloadUrls(iTemp), "|")
            End If
        End If
    End If
    
    If DownloadUrl = "" Or DownloadUrl = "http://" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�Ҳ�����Ч���ص�ַ��</li>"
        
        Exit Function
    End If
    GetUrlName = DownloadUrl
End Function
'=================================================
'��������UpdateDownloadUrl
'��  �ã���UrlName�������ص�ַ�о���ĳ��UrlID��������ַ
'��  ����DownloadUrls  ----���ص�ַ
'        UrlID ----�����ı��ID
'        UrlName---����
'����ֵ���滻���º�����ص�ַ
'=================================================
Function UpdateDownloadUrl(DownloadUrls, ByVal UrlID, ByVal UrlName)
     Dim iTemp, arrDownloadUrls, strDownloadUrl
     arrDownloadUrls = Split(DownloadUrls, "$$$")
     
     If UrlID > 0 And UrlID < UBound(arrDownloadUrls) Then
     For iTemp = 0 To UBound(arrDownloadUrls)
          If iTemp = UrlID Then
              strDownloadUrl = Replace(DownloadUrls, GetUrlName(DownloadUrls, iTemp), UrlName)
          End If
     Next
     Else
     strDownloadUrl = Replace(DownloadUrls, GetUrlName(DownloadUrls, 1), UrlName)
     End If
     UpdateDownloadUrl = strDownloadUrl
End Function
'=================================================
'��������DownError
'���ã� ���ش�����Ϣ������
'������ ��
'=================================================
Sub DownError()
    Dim rsDownErrorList, sqlDownErrorList, UrlName
    Dim rsUrl, sqlurl
    Dim rs, sql, imgUrl
    Dim iCount, strKeyword
    If AdminPurview > 1 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    
    strKeyword = Request.Form("keyword")
    
    Call ShowJS_DownError
    Response.Write "<br>" & vbCrLf
    Response.Write "<form name='myform' method='Post' action='Admin_Soft.asp' onsubmit=""return confirm('ȷ��Ҫɾ��ѡ������Ϣ��');"">"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='30'><strong>ѡ��</strong></td>"
    Response.Write "    <td width='30' height='22'><strong>���</strong></td>"
    Response.Write "    <td height='22'><strong>�����</strong></td>"
    Response.Write "    <td width='185' height='22'><strong>���ص�ַ</strong></td>"
    Response.Write "    <td width='60' height='22'><strong>�����˴�</strong></td>"
    Response.Write "    <td  height='22'><strong>�� ��</strong></td>"
    Response.Write "  </tr>"

    Set rsDownErrorList = Server.CreateObject("Adodb.RecordSet")
    sqlDownErrorList = "select D.ErrorID,D.ChannelID,D.InfoID,D.UrlID,D.ErrorTimes,S.SoftID,S.ChannelID,S.SoftName,S.DownloadUrl from PE_DownError D "
    sqlDownErrorList = sqlDownErrorList & " left join PE_Soft S on D.InfoID=S.SoftID where D.ChannelID=" & ChannelID & ""
    If strKeyword <> "" Then
            sqlDownErrorList = sqlDownErrorList & " and D.InfoID In (select SoftID from PE_Soft where SoftName like '%" & strKeyword & "%' )"
    End If
    sqlDownErrorList = sqlDownErrorList & " order by D.ErrorTimes Desc"
    rsDownErrorList.Open sqlDownErrorList, Conn, 1, 1
    If rsDownErrorList.BOF And rsDownErrorList.EOF Then
        rsDownErrorList.Close
        Set rsDownErrorList = Nothing
        Response.Write "  <tr class='tdbg'><td colspan='6' align='center'><br>û���κ����ش�����Ϣ��<br><br></td></tr></Table>"
        Exit Sub
    End If

    totalPut = rsDownErrorList.RecordCount
    If CurrentPage < 1 Then
        CurrentPage = 1
    End If
    If (CurrentPage - 1) * MaxPerPage > totalPut Then
        If (totalPut Mod MaxPerPage) = 0 Then
            CurrentPage = totalPut \ MaxPerPage
        Else
            CurrentPage = totalPut \ MaxPerPage + 1
        End If
    End If
    If CurrentPage > 1 Then
        If (CurrentPage - 1) * MaxPerPage < totalPut Then
            rsDownErrorList.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If
    
    'If InStr(rsDownErrorList("DownloadUrl"),"@@@") > 0 Then

    
    Do While Not rsDownErrorList.EOF
        Response.Write " <tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "    <td><input name='ErrorID' type='checkbox' id='ErrorID' value='" & rsDownErrorList("ErrorID") & "'"
        Response.Write " onclick='unselectall()'></td>"
        Response.Write "    <td>" & rsDownErrorList("ErrorID") & "</td>"
        Response.Write "    <td>" & GetSubStr(rsDownErrorList("SoftName"), 40, True) & "</td>"
        '�Ƿ�Ϊ����ģʽ
        If InStr(rsDownErrorList("DownloadUrl"), "@@@") > 0 Then
            sql = "select * from PE_DownServer where ServerID=" & rsDownErrorList("UrlID")
            Set rs = Server.CreateObject("ADODB.Recordset")
            rs.Open sql, Conn, 1, 3
            If Not rs.BOF And Not rs.EOF Then
                imgUrl = rs("ServerUrl")
            End If
            UrlName = Trim(Replace(rsDownErrorList("DownloadUrl"), "@@@", ""))
            imgUrl = imgUrl & UrlName
            Response.Write "    <td>" & imgUrl & "</td>"
            Response.Write "    <td>" & rsDownErrorList("ErrorTimes") & "</td>"
            Response.Write "<td>"
            Response.Write "<a href=" & imgUrl & ">����</a>&nbsp;&nbsp;"
            Response.Write "�޸�&nbsp;&nbsp;"
        Else
            Response.Write "    <td>" & GetUrlName(rsDownErrorList("DownloadUrl"), rsDownErrorList("UrlID")) & "</td>"
            Response.Write "    <td>" & rsDownErrorList("ErrorTimes") & "</td>"
            Response.Write "<td>"

            UrlName = GetUrlName(rsDownErrorList("DownloadUrl"), rsDownErrorList("UrlID"))
            Response.Write "<a href=" & UrlName & ">����</a>&nbsp;&nbsp;"
            Response.Write "<a href='Admin_Soft.asp?action=ModifyDownError&ChannelID=" & ChannelID & "&ErrorID=" & rsDownErrorList("ErrorID") & "'>�޸�</a>&nbsp;&nbsp;"
        End If
        Response.Write "<a href='Admin_Soft.asp?action=DelDownError&ChannelID=" & ChannelID & "&ErrorID=" & rsDownErrorList("ErrorID") & "&InfoID=" & rsDownErrorList("InfoID") & "' onClick=""return confirm('ȷ��Ҫɾ�������ش�����Ϣ��');"">ɾ��</a>"
        Response.Write "</td>"
        Response.Write "</tr>"
        iCount = iCount + 1
        If iCount >= MaxPerPage Then Exit Do
        rsDownErrorList.MoveNext
    Loop
    rsDownErrorList.Close
    Set rsDownErrorList = Nothing

    Response.Write "</table>  "
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'> ѡ�б�ҳ��ʾ�����д�����Ϣ</td>"
    Response.Write "    <td><input name='action' type='hidden' id='action' value='DelDownError'>"
    Response.Write "    <td><input name='ChannelID' type='hidden' id='ChannelID' value=" & ChannelID & ">"
    Response.Write "    <input name='Submit' type='submit' id='Submit' value='ɾ��ѡ�е����ش�����Ϣ'>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<input type='submit' name='Submit2' value='ɾ����Ƶ��ȫ�����ش�����Ϣ' onClick=""document.myform.action.value='DelAllDownError'"">"
    Response.Write "  </tr>"
    Response.Write "</table></form>"
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "��������Ϣ", True)
    Response.Write "<br><table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "<tr class='tdbg'><td width='100' align='right'><strong>������Ϣ������</td>"
    Response.Write "<td><table border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "<form method='Post' name='SearchForm' action='Admin_Soft.asp?action=DownError'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
   ' Response.Write "<input name='TypeSelect' type='hidden' id='TypeSelect' value='DownError'>"
    Response.Write "<tr><td height='28' align='center'>"
    Response.Write "<select name='Field' size='1'><option value='name' selected>�����</option></select>"
    Response.Write "<input type='text' name='keyword' size='20' value='"
    If strKeyword <> "" Then
        Response.Write strKeyword
    Else
        Response.Write "�����������"
    End If
    Response.Write "' maxlength='50'>"
    Response.Write "<input type='submit' name='Submit' value='����'>"
    Response.Write "</td></tr></form></table></td></tr></table>"
End Sub

Sub ModifyDownError()
    Dim ErrorID, sqlDownErrorUrl, rsDownErrorUrl
    Dim rsDownError, sqlDownError, rs, sql
    Dim strUrlName
    ErrorID = PE_CLng(Trim(Request("ErrorID")))
    If ErrorID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ����ش�����ϢID</li>"
        Exit Sub
    End If
    sqlDownError = "Select * from PE_DownError where ErrorID=" & ErrorID
    Set rsDownError = Server.CreateObject("Adodb.RecordSet")
    rsDownError.Open sqlDownError, Conn, 1, 3
    If rsDownError.BOF And rsDownError.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����ڴ����ش�����Ϣ��</li>"
    Else
        Response.Write "<form method='post' action='Admin_Soft.asp?action=SaveModifyDownError' name='myform'>"
        Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Response.Write "    <tr class='title'> "
        Response.Write "      <td height='22' colspan='2'> <div align='center'><font size='2'><strong>�� �� �� �� �� �� ��ַ</strong></font></div></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr> "
        Response.Write "      <td width='100%' class='tdbg' align='center'><strong>���ص�ַ ��</strong><input name='UrlName' type='text' size=60  value='"
        If rsDownError("UrlID") <> "" Then
            sqlDownErrorUrl = "Select * from PE_Soft where ErrorTimes>0 and SoftID=" & rsDownError("InfoID")
            Set rsDownErrorUrl = Conn.Execute(sqlDownErrorUrl)
            If rsDownErrorUrl.BOF And rsDownErrorUrl.EOF Then
                Response.Write "û���ҵ��������������ص�ַ"
            Else
           '����ģʽ
                If InStr(rsDownErrorUrl("DownloadUrl"), "@@@") > 0 Then
                    sql = "select * from PE_DownServer where ServerID=" & rsDownError("UrlID")
                    Set rs = Server.CreateObject("ADODB.Recordset")
                    rs.Open sql, Conn, 1, 3
                    If Not rs.BOF And Not rs.EOF Then
                    strUrlName = rs("ServerUrl")
                    End If
                    strUrlName = strUrlName & Trim(Replace(rsDownErrorUrl("DownloadUrl"), "@@@", ""))
                Else
                    strUrlName = GetUrlName(rsDownErrorUrl("DownloadUrl"), rsDownError("UrlID"))
                End If
                Response.Write strUrlName
            End If
        End If
        Response.Write "'>"
        Response.Write "</td>"
        Response.Write "</tr>"
        Response.Write "    <tr>"
        Response.Write "      <td colspan='2' align='center' class='tdbg'>"
        Response.Write "      <input name='ChannelID' type='hidden' id='ChannelIE' value=" & ChannelID & ">"
        Response.Write "      <input name='strUrlName' type='hidden' id='strUrlName' value=" & strUrlName & ">"
        Response.Write "      <input name='InfoID' type='hidden' id='InfoID' value=" & rsDownError("InfoID") & ">"
        Response.Write "      <input name='UrlID' type='hidden' id='UrlID' value=" & rsDownError("UrlID") & ">"
        Response.Write "      <input  type='submit' name='Submit' value='�����޸Ľ��'  style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Soft.asp?ChannelID=" & ChannelID & "&action=DownError'"" style='cursor:hand;'></td>"
        Response.Write "    </tr>"
        Response.Write "  </table>"
        Response.Write "</form>"
    End If
    rsDownError.Close
    Set rsDownError = Nothing
    rsDownErrorUrl.Close
    Set rsDownErrorUrl = Nothing
End Sub

 Sub DelDownError()
    Dim sqlDelDownError, rsDelDownError, ErrTimes, Times, Num
    Dim ErrorID, SoftID
    ErrorID = Trim(Request("ErrorID"))
    SoftID = PE_CLng(Trim(Request("InfoID")))
    If IsValidID(ErrorID) = False Then
        ErrorID = ""
    End If
    If ErrorID = "" Or SoftID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫɾ�������ش�����ϢID</li>"
        Exit Sub
    End If
   
    sqlDelDownError = "select S.SoftID,D.ErrorID,D.InfoID,D.ErrorTimes as Times from PE_Soft S left join PE_DownError D on S.SoftID=D.InfoID"
    If InStr(ErrorID, ",") > 0 Then
        sqlDelDownError = sqlDelDownError & " where D.ErrorID in (" & ErrorID & ") order by D.ErrorID"
    Else
        sqlDelDownError = sqlDelDownError & " where D.ErrorID=" & ErrorID
    End If
    Set rsDelDownError = Server.CreateObject("ADODB.Recordset")
    rsDelDownError.Open sqlDelDownError, Conn, 1, 3
    Do While Not rsDelDownError.EOF
        Conn.Execute ("update PE_Soft set ErrorTimes=ErrorTimes-" & rsDelDownError("Times") & " where  SoftID=" & rsDelDownError("SoftID") & "")
        Conn.Execute ("delete from PE_DownError where ErrorID=" & rsDelDownError("ErrorID") & "")
        rsDelDownError.MoveNext
    Loop
    rsDelDownError.Close
    Set rsDelDownError = Nothing
    Call CloseConn
   Response.Redirect "Admin_Soft.asp?ChannelID=" & ChannelID & "&action=DownError"
End Sub

Sub DelAllDownError()
    Conn.Execute ("delete from PE_DownError where ChannelID=" & ChannelID)
     Conn.Execute ("update PE_Soft set ErrorTimes=0 where  ChannelID=" & ChannelID)
    Call CloseConn
    Response.Redirect "Admin_Soft.asp?ChannelID=" & ChannelID & "&action=DownError"
End Sub

Sub ShowJS_DownError()
    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "function unselectall(){" & vbCrLf
    Response.Write "    if(document.myform.chkAll.checked){" & vbCrLf
    Response.Write " document.myform.chkAll.checked = document.myform.chkAll.checked&0;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckAll(form){" & vbCrLf
    Response.Write "  for (var i=0;i<form.elements.length;i++){" & vbCrLf
    Response.Write "    var e = form.elements[i];" & vbCrLf
    Response.Write "    if (e.Name != 'chkAll'&&e.disabled==false)" & vbCrLf
    Response.Write "       e.checked = form.chkAll.checked;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "function ConfirmDel(){" & vbCrLf
    Response.Write " if(document.myform.Action.value=='Del'){" & vbCrLf
    Response.Write "     if(confirm('ȷ��Ҫɾ��ѡ�е����ش�����Ϣ��'))" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "     else" & vbCrLf
    Response.Write "         return false;" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
End Sub

Sub ShowDownloadUrls(DownloadUrls)
    Dim arrDownloadUrls, arrUrls, iTemp
    DownloadUrls = Replace(DownloadUrls,  "&nbsp;", " ")
    arrDownloadUrls = Split(DownloadUrls, "$$$")
    For iTemp = 0 To UBound(arrDownloadUrls)
        arrUrls = Split(arrDownloadUrls(iTemp), "|")
        If UBound(arrUrls) >= 1 Then
            If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                Response.Write arrUrls(0) & "��<a href='" & InstallDir & ChannelDir & "/" & UploadDir & "/" & arrUrls(1) & "'>" & arrUrls(1) & "</a><br>"
            Else
                Response.Write arrUrls(0) & "��<a href='" & GetFirstSeparatorToEnd(arrDownloadUrls(iTemp), "|") & "'>" & GetFirstSeparatorToEnd(arrDownloadUrls(iTemp), "|") & "</a><br>"
            End If
        End If
    Next
End Sub
%>
