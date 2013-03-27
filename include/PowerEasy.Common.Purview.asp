<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Function CheckClassPurview(Action, ClassID)
    Dim PurviewType, PurviewChecked, CheckParentPath, i
    PresentExp = 0
    If ClassID = "" Or IsNull(ClassID) Or Not IsNumeric(ClassID) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>δָ��������Ŀ������ָ������Ŀ������˲�����</li>"
    Else
        PurviewType = LCase(Action)
        ClassID = PE_CLng(ClassID)
        Select Case ClassID
        Case 0
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ָ������Ŀ������˲�����</li>"
        Case -1
            If AdminPurview = 2 And AdminPurview_Channel >= 3 And PurviewType <> "show" And PurviewType <> "preview" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
            Else
                ClassName = "��ָ���κ���Ŀ"
                Depth = -1
                ParentPath = ""
            End If
        Case Else
            Set tClass = Conn.Execute("select ClassName,ClassType,Depth,ParentID,ParentPath,Child,EnableAdd,PresentExp from PE_Class where ClassID=" & ClassID)
            If tClass.BOF And tClass.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�Ҳ���ָ������Ŀ��</li>"
            Else
                ClassName = tClass("ClassName")
                Depth = tClass("Depth")
                ParentPath = tClass("ParentPath")
                ParentID = tClass("ParentID")
                Child = tClass("Child")
                PresentExp = tClass("PresentExp")

                If PurviewType = "saveadd" Or PurviewType = "savemodify" Or PurviewType = "input" Then
                    If Child > 0 And tClass("EnableAdd") = False Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>ָ������Ŀ���������" & ChannelShortName & "</li>"
                    End If
                    If tClass("ClassType") = 2 Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>����ָ��Ϊ�ⲿ��Ŀ</li>"
                    End If
                End If
                If AdminPurview = 2 And AdminPurview_Channel = 3 Then
                    If ParentID > 0 Then
                        CheckParentPath = ParentPath & "," & ClassID
                    Else
                        CheckParentPath = ClassID
                    End If
                    Select Case PurviewType
                    Case "show", "preview"
                        PurviewChecked = CheckPurview_Class(arrClass_View, CheckParentPath)
                    Case "saveadd", "savemodify", "input"
                        PurviewChecked = CheckPurview_Class(arrClass_Input, CheckParentPath)
                    Case "setpassed", "cancelpassed", "check"
                        PurviewChecked = CheckPurview_Class(arrClass_Check, CheckParentPath)
                    Case Else
                        PurviewChecked = CheckPurview_Class(arrClass_Manage, CheckParentPath)
                    End Select
                    If PurviewChecked = False Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>�Բ�����û����Ŀ����Ӧ����Ȩ�ޣ�</li>"
                    End If
                End If
            End If
            Set tClass = Nothing
        End Select
    End If
    If FoundErr = True Then
        CheckClassPurview = False
    Else
        CheckClassPurview = True
    End If
End Function

Function CheckClassMaster(AllMaster, MasterName)
    If IsNull(AllMaster) Or AllMaster = "" Or MasterName = "" Then
        CheckClassMaster = False
        Exit Function
    End If
    CheckClassMaster = False
    If InStr(AllMaster, "|") > 0 Then
        Dim arrMaster, i
        arrMaster = Split(AllMaster, "|")
        For i = 0 To UBound(arrMaster)
            If Trim(arrMaster(i)) = MasterName Then
                CheckClassMaster = True
                Exit For
            End If
        Next
    Else
        If AllMaster = MasterName Then
            CheckClassMaster = True
        End If
    End If
End Function

'**************************************************
'��������CheckPurview_Other
'��  �ã�����Ȩ��������
'��  ����AllPurviews ---- Ҫ�Ƚ�����
'        strPurview ---- �Ƚ��ַ�
'����ֵ��True  ---- ����
'**************************************************
Function CheckPurview_Other(AllPurviews, strPurview)
    If IsNull(AllPurviews) Or AllPurviews = "" Or strPurview = "" Then
        CheckPurview_Other = False
        Exit Function
    End If
    CheckPurview_Other = False
    If InStr(AllPurviews, ",") > 0 Then
        Dim arrPurviews, i
        arrPurviews = Split(AllPurviews, ",")
        For i = 0 To UBound(arrPurviews)
            If Trim(arrPurviews(i)) = Trim(strPurview) Then
                CheckPurview_Other = True
                Exit For
            End If
        Next
    Else
        If Trim(AllPurviews) = Trim(strPurview) Then
            CheckPurview_Other = True
        End If
    End If
End Function

'**************************************************
'��������CheckPurview_Class
'��  �ã���ĿȨ��������
'��  ����str1 ---- Ҫ�Ƚ�����1
'        str2 ---- Ҫ�Ƚ�����2
'����ֵ��True  ---- ����
'**************************************************
Function CheckPurview_Class(str1, str2)
    Dim arrTemp, arrTemp2, i, j
    CheckPurview_Class = False
    If IsNull(str1) Or IsNull(str2) Or str1 = "" Or str2 = "" Then
        Exit Function
    End If
    arrTemp = Split(str1 & ",", ",")
    arrTemp2 = Split(str2 & ",", ",")
    For i = 0 To UBound(arrTemp)
        For j = 0 To UBound(arrTemp2)
            If Trim(arrTemp2(j)) <> "" And Trim(arrTemp(i)) <> "" And Trim(arrTemp2(j)) = Trim(arrTemp(i)) Then
                CheckPurview_Class = True
                Exit Function
            End If
        Next
    Next
End Function

'**************************************************
'��������CheckPurview_Channel
'��  �ã�Ƶ�������û�����
'��  ����ChannelPurview ---- Ƶ����Ȩ�ޡ�0Ϊ����Ƶ�����κ��˿����������1Ϊ��֤Ƶ�����οͲ��������
'        ChannelArrGroupID ---- ������ʵ��û���
'        GroupID ---- �û��������û���
'����ֵ��True  ---- ��Ȩ�޷���
'**************************************************
Function CheckPurview_Channel(ChannelPurview, ChannelArrGroupID, UserLogined, GroupID)
    ChannelPurview = PE_CLng(ChannelPurview)
    CheckPurview_Channel = False
    If ChannelPurview = 0 Then
        CheckPurview_Channel = True
    Else
        If UserLogined = True Then
            If FoundInArr(ChannelArrGroupID, GroupID, ",") = True Then
                CheckPurview_Channel = True
            End If
        End If
    End If
End Function


'==================================================
'��������WriteEntry
'��  �ã�����Ϣд����־
'��  ����LogType��1--��Ҫ����  2--ϵͳ����  3--Ƶ������  4--��¼ʧ��  5--�������
'        UserName    ------ ������
'        LogContent  ------ ������Ϣ
'
'˵  ����LogType��ChannelID��ȡֵ
'==================================================
Public Sub WriteEntry(LogType, UserName, LogContent)
    Dim sqlLog, rsLog
    sqlLog = "select top 1 * from PE_Log"
    Set rsLog = Server.CreateObject("Adodb.RecordSet")
    rsLog.Open sqlLog, Conn, 1, 3
    rsLog.addnew
    rsLog("LogType") = LogType
    rsLog("ChannelID") = 0      '�����ֶ�
    rsLog("UserName") = UserName
    rsLog("LogContent") = LogContent
    rsLog("LogTime") = Now()
    rsLog("UserIP") = UserTrueIP
    rsLog("ScriptName") = GetScriptName()
    rsLog("PostString") = GetPostString()
    rsLog.Update
    rsLog.Close
    Set rsLog = Nothing
End Sub
'**************************************************
'��������GetScriptName
'��  �ã���ȡ�������ļ�Ŀ¼
'����ֵ���ļ�Ŀ¼
'**************************************************
Function GetScriptName()
    Dim ScriptName
    ScriptName = Trim(Request.ServerVariables("SCRIPT_NAME"))
    If InStr(ScriptName, "?") > 0 Then
        ScriptName = Left(ScriptName, InStr(ScriptName, "?"))
    End If
    GetScriptName = ScriptName
End Function
'**************************************************
'��������GetPostString
'��  �ã��ж϶Է������� Form ���� QueryString
'����ֵ���Է����ʵ�����
'**************************************************
Function GetPostString()
    Dim PostString, PostItem
    PostString = ""
    If Request.Form <> "" Then
        PostString = PostString & "Request.Form��"
        For Each PostItem In Request.Form
            PostString = PostString & PostItem & "=" & Request.Form(PostItem) & "&"
        Next
        If Right(PostString, 1) = "&" Then PostString = Left(PostString, Len(PostString) - 1)
    End If
    If Request.QueryString <> "" Then
        If PostString <> "" Then PostString = PostString & vbCrLf
        PostString = PostString & "Request.QueryString��"
        For Each PostItem In Request.QueryString
            PostString = PostString & PostItem & "=" & Request.QueryString(PostItem) & "&"
        Next
        If Right(PostString, 1) = "&" Then PostString = Left(PostString, Len(PostString) - 1)
    End If
    GetPostString = PostString
End Function

%>
