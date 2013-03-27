<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Function CheckClassPurview(Action, ClassID)
    Dim PurviewType, PurviewChecked, CheckParentPath, i
    PresentExp = 0
    If ClassID = "" Or IsNull(ClassID) Or Not IsNumeric(ClassID) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>未指定所属栏目，或者指定的栏目不允许此操作！</li>"
    Else
        PurviewType = LCase(Action)
        ClassID = PE_CLng(ClassID)
        Select Case ClassID
        Case 0
            FoundErr = True
            ErrMsg = ErrMsg & "<li>指定的栏目不允许此操作！</li>"
        Case -1
            If AdminPurview = 2 And AdminPurview_Channel >= 3 And PurviewType <> "show" And PurviewType <> "preview" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
            Else
                ClassName = "不指定任何栏目"
                Depth = -1
                ParentPath = ""
            End If
        Case Else
            Set tClass = Conn.Execute("select ClassName,ClassType,Depth,ParentID,ParentPath,Child,EnableAdd,PresentExp from PE_Class where ClassID=" & ClassID)
            If tClass.BOF And tClass.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>找不到指定的栏目！</li>"
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
                        ErrMsg = ErrMsg & "<li>指定的栏目不允许添加" & ChannelShortName & "</li>"
                    End If
                    If tClass("ClassType") = 2 Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>不能指定为外部栏目</li>"
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
                        ErrMsg = ErrMsg & "<li>对不起，你没有栏目的相应操作权限！</li>"
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
'函数名：CheckPurview_Other
'作  用：其他权限数组检测
'参  数：AllPurviews ---- 要比较数组
'        strPurview ---- 比较字符
'返回值：True  ---- 存在
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
'函数名：CheckPurview_Class
'作  用：栏目权限数组检测
'参  数：str1 ---- 要比较数组1
'        str2 ---- 要比较数组2
'返回值：True  ---- 存在
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
'函数名：CheckPurview_Channel
'作  用：频道允许用户组检测
'参  数：ChannelPurview ---- 频道的权限　0为开放频道（任何人可以浏览），1为认证频道（游客不能浏览）
'        ChannelArrGroupID ---- 允许访问的用户组
'        GroupID ---- 用户所属的用户组
'返回值：True  ---- 有权限访问
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
'过程名：WriteEntry
'作  用：将信息写入日志
'参  数：LogType：1--重要操作  2--系统操作  3--频道操作  4--登录失败  5--管理错误
'        UserName    ------ 操作人
'        LogContent  ------ 操作信息
'
'说  明：LogType及ChannelID的取值
'==================================================
Public Sub WriteEntry(LogType, UserName, LogContent)
    Dim sqlLog, rsLog
    sqlLog = "select top 1 * from PE_Log"
    Set rsLog = Server.CreateObject("Adodb.RecordSet")
    rsLog.Open sqlLog, Conn, 1, 3
    rsLog.addnew
    rsLog("LogType") = LogType
    rsLog("ChannelID") = 0      '保留字段
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
'函数名：GetScriptName
'作  用：获取被访问文件目录
'返回值：文件目录
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
'函数名：GetPostString
'作  用：判断对方访问是 Form 还是 QueryString
'返回值：对方访问的数据
'**************************************************
Function GetPostString()
    Dim PostString, PostItem
    PostString = ""
    If Request.Form <> "" Then
        PostString = PostString & "Request.Form："
        For Each PostItem In Request.Form
            PostString = PostString & PostItem & "=" & Request.Form(PostItem) & "&"
        Next
        If Right(PostString, 1) = "&" Then PostString = Left(PostString, Len(PostString) - 1)
    End If
    If Request.QueryString <> "" Then
        If PostString <> "" Then PostString = PostString & vbCrLf
        PostString = PostString & "Request.QueryString："
        For Each PostItem In Request.QueryString
            PostString = PostString & PostItem & "=" & Request.QueryString(PostItem) & "&"
        Next
        If Right(PostString, 1) = "&" Then PostString = Left(PostString, Len(PostString) - 1)
    End If
    GetPostString = PostString
End Function

%>
