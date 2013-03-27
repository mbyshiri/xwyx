<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_Content.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<!--#include file="../Include/PowerEasy.SendMail.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 3   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

Dim HtmlDir, PurviewChecked, AddType
Dim ManageType, Status, MyStatus, arrStatus
Dim ClassID, SpecialID, OnTop, IsElite, IsHot, Created

Dim tClass, ClassName, RootID, ParentID, Depth, ParentPath, Child, arrChildID, ParentDir, ClassDir, ClassPurview

Dim PhotoID
Dim VoteID

If ChannelID = 0 Then
    Response.Write "频道参数丢失！"
    FoundErr = True
    Response.End
End If
If ModuleType <> 3 Then
    FoundErr = True
    Response.Write "<li>指定的频道ID不对！</li>"
    Response.End
End If
ModuleName = "Photo"
SheetName = "PE_Photo"

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
arrStatus = Array("待审核", "一审通过", "二审通过", "终审通过")

HtmlDir = InstallDir & ChannelDir

ManageType = Trim(Request("ManageType"))
Status = Trim(Request("Status"))
Created = Trim(Request("Created"))
OnTop = Trim(Request("OnTop"))
IsElite = Trim(Request("IsElite"))
IsHot = Trim(Request("IsHot"))
ClassID = PE_CLng(Trim(Request("ClassID")))
SpecialID = PE_CLng(Trim(Request("SpecialID")))
PhotoID = Trim(Request("PhotoID"))
AddType = Trim(Request("AddType"))

If Action = "" Then
    Action = "Manage"
End If
If Status = "" Then
    Status = 9
Else
    Status = PE_CLng(Status) '图片状态   9－－所有图片，-1－－草稿，0－－待审核，1－－已审核
End If
If IsValidID(PhotoID) = False Then
    PhotoID = ""
End If
If AddType = "" Then
    AddType = 1
Else
    AddType = PE_CLng(AddType)
End If

FileName = "Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=" & Action & "&ManageType=" & ManageType
strFileName = FileName & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&Field=" & strField & "&keyword=" & Keyword

If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))
    
    
Response.Write "<html><head><title>" & ChannelShortName & "管理</title>" & vbCrLf
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
strTitle = ChannelName & "管理----"
Select Case Action
Case "Add"
    strTitle = strTitle & "添加" & ChannelShortName
Case "Modify"
    strTitle = strTitle & "修改" & ChannelShortName
Case "Check"
    strTitle = strTitle & "审核" & ChannelShortName
Case "SaveAdd", "SaveModify"
    strTitle = strTitle & "保存" & ChannelShortName
Case "Move"
    strTitle = strTitle & "移动" & ChannelShortName
Case "Preview", "Show"
    strTitle = strTitle & "预览" & ChannelShortName
Case "Batch", "DoBatch"
    strTitle = strTitle & "批量修改" & ChannelShortName & "属性"
Case "MoveToClass"
    strTitle = strTitle & "批量移动" & ChannelShortName
Case "AddToSpecial"
    strTitle = strTitle & "添加" & ChannelShortName & "到专题"
Case "MoveToSpecial"
    strTitle = strTitle & "移动" & ChannelShortName & "到专题"
Case "Manage"
    Select Case ManageType
    Case "Check"
        strTitle = strTitle & ChannelShortName & "审核"
    Case "HTML"
        strTitle = strTitle & ChannelShortName & "生成"
    Case "Recyclebin"
        strTitle = strTitle & ChannelShortName & "回收站管理"
    Case "Special"
        strTitle = strTitle & "专题" & ChannelShortName & "管理"
    Case Else
        strTitle = strTitle & ChannelShortName & "管理首页"
    End Select
End Select
Call ShowPageTitle(strTitle, 10131)
Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30' ><strong>管理导航：</strong></td><td colspan='5'>"
Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Status=9'>" & ChannelShortName & "管理首页</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Add&AddType=1&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>添加" & ChannelShortName & "</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Add&AddType=3&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>添加" & ChannelShortName & "（批量模式）</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&ManageType=Check&Status=0'>审核" & ChannelShortName & "</a>"
If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&ManageType=Special'>专题" & ChannelShortName & "管理</a>"
End If
If AdminPurview = 1 Or AdminPurview_Channel = 1 Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&ManageType=Recyclebin&Status=9'>" & ChannelShortName & "回收站管理</a>"
End If
If UseCreateHTML > 0 And ObjInstalled_FSO = True Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&ManageType=HTML&Status=1'><b>生成HTML管理</b></a>"
End If
Response.Write "</td></tr>" & vbCrLf

If Action = "Manage" Then
    Response.Write "<form name='form3' method='Post' action='" & strFileName & "'><tr class='tdbg'>"
    Response.Write "  <td width='70' height='30' ><strong>" & ChannelShortName & "选项：</strong></td><td>"
    If ManageType = "HTML" Then
        Response.Write "<input name='Created' type='radio' onclick='submit();' " & RadioValue(Created, "") & ">所有" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='Created' type='radio' onclick='submit();' " & RadioValue(Created, "False") & ">未生成的" & ChannelShortName & "&nbsp;&nbsp;&nbsp"
        Response.Write "<input name='Created' type='radio' onclick='submit();' " & RadioValue(Created, "True") & ">已生成的" & ChannelShortName
    Else
        Response.Write "<input name='Status' type='radio' onclick='submit();' " & RadioValue(Status, 9) & ">所有" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='Status' type='radio' onclick='submit();' " & RadioValue(Status, -1) & ">草稿&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='Status' type='radio' onclick='submit();' " & RadioValue(Status, 0) & ">待审核&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='Status' type='radio' onclick='submit();' " & RadioValue(Status, 1) & ">已审核&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='Status' type='radio' onclick='submit();' " & RadioValue(Status, -2) & ">退稿"
        Response.Write "</td><td>"
        Response.Write "<input name='OnTop' type='checkbox' onclick='submit()' " & RadioValue(OnTop, "True") & "> 固顶" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='IsElite' type='checkbox' onclick='submit()' " & RadioValue(IsElite, "True") & "> 推荐" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "<input name='IsHot' type='checkbox' onclick='submit()' " & RadioValue(IsHot, "True") & "> 热门" & ChannelShortName
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
    Call SavePhoto
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
Case "Manage"
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub main()
    If ManageType = "HTML" And UseCreateHTML = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>本频道设置了不生成HTML，所以不用进行生成管理！</li>"
        Exit Sub
    End If
    Dim rsPhotoList, sql, Querysql
    PurviewChecked = False
    If ClassID = 0 Then
        If strField = "" And AdminPurview = 2 And AdminPurview_Channel = 3 And ManageType <> "My" Then
            If ManageType = "Check" Then
                If arrClass_Check = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>对不起，您没有在此频道审核" & ChannelShortName & "的权限！</li>"
                    Exit Sub
                End If
                Set tClass = Conn.Execute("select top 1 ClassID,ClassName,RootID,ParentID,Depth,ParentPath,Child,arrChildID,ClassPurview,ParentDir,ClassDir from PE_Class where ChannelID=" & ChannelID & " and ClassID In (" & DelRightComma(arrClass_Check) & ")")
            Else
                If arrClass_Manage = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>对不起，您没有在此频道管理" & ChannelShortName & "的权限！</li>"
                    Exit Sub
                End If
                Set tClass = Conn.Execute("select top 1 ClassID,ClassName,RootID,ParentID,Depth,ParentPath,Child,arrChildID,ClassPurview,ParentDir,ClassDir from PE_Class where ChannelID=" & ChannelID & " and ClassID In (" & DelRightComma(arrClass_Manage) & ")")
            End If
            If tClass.BOF And tClass.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>对不起，您没有在此频道的管理权限！</li>"
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
            ErrMsg = ErrMsg & "<li>找不到指定的栏目</li>"
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
        Call ShowContentManagePath(ChannelShortName & "审核")
    Case "HTML"
        Call ShowContentManagePath(ChannelShortName & "生成")
    Case "Recyclebin"
        Call ShowContentManagePath(ChannelShortName & "回收站管理")
    Case "Special"
        Call ShowContentManagePath("专题" & ChannelShortName & "管理")
    Case Else
        Call ShowContentManagePath(ChannelShortName & "管理")
    End Select

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='2' class='border'>"
    Response.Write "<form name='myform' method='Post' action='Admin_Photo.asp' onsubmit='return ConfirmDel();'><tr>"

    If ManageType = "Special" Then
        sql = "select top " & MaxPerPage & " I.InfoID,I.SpecialID,P.PhotoID,SP.SpecialName,P.PhotoName,P.Keyword,P.Author,P.UpdateTime,P.Inputer,"
        sql = sql & "P.PhotoThumb,P.Hits,P.OnTop,P.Elite,P.Status,P.Stars,P.InfoPoint,P.VoteID "
        sql = sql & " from PE_Photo P right join (PE_InfoS I left join PE_Special SP on I.SpecialID=SP.SpecialID) on P.PhotoID=I.ItemID "
    Else
        If ClassID = -1 Or (ClassID > 0 And Child = 0) Then
            sql = "select top " & MaxPerPage & " P.ClassID,P.PhotoID,P.PhotoName,P.Keyword,P.Author,P.UpdateTime,P.Inputer,"
            sql = sql & "P.PhotoThumb,P.Hits,P.OnTop,P.Elite,P.Status,P.Stars,P.InfoPoint,P.VoteID "
            sql = sql & " from PE_Photo P "
        Else
            sql = "select top " & MaxPerPage & " P.ClassID,P.PhotoID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview,P.PhotoName,P.Keyword,P.Author,P.UpdateTime,P.Inputer,"
            sql = sql & "P.PhotoThumb,P.Hits,P.OnTop,P.Elite,P.Status,P.Stars,P.InfoPoint,P.VoteID "
            sql = sql & " from PE_Photo P left join PE_Class C on P.ClassID=C.ClassID "
        End If
    End If
    
    Querysql = " where P.ChannelID=" & ChannelID
    If ManageType = "Special" Then
        Querysql = Querysql & " and I.ModuleType=" & ModuleType
    End If
    If ManageType = "Recyclebin" Then
        Querysql = Querysql & " and P.Deleted=" & PE_True & ""
    Else
        Querysql = Querysql & " and P.Deleted=" & PE_False & ""
    End If
    If ManageType = "HTML" Then
        If Created = "False" Then
            Querysql = Querysql & " and (P.CreateTime<=P.UpdateTime or P.CreateTime is Null)"
        ElseIf Created = "True" Then
            Querysql = Querysql & " and P.CreateTime>P.UpdateTime"
        End If
        Querysql = Querysql & " and P.Status=3" '当图片为已审核时，才在生成管理中出现
    Else
        Select Case Status
        Case -2 '退稿
            Querysql = Querysql & " and P.Status=-2"
        Case -1 '草稿
            Querysql = Querysql & " and P.Status=-1"
        Case 0  '待审核
            Querysql = Querysql & " and P.Status>=0 and P.Status<" & MyStatus
        Case 1  '已审核
            Querysql = Querysql & " and P.Status>=" & MyStatus
        Case Else
            Querysql = Querysql & " and P.Status>-1"
        End Select
        If OnTop = "True" Then
            Querysql = Querysql & " and P.OnTop=" & PE_True & ""
        End If
        If IsElite = "True" Then
            Querysql = Querysql & " and P.Elite=" & PE_True & ""
        End If
        If IsHot = "True" Then
            Querysql = Querysql & " and P.Hits>=" & HitsOfHot & ""
        End If
    End If

    If ClassID <> 0 Then
        If Child > 0 Then
            Querysql = Querysql & " and P.ClassID in (" & arrChildID & ")"
        Else
            Querysql = Querysql & " and P.ClassID=" & ClassID
        End If
    End If
    If SpecialID > 0 Then
        Querysql = Querysql & " and I.SpecialID=" & SpecialID
    End If
    If ManageType = "My" Then
        Querysql = Querysql & " and P.Inputer='" & UserName & "' "
    End If
    If Keyword <> "" Then
        Select Case strField
        Case "PhotoName"
            Querysql = Querysql & " and P.PhotoName like '%" & Keyword & "%' "
        Case "PhotoIntro"
            Querysql = Querysql & " and P.PhotoIntro like '%" & Keyword & "%' "
        Case "Author"
            Querysql = Querysql & " and P.Author like '%" & Keyword & "%' "
        Case "Inputer"
            Querysql = Querysql & " and P.Inputer='" & Keyword & "' "
        Case Else
            Querysql = Querysql & " and P.PhotoName like '%" & Keyword & "%' "
        End Select
    End If
    If ManageType = "Special" Then
        totalPut = PE_CLng(Conn.Execute("select Count(*) from PE_InfoS I inner join PE_Photo P on I.ItemID=P.PhotoID " & Querysql)(0))
    Else
        totalPut = PE_CLng(Conn.Execute("select Count(*) from PE_Photo P " & Querysql)(0))
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
            Querysql = Querysql & " and I.InfoID < (select min(InfoID) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " I.InfoID from PE_InfoS I inner join PE_Photo P on I.ItemID=P.PhotoID " & Querysql & " order by I.InfoID desc) as QueryPhoto)"
        Else
            Querysql = Querysql & " and P.PhotoID < (select min(PhotoID) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " P.PhotoID from PE_Photo P " & Querysql & " order by P.PhotoID desc) as QueryPhoto)"
        End If
    End If
    If ManageType = "Special" Then
        sql = sql & Querysql & " order by I.InfoID desc"
    Else
        sql = sql & Querysql & " order by P.PhotoID desc"
    End If

    Set rsPhotoList = Server.CreateObject("ADODB.Recordset")
    rsPhotoList.Open sql, Conn, 1, 1
    If rsPhotoList.BOF And rsPhotoList.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>"

        If ClassID > 0 Then
            Response.Write "此栏目及其子栏目中没有任何"
        Else
            Response.Write "没有任何"
        End If

        Select Case Status
        Case -2
            Response.Write "退稿"
        Case -1
            Response.Write "草稿"
        Case 0
            Response.Write "<font color=blue>待审核</font>的" & ChannelShortName & "！"
        Case 1
            Response.Write "<font color=green>已审核</font>的" & ChannelShortName & "！"
        Case Else
            Response.Write ChannelShortName & "！"
        End Select
        Response.Write "<br><br></td></tr>"
    Else
        Dim PhotoNum, PhotoPath
        PhotoNum = 0
        Do While Not rsPhotoList.EOF
            Response.Write "<td class='tdbg'><table width='100%'  cellpadding='0' cellspacing='0' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "<tr><td colspan='2' align='center'><a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Show&PhotoID=" & rsPhotoList("PhotoID") & "'><img src='" & GetPhotoThumb(rsPhotoList("PhotoThumb")) & "' width='130' height='90' border='0'></a></td></tr>"
            If ManageType = "Special" Then
                Response.Write "<tr><td align='right'>专题名称：</td><td>"
                If rsPhotoList("SpecialID") > 0 Then
                    Response.Write "<a href='" & FileName & "&SpecialID=" & rsPhotoList("SpecialID") & "'>" & rsPhotoList("SpecialName") & "</a>"
                Else
                    Response.Write "&nbsp;"
                End If
                Response.Write "</td></tr>"
            Else
                If rsPhotoList("ClassID") <> ClassID And ClassID <> -1 Then
                    Response.Write "<tr><td align='right'>栏目名称：</td><td><a href='" & FileName & "&ClassID=" & rsPhotoList("ClassID") & "'>["
                    If rsPhotoList("ClassName") <> "" Then
                        Response.Write rsPhotoList("ClassName")
                    Else
                        Response.Write "<font color='gray'>不属于任何栏目</font>"
                    End If
                    Response.Write "]</a></td></tr>"
                End If
            End If
            Response.Write "<tr><td align='right'>" & ChannelShortName & "名称：</td><td>"
            Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Show&PhotoID=" & rsPhotoList("PhotoID") & "'"
            Response.Write " title='名&nbsp;&nbsp;&nbsp;&nbsp;称：" & rsPhotoList("PhotoName") & vbCrLf & "作&nbsp;&nbsp;&nbsp;&nbsp;者：" & rsPhotoList("Author") & vbCrLf & "更新时间：" & rsPhotoList("UpdateTime") & vbCrLf
            Response.Write "查看次数：" & rsPhotoList("Hits") & vbCrLf & "关 键 字：" & Mid(rsPhotoList("Keyword"), 2, Len(rsPhotoList("Keyword")) - 2) & vbCrLf & "推荐等级："
            If rsPhotoList("Stars") = 0 Then
                Response.Write "无"
            Else
                Response.Write String(rsPhotoList("Stars"), "★")
            End If
            Response.Write vbCrLf & "查看" & PointName & "数：" & rsPhotoList("InfoPoint")
            Response.Write "'>" & rsPhotoList("PhotoName") & "</a></td></tr>"
            Response.Write "<tr><td align='right'>添 加 者：</td><td><a href='" & FileName & "&field=Inputer&keyword=" & rsPhotoList("Inputer") & "' title='点击将查看此用户录入的所有" & ChannelShortName & "'>" & rsPhotoList("Inputer") & "</a></td></tr>"
            Response.Write "<tr><td align='right'>点 击 数：</td><td>" & rsPhotoList("Hits") & "</td></tr>"
            Response.Write "<tr><td align='right'>" & ChannelShortName & "属性：</td><td>"
            If rsPhotoList("OnTop") = True Then
                Response.Write "<font color=blue>顶</font>&nbsp;"
            Else
                Response.Write "&nbsp;&nbsp;&nbsp;"
            End If
            If rsPhotoList("Hits") >= HitsOfHot Then
                Response.Write "<font color=red>热</a>&nbsp;"
            Else
                Response.Write "&nbsp;&nbsp;&nbsp;"
            End If
            If rsPhotoList("Elite") = True Then
                Response.Write "<font color=green>荐</a>"
            Else
                Response.Write "&nbsp;&nbsp;"
            End If
            If rsPhotoList("VoteID") > 0 Then
                Response.Write "<a href='" & InstallDir & "Vote.asp?ID=" & rsPhotoList("VoteID") & "&Action=Show' target='_blank'>调</a>"
            Else
                Response.Write "&nbsp;&nbsp;"
            End If
            Response.Write "</td></tr>"
            Response.Write "<tr><td align='right'>" & ChannelShortName & "状态：</td><td>"
            Select Case rsPhotoList("Status")
            Case -2
                Response.Write "<font color=gray>退稿</font>"
            Case -1
                Response.Write "<font color=gray>草稿</font>"
            Case 0
                Response.Write "<font color=red>待审核</font>"
            Case 1
                Response.Write "<font color=blue>一审通过</font>"
            Case 2
                Response.Write "<font color=green>二审通过</font>"
            Case 3
                Response.Write "<font color=black>终审通过</font>"
            End Select
            Response.Write "</td></tr>"

            Dim iClassPurview
            If UseCreateHTML > 0 And ObjInstalled_FSO = True And ManageType <> "Special" Then
                Response.Write "<tr><td align='right'>已 生 成：</td><td>"
                If ClassID = -1 Or (ClassID > 0 And Child = 0) Then
                    iClassPurview = ClassPurview
                    PhotoPath = HtmlDir & GetItemPath(StructureType, ParentDir, ClassDir, rsPhotoList("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsPhotoList("UpdateTime"), rsPhotoList("PhotoID")) & FileExt_Item
                Else
                    iClassPurview = PE_CLng(rsPhotoList("ClassPurview"))
                    PhotoPath = HtmlDir & GetItemPath(StructureType, rsPhotoList("ParentDir"), rsPhotoList("ClassDir"), rsPhotoList("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsPhotoList("UpdateTime"), rsPhotoList("PhotoID")) & FileExt_Item
                End If
                If PE_CLng(iClassPurview) > 0 Or rsPhotoList("InfoPoint") > 0 Then
                    Response.Write "<a href='#' title='因为设置了查看权限，所以不用生成HTML'><font color=green><b>¤</b></font></a>"
                Else
                    If fso.FileExists(Server.MapPath(PhotoPath)) Then
                        Response.Write "<a href='#' title='文件位置：" & PhotoPath & "'><b>√</b></a>"
                    Else
                        Response.Write "<font color=red><b>×</b></font>"
                    End If
                End If
                Response.Write "</td></tr>"
            End If
            Response.Write "<tr><td align='right'>操作选项：</td><td>"
            If ManageType = "Special" Then
                Response.Write "<input name='InfoID' type='checkbox' onclick='CheckItem(this,""TABLE"")' id='InfoID' value='" & rsPhotoList("InfoID") & "'>"
            Else
                Response.Write "<input name='PhotoID' type='checkbox' onclick='CheckItem(this,""TABLE"")' id='PhotoID' value='" & rsPhotoList("PhotoID") & "'>"
            End If
            Response.Write "</td></tr>"
            Response.Write "<tr><td colspan='2' align='center'>"
            If ManageType = "Check" Then
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
                    If rsPhotoList("Status") <= MyStatus Then
                        If rsPhotoList("Status") > -1 Then
                            Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Reject&PhotoID=" & rsPhotoList("PhotoID") & "'>直接退稿</a>&nbsp;&nbsp;"
                        End If
                        If rsPhotoList("Status") < MyStatus Then
                            Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Check&PhotoID=" & rsPhotoList("PhotoID") & "'>审核</a>&nbsp;&nbsp;"
                            Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=SetPassed&PhotoID=" & rsPhotoList("PhotoID") & "'>通过</a>&nbsp;&nbsp;"
                        Else
                            Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=CancelPassed&PhotoID=" & rsPhotoList("PhotoID") & "'>取消审核</a>"
                        End If
                    End If
                End If
            ElseIf ManageType = "HTML" Then
                If iClassPurview = 0 And rsPhotoList("InfoPoint") = 0 And rsPhotoList("Status") = 3 And (AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True) Then
                    Response.Write "<a href='Admin_CreatePhoto.asp?ChannelID=" & ChannelID & "&Action=CreatePhoto&PhotoID=" & rsPhotoList("PhotoID") & "' title='生成本" & ChannelShortName & "的HTML页面'>生成文件</a>&nbsp;"
                    If fso.FileExists(Server.MapPath(PhotoPath)) Then
                        Response.Write "<a href='" & PhotoPath & "' target='_blank' title='查看本" & ChannelShortName & "的HTML页面'>查看文件</a>&nbsp;"
                        Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=DelFile&PhotoID=" & rsPhotoList("PhotoID") & "' title='删除本" & ChannelShortName & "的HTML页面' onclick=""return confirm('确定要删除此" & ChannelShortName & "的HTML页面吗？');"">删除文件</a>&nbsp;"
                    End If
                End If
            ElseIf ManageType = "Recyclebin" Then
                Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=ConfirmDel&PhotoID=" & rsPhotoList("PhotoID") & "' onclick=""return confirm('确定要彻底删除此" & ChannelShortName & "吗？彻底删除后将无法还原！');"">彻底删除</a> "
                Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Restore&PhotoID=" & rsPhotoList("PhotoID") & "'>还原</a>"
            ElseIf ManageType = "Special" Then
                If rsPhotoList("SpecialID") > 0 Then
                    Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=DelFromSpecial&InfoID=" & rsPhotoList("InfoID") & "' onclick=""return confirm('确定要将此" & ChannelShortName & "从其所属专题中删除吗？操作成功后此" & ChannelShortName & "将不属于任何专题。');"">从所属专题中删除</a> "
                End If
            Else
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or CheckPurview_Class(arrClass_Input, ParentPath & "," & ClassID) Or UserName = rsPhotoList("Inputer") Then
                    Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Modify&PhotoID=" & rsPhotoList("PhotoID") & "'>修改</a>&nbsp;"
                Else
                    Response.Write "<font color='gray'>修改&nbsp;</font>"
                End If
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Or UserName = rsPhotoList("Inputer") Then
                    Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Del&PhotoID=" & rsPhotoList("PhotoID") & "' onclick=""return confirm('确定要删除此" & ChannelShortName & "吗？删除后你还可以从回收站中还原。');"">删除</a>&nbsp;"
                Else
                    Response.Write "<font color='gray'>删除&nbsp;</font>"
                End If
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
                    If rsPhotoList("OnTop") = False Then
                        Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=SetOnTop&PhotoID=" & rsPhotoList("PhotoID") & "'>固顶</a>&nbsp;"
                    Else
                        Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=CancelOnTop&PhotoID=" & rsPhotoList("PhotoID") & "'>解固</a>&nbsp;"
                    End If
                    If rsPhotoList("Elite") = False Then
                        Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=SetElite&PhotoID=" & rsPhotoList("PhotoID") & "'>设为推荐</a>"
                    Else
                        Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=CancelElite&PhotoID=" & rsPhotoList("PhotoID") & "'>取消推荐</a>"
                    End If
                End If
            End If
            Response.Write "</td></tr>"
            Response.Write "</table></td>"

            PhotoNum = PhotoNum + 1
            If PhotoNum Mod 4 = 0 Then
                Response.Write "</tr><tr>"
            End If
            If PhotoNum >= MaxPerPage Then Exit Do
            rsPhotoList.MoveNext
        Loop
    End If
    rsPhotoList.Close
    Set rsPhotoList = Nothing
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form,""TABLE"")' value='checkbox'>选中本页显示的所有" & ChannelShortName & "</td><td>"
    Select Case ManageType
    Case "Check"
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
            Response.Write "<input name='submit1' type='submit' value=' 审核通过 ' onClick=""document.myform.Action.value='SetPassed'"">&nbsp;&nbsp;"
            Response.Write "<input name='submit2' type='submit' value=' 取消审核 ' onClick=""document.myform.Action.value='CancelPassed'"">&nbsp;&nbsp;"
            If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
                Response.Write "<input name='submit3' type='submit' value=' 批量删除 ' onClick=""document.myform.Action.value='Del'"">"
            End If
        End If
    Case "HTML"
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
            Response.Write "<input name='CreateType' type='hidden' id='CreateType' value='1'>"
            Response.Write "<input name='ClassID' type='hidden' id='ClassID' value='" & ClassID & "'>"
            If ClassID > 0 Then
                If UseCreateHTML = 1 Or UseCreateHTML = 3 And ClassPurview < 2 Then
                    Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateClass';document.myform.CreateType.value='1';document.myform.action='Admin_CreatePhoto.asp';"" value='生成当前栏目列表页'>&nbsp;&nbsp;"
                End If
                If ClassPurview = 0 Then
                    Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreatePhoto';document.myform.CreateType.value='2';document.myform.action='Admin_CreatePhoto.asp';"" value='生成当前栏目的" & ChannelShortName & "'>&nbsp;&nbsp;"
                End If
            Else
                Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateIndex';document.myform.CreateType.value='1';document.myform.action='Admin_CreatePhoto.asp';"" value='生成首页'>&nbsp;&nbsp;"
                If UseCreateHTML = 1 Or UseCreateHTML = 3 Then
                    Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateClass';document.myform.CreateType.value='2';document.myform.action='Admin_CreatePhoto.asp';"" value='生成所有栏目列表页'>&nbsp;&nbsp;"
                End If
                Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreatePhoto';document.myform.CreateType.value='3';document.myform.action='Admin_CreatePhoto.asp';"" value='生成所有" & ChannelShortName & "'>&nbsp;&nbsp;"
            End If
            Response.Write "<input name='submit2' type='submit' id='submit2' onClick=""document.myform.Action.value='CreatePhoto';document.myform.action='Admin_CreatePhoto.asp';"" value='生成选定的" & ChannelShortName & "'>&nbsp;&nbsp;&nbsp;&nbsp;"
            Response.Write "<input name='Submit3' type='submit' id='submit3' onClick=""document.myform.Action.value='DelFile';document.myform.action='Admin_Photo.asp'"" value='删除选定" & ChannelShortName & "的HTML文件'>"
        End If
    Case "Recyclebin"
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
            Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='ConfirmDel'"" value=' 彻底删除 '>&nbsp;"
            Response.Write "<input name='Submit2' type='submit' id='Submit2' onClick=""document.myform.Action.value='ClearRecyclebin'"" value='清空回收站'>&nbsp;&nbsp;&nbsp;&nbsp;"
            Response.Write "<input name='Submit3' type='submit' id='Submit3' onClick=""document.myform.Action.value='Restore'"" value='还原选定的" & ChannelShortName & "'>&nbsp;"
            Response.Write "<input name='Submit4' type='submit' id='Submit4' onClick=""document.myform.Action.value='RestoreAll'"" value='还原所有" & ChannelShortName & "'>"
        End If
    Case "Special"
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
            Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='DelFromSpecial'"" value='从所属专题中移除'> "
            Response.Write "<input name='Submit2' type='submit' id='Submit2' onClick=""document.myform.Action.value='AddToSpecial'"" value='添加到其他专题中'> "
            Response.Write "<input name='Submit2' type='submit' id='Submit2' onClick=""document.myform.Action.value='MoveToSpecial'"" value='移动到另一专题中'>"
        End If
    Case Else
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
            Response.Write "<input name='submit1' type='submit' value=' 批量删除 ' onClick=""document.myform.Action.value='Del'""> "
            If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
                Response.Write "<input type='submit' name='Submit4' value=' 批量移动 ' onClick=""document.myform.Action.value='MoveToClass'""> "
                Response.Write "<input type='submit' name='Submit3' value=' 批量设置 ' onClick=""document.myform.Action.value='Batch'""> "
                Response.Write "<input name='submit1' type='submit' value=' 审核通过 ' onClick=""document.myform.Action.value='SetPassed'""> "
                Response.Write "<input name='submit2' type='submit' value=' 取消审核 ' onClick=""document.myform.Action.value='CancelPassed'""> "
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
    Response.Write "   <td width='80' align='right'><strong>" & ChannelShortName & "搜索：</strong></td>"
    Response.Write "   <td>"
    Response.Write "<select name='Field' size='1'>"
    Response.Write "<option value='PhotoName' selected>" & ChannelShortName & "名称</option>"
    Response.Write "<option value='PhotoIntro'>" & ChannelShortName & "简介</option>"
    Response.Write "<option value='Author'>" & ChannelShortName & "作者</option>"
    If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
        Response.Write "<option value='Inputer'>录入者</option>"
    End If
    Response.Write "<option value='ID'>" & ChannelShortName & "ID</option>"
    Response.Write "</select>"
    Response.Write "<select name='ClassID'><option value=''>所有栏目</option>" & GetClass_Option(1, 0) & "</select>"
    Response.Write "<input type='text' name='keyword'  size='20' value='关键字' maxlength='50' onFocus='this.select();'>"
    Response.Write "<input type='submit' name='Submit'  value='搜索'>"
    Response.Write "<input name='ManageType' type='hidden' id='ManageType' value='" & ManageType & "'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "</td></tr></table></form>"
    Response.Write "<br><b>说明：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "属性中的各项含义：<font color=blue>顶</font>----固顶" & ChannelShortName & "，<font color=red>热</font>----热门" & ChannelShortName & "，<font color=green>荐</font>----推荐" & ChannelShortName & "，<font color=black>调</font>----有调查<br><br>"
End Sub

Sub ShowJS_Photo()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function SelectPhoto(iType){" & vbCrLf
    Response.Write "  var arr=showModalDialog('Admin_SelectFile.asp?ChannelID=" & ChannelID & "&DialogType=photo', '', 'dialogWidth:820px; dialogHeight:600px; help: no; scroll: yes; status: no');" & vbCrLf
    Response.Write "  if(arr!=null){" & vbCrLf
    Response.Write "    var ss=arr.split('|');" & vbCrLf
    Response.Write "    var strPhotoUrl=ss[0];" & vbCrLf
    Response.Write "    if(iType==0){document.myform.PhotoThumb.value=ss[0];}" & vbCrLf
    Response.Write "    else{" & vbCrLf
    Response.Write "    var url='" & ChannelShortName & "地址'+(document.myform.PhotoUrl.length+1)+'|'+strPhotoUrl;" & vbCrLf
    Response.Write "    document.myform.PhotoUrl.options[document.myform.PhotoUrl.length]=new Option(url,url);}" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function AddUrl(){" & vbCrLf
    Response.Write "  var thisurl='" & ChannelShortName & "地址'+(document.myform.PhotoUrl.length+1)+'|http://'; " & vbCrLf
    Response.Write "  var url=prompt('请输入" & ChannelShortName & "地址名称和链接，中间用“|”隔开：',thisurl);" & vbCrLf
    Response.Write "  if(url!=null&&url!=''){document.myform.PhotoUrl.options[document.myform.PhotoUrl.length]=new Option(url,url);}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function ModifyUrl(){" & vbCrLf
    Response.Write "  if(document.myform.PhotoUrl.length==0) return false;" & vbCrLf
    Response.Write "  var thisurl=document.myform.PhotoUrl.value; " & vbCrLf
    Response.Write "  if (thisurl=='') {alert('请先选择一个" & ChannelShortName & "地址，再点修改按钮！');return false;}" & vbCrLf
    Response.Write "  var url=prompt('请输入" & ChannelShortName & "地址名称和链接，中间用“|”隔开：',thisurl);" & vbCrLf
    Response.Write "  if(url!=thisurl&&url!=null&&url!=''){document.myform.PhotoUrl.options[document.myform.PhotoUrl.selectedIndex]=new Option(url,url);}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function DelUrl(){" & vbCrLf
    Response.Write "  if(document.myform.PhotoUrl.length==0) return false;" & vbCrLf
    Response.Write "  var thisurl=document.myform.PhotoUrl.value; " & vbCrLf
    Response.Write "  if (thisurl=='') {alert('请先选择一个" & ChannelShortName & "地址，再点删除按钮！');return false;}" & vbCrLf
    Response.Write "  document.myform.PhotoUrl.options[document.myform.PhotoUrl.selectedIndex]=null;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  var CurrentMode=editor.CurrentMode;" & vbCrLf
    Response.Write "  if (CurrentMode==0){" & vbCrLf
    Response.Write "    document.myform.PhotoIntro.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else if(CurrentMode==1){" & vbCrLf
    Response.Write "   document.myform.PhotoIntro.value=editor.HtmlEdit.document.body.innerText;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    alert('预览状态不能保存！请先回到编辑状态后再保存');" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.PhotoName.value==''){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "名称不能为空！');" & vbCrLf
    Response.Write "    document.myform.PhotoName.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Keyword.value==''){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('关键字不能为空！');" & vbCrLf
    Response.Write "    document.myform.Keyword.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.PhotoThumb.value==''){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('缩略图地址不能为空！');" & vbCrLf
    Response.Write "    document.myform.PhotoThumb.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    If AddType < 3 Then
        Response.Write "  if(document.myform.PhotoUrl.length==0){" & vbCrLf
        Response.Write "    ShowTabs(0);" & vbCrLf
        Response.Write "    alert('" & ChannelShortName & "地址不能为空！');" & vbCrLf
        Response.Write "    document.myform.PhotoUrl.focus();" & vbCrLf
        Response.Write "    return false;" & vbCrLf
        Response.Write "  }" & vbCrLf
        Response.Write "  document.myform.PhotoUrls.value='';" & vbCrLf
        Response.Write "  if (document.myform.Action.value!='Preview'){" & vbCrLf
        Response.Write "    for(var i=0;i<document.myform.PhotoUrl.length;i++){" & vbCrLf
        Response.Write "      if (document.myform.PhotoUrls.value=='') document.myform.PhotoUrls.value=document.myform.PhotoUrl.options[i].value;" & vbCrLf
        Response.Write "      else document.myform.PhotoUrls.value+='$$$'+document.myform.PhotoUrl.options[i].value;" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "  }" & vbCrLf
    End If
    Response.Write "  var obj=document.myform.ClassID;" & vbCrLf
    Response.Write "  var iCount=0;" & vbCrLf
    Response.Write "  for(var i=0;i<obj.length;i++){" & vbCrLf
    Response.Write "    if(obj.options[i].selected==true){" & vbCrLf
    Response.Write "      iCount=iCount+1;" & vbCrLf
    Response.Write "      if(obj.options[i].value==''){" & vbCrLf
    Response.Write "        ShowTabs(0);" & vbCrLf
    Response.Write "        alert('" & ChannelShortName & "所属栏目不能指定为外部栏目！');" & vbCrLf
    Response.Write "        document.myform.ClassID.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "      }" & vbCrLf
    Response.Write "      if(obj.options[i].selected==true&&obj.options[i].value=='0'){" & vbCrLf
    Response.Write "        ShowTabs(0);" & vbCrLf
    Response.Write "        alert('指定的栏目不允许添加" & ChannelShortName & "！只允许在其子栏目中添加" & ChannelShortName & "。');" & vbCrLf
    Response.Write "        document.myform.ClassID.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "      }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (iCount==0){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('请选择所属栏目！');" & vbCrLf
    Response.Write "    document.myform.ClassID.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
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
    Response.Write "     document.myform.VoteTitle.value = document.myform.PhotoName.value;" & vbCrLf
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
    'Response.Write "}" & vbCrLf
    'Response.Write "document.onkeypress = getKey;" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub ShowTabs_Title()
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>基本信息</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>所属专题</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>属性设置</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'"
    If FoundInArr(arrEnabledTabs, "Charge", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">收费选项</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(4);CopyTitle()'"
    If FoundInArr(arrEnabledTabs, "Vote", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">调查设置</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(5)'>自定义选项</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub

Sub ShowTabs_Bottom()
    Response.Write "<table id='Tabs_Bottom' width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center' height='24'>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title4' onclick='ShowTabs(0)'>基本信息</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(1)'>所属专题</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(2)'>属性设置</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(3)'"
    If FoundInArr(arrEnabledTabs, "Charge", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">收费选项</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(4);CopyTitle()'"
    If FoundInArr(arrEnabledTabs, "Vote", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">调查设置</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(5)'>自定义选项</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub

Sub Add()
    Call ShowJS_Photo
    
    
    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "'>" & ChannelName & "管理</a>&nbsp;&gt;&gt;&nbsp;添加" & ChannelShortName & "</td></tr></table>"
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Photo.asp' target='_self'>"

    Call ShowTabs_Title

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf

    Call ShowTr_Class
    
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "名称：</td>"
    Response.Write "            <td><div style=""clear: both;""><input name='PhotoName' type='text' value='' autocomplete='off' size='50' maxlength='255' onPropertyChange=""moreitem('PhotoName',10," & ChannelID & ",'satitle2');"" onBlur=""setTimeout('Element.hide(satitle2)',500);""> <font color='#FF0000'>*</font><input type='button' name='checksame' value='检查是否存在相同的" & ChannelShortName & "名' onclick=""showModalDialog('Admin_CheckSameTitle.asp?ModuleType=" & ModuleType & "&Title='+document.myform.PhotoName.value,'checksame','dialogWidth:350px; dialogHeight:250px; help: no; scroll: no; status: no');"">"
    Response.Write "                </div><div id=""satitle2"" style='display:none'></div>"
    If AddType = 3 Then
        Response.Write " <font color='blue'>可以输入数字序号通配符#</font>"
    End If
    Response.Write "</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>关键字：</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='Keyword' type='text' id='Keyword' value='" & Trim(Session("Keyword")) & "' autocomplete='off' size='50' maxlength='255' onPropertyChange=""moreitem('Keyword',10," & ChannelID & ",'skey');"" onBlur=""setTimeout('Element.hide(skey)',500);""> <font color='#FF0000'>*</font> " & GetKeywordList("Admin", ChannelID)
    Response.Write "              </div><div id=""skey"" style='display:none'></div>"
    Response.Write "              <font color='#0000FF'>用来查找相关" & ChannelShortName & "，可输入多个关键字，中间用<font color='#FF0000'>“|”</font>隔开。不能出现&quot;'&?;:()等字符。</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "作者：</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Author' type='text' id='Author' value='" & Trim(Session("Author")) & "' autocomplete='off' size='50' maxlength='100' onPropertyChange=""moreitem('Author',10," & ChannelID & ",'sauthor1');"" onBlur=""setTimeout('Element.hide(sauthor1)',500);"">" & GetAuthorList("Admin", ChannelID, UserName)
    Response.Write "              <div id=""sauthor1"" style='display:none'></div>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "来源：</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='CopyFrom' type='text' id='CopyFrom' value='" & Trim(Session("CopyFrom")) & "' autocomplete='off' size='50' maxlength='100' onPropertyChange=""moreitem('CopyFrom',10," & ChannelID & ",'scopyfrom1');"" onBlur=""setTimeout('Element.hide(scopyfrom1)',500);"">" & GetCopyFromList("Admin", ChannelID)
    Response.Write "              <div id=""scopyfrom1"" style='display:none'></div>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "简介：</td>"
    Response.Write "            <td>"
    Response.Write "              <textarea name='PhotoIntro' cols='67' rows='5' id='PhotoIntro' style='display:none'></textarea>"
    Response.Write "               <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=3&tContentid=PhotoIntro' frameborder='1' scrolling='no' width='700' height='200' ></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    If AddType < 3 Then
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'>缩略图：</td>"
        Response.Write "            <td>"
        Response.Write "              <input name='PhotoThumb' type='text' id='PhotoThumb' size='60' maxlength='200'>"
        Response.Write "              <input type='button' name='Button2' value='从已上传缩略图中选择' onclick='SelectPhoto(0)'>"
        Response.Write "            </td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "地址：</td>"
        Response.Write "            <td>"
        Response.Write "              <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
        Response.Write "                <tr>"
        Response.Write "                  <td width='410'>"
        Response.Write "                    <input type='hidden' name='PhotoUrls' value=''>"
        Response.Write "                    <select name='PhotoUrl' style='width:400;height:100' size='2' ondblclick='return ModifyUrl();'></select>"
        Response.Write "                  </td>"
        Response.Write "                  <td>"
        Response.Write "                    <input type='button' name='photoselect' value='从已上传" & ChannelShortName & "中选择' onclick='SelectPhoto(1)'><br><br>"
        Response.Write "                    <input type='button' name='addurl' value='添加外部地址' onclick='AddUrl();'><br>"
        Response.Write "                    <input type='button' name='modifyurl' value='修改当前地址' onclick='return ModifyUrl();'><br>"
        Response.Write "                    <input type='button' name='delurl' value='删除当前地址' onclick='DelUrl();'>"
        Response.Write "                  </td>"
        Response.Write "                </tr>"
        Response.Write "              </table>"
        Response.Write "            </td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'>上传" & ChannelShortName & "：</td>"
        Response.Write "            <td><iframe style='top:2px' id='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=photos' frameborder=0 scrolling=no width='650' height='150'></iframe></td>"
        Response.Write "          </tr>"
    Else
        Dim yyyy, mm, DD, ymd
        yyyy = Year(Date)
        mm = Right("0" & Month(Date), 2)
        DD = Right("0" & Day(Date), 2)
        ymd = yyyy & mm & DD
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='100' align='right' class='tdbg5'>缩略图：</td>"
        Response.Write "            <td colspan='3'>文件名格式：<input name='PhotoThumb' type='text' id='PhotoThumb' size='40' maxlength='200' value='" & yyyy & mm & "/" & ymd & "#_S.jpg'> 数字序号通配符为#，注意通配符只用一个#即可<br>开始ID：<input type='text' name='BeginID' value='01' size='6' maxlength='6' style='text-align:center'> 结束ID：<input type='text' name='EndID' value='99' size='6' maxlength='6' style='text-align:center'></td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "地址：</td>"
        Response.Write "            <td colspan='3'>文件名格式：<input name='PhotoUrls' type='text' id='PhotoUrl' size='40' maxlength='200' value='图片地址|" & yyyy & mm & "/" & ymd & "#.jpg'></td>"
        Response.Write "          </tr>"
    End If

    Call ShowTabs_Status_Add

    Response.Write "        </tbody>" & vbCrLf

    Call ShowTabs_Special(SpecialID, "")

    Call ShowTabs_Property_Add
    
    Call ShowTabs_Purview_Add("查看")
    
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
    Response.Write "   <input name='add' type='submit'  id='Add' value=' 添 加 ' onClick=""document.myform.Action.value='SaveAdd';document.myform.target='_self';"" style='cursor:hand;'>&nbsp; "
    Response.Write "   <input name='Preview' type='submit'  id='Preview' value=' 预 览 ' onClick=""document.myform.Action.value='Preview';document.myform.target='_blank';"" style='cursor:hand;'>"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
End Sub

Sub Modify()
    Dim rsPhoto, sql, tmpAuthor, tmpCopyFrom
    
    If PhotoID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的" & ChannelShortName & "ID</li>"
        Exit Sub
    Else
        PhotoID = PE_CLng(PhotoID)
    End If
    sql = "select * from PE_Photo where PhotoID=" & PhotoID & ""
    Set rsPhoto = Conn.Execute(sql)
    If rsPhoto.BOF And rsPhoto.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到" & ChannelShortName & "</li>"
        rsPhoto.Close
        Set rsPhoto = Nothing
        Exit Sub
    End If

    ClassID = rsPhoto("ClassID")
    Dim arrSpecialID
    arrSpecialID = GetSpecialIDArr(ModuleType, PhotoID)
    
    If rsPhoto("Inputer") <> UserName Then
        Call CheckClassPurview(Action, ClassID)
    End If
    
    If FoundErr = True Then
        rsPhoto.Close
        Set rsPhoto = Nothing
        Exit Sub
    End If
    tmpAuthor = rsPhoto("Author")
    tmpCopyFrom = rsPhoto("CopyFrom")
    EmailOfReject = Replace(EmailOfReject, "{$Title}", rsPhoto("PhotoName"))
    EmailOfPassed = Replace(EmailOfPassed, "{$Title}", rsPhoto("PhotoName"))

    Call ShowJS_Photo
    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "'>" & ChannelName & "管理</a>&nbsp;&gt;&gt;&nbsp;修改" & ChannelShortName & "</td></tr></table>"
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Photo.asp' target='_self'>"

    Call ShowTabs_Title

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf

    Call ShowTr_Class

    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "名称：</td>"
    Response.Write "            <td><div style=""clear: both;""><input name='PhotoName' type='text' value='" & rsPhoto("PhotoName") & "' autocomplete='off' size='50' maxlength='255' onPropertyChange=""moreitem('PhotoName',10," & ChannelID & ",'satitle2');"" onBlur=""setTimeout('Element.hide(satitle2)',500);""><font color='#FF0000'>*</font>"
    Response.Write "                </div><div id=""satitle2"" style='display:none'></div></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>关键字：</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='Keyword' type='text' id='Keyword' value='" & Mid(rsPhoto("Keyword"), 2, Len(rsPhoto("Keyword")) - 2) & "' autocomplete='off' size='50' maxlength='255' onPropertyChange=""moreitem('Keyword',10," & ChannelID & ",'skey');"" onBlur=""setTimeout('Element.hide(skey)',500);""> <font color='#FF0000'>*</font> " & GetKeywordList("Admin", ChannelID)
    Response.Write "              </div><div id=""skey"" style='display:none'></div>"
    Response.Write "              <font color='#0000FF'>用来查找相关" & ChannelShortName & "，可输入多个关键字，中间用<font color='#FF0000'>“|”</font>隔开。不能出现&quot;'&?;:()等字符。</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "作者：</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Author' type='text' id='Author' value='" & tmpAuthor & "' autocomplete='off' size='50' maxlength='100' onPropertyChange=""moreitem('Author',10," & ChannelID & ",'sauthor1');"" onBlur=""setTimeout('Element.hide(sauthor1)',500);"">" & GetAuthorList("Admin", ChannelID, UserName)
    Response.Write "              <div id=""sauthor1"" style='display:none'></div>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "来源：</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='CopyFrom' type='text' id='CopyFrom' value='" & tmpCopyFrom & "' autocomplete='off' size='50' maxlength='100' onPropertyChange=""moreitem('CopyFrom',10," & ChannelID & ",'scopyfrom1');"" onBlur=""setTimeout('Element.hide(scopyfrom1)',500);"">" & GetCopyFromList("Admin", ChannelID)
    Response.Write "              <div id=""scopyfrom1"" style='display:none'></div>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "简介：</td>"
    Response.Write "            <td>"
    Response.Write "              <textarea name='PhotoIntro' cols='67' rows='5' id='PhotoIntro' style='display:none'>" & Server.HTMLEncode(FilterBadTag(rsPhoto("PhotoIntro"), rsPhoto("Inputer"))) & "</textarea>"
    Response.Write "              <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=3&tContentid=PhotoIntro' frameborder='1' scrolling='no' width='700' height='200' ></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>缩略图：</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='PhotoThumb' type='text' id='PhotoThumb' size='60' maxlength='200' value='" & rsPhoto("PhotoThumb") & "'>"
    Response.Write "              <input type='button' name='Button2' value='从已上传缩略图中选择' onclick='SelectPhoto(0)'>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "地址：</td>"
    Response.Write "            <td>"
    Response.Write "              <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='410'>"
    Response.Write "                    <input type='hidden' name='PhotoUrls' value=''>"
    Response.Write "                    <select name='PhotoUrl' style='width:400;height:100' size='2' ondblclick='return ModifyUrl();'>"
    Dim PhotoUrls, arrPhotoUrls, iTemp
    PhotoUrls = rsPhoto("PhotoUrl")
    If InStr(PhotoUrls, "$$$") > 1 Then
        arrPhotoUrls = Split(PhotoUrls, "$$$")
        For iTemp = 0 To UBound(arrPhotoUrls)
            Response.Write "<option value='" & arrPhotoUrls(iTemp) & "'>" & arrPhotoUrls(iTemp) & "</option>"
        Next
    Else
        Response.Write "<option value='" & PhotoUrls & "'>" & PhotoUrls & "</option>"
    End If
    Response.Write "                    </select>"
    Response.Write "                  </td>"
    Response.Write "                  <td>"
    Response.Write "                    <input type='button' name='photoselect' value='从已上传" & ChannelShortName & "中选择' onclick='SelectPhoto(1)'><br><br>"
    Response.Write "                    <input type='button' name='addurl' value='添加外部地址' onclick='AddUrl();'><br>"
    Response.Write "                    <input type='button' name='modifyurl' value='修改当前地址' onclick='ModifyUrl();'><br>"
    Response.Write "                    <input type='button' name='delurl' value='删除当前地址' onclick='DelUrl();'>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>上传" & ChannelShortName & "：</td>"
    Response.Write "            <td><iframe style='top:2px' id='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=photos' frameborder=0 scrolling=no width='650' height='150'></iframe></td>"
    Response.Write "          </tr>"
    Call ShowTabs_Status_Modify(rsPhoto)
    Response.Write "        </tbody>" & vbCrLf
    
    Call ShowTabs_Special(arrSpecialID, "")
    
    Call ShowTabs_Property_Modify(rsPhoto)
    
    Call ShowTabs_Purview_Modify("查看", rsPhoto, "")
    
    Call ShowTabs_Vote_Modify(rsPhoto)

    Call ShowTabs_MyField_Modify(rsPhoto)
    
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>" & vbCrLf
    Call ShowTabs_Bottom
    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='SaveModify'><input name='AddType' type='hidden' id='AddType' value='1'>"
    Response.Write "   <input name='PhotoID' type='hidden' id='PhotoID' value='" & rsPhoto("PhotoID") & "'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='Save' type='submit' value='保存修改结果' onClick=""document.myform.Action.value='SaveModify';document.myform.target='_self';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Save' type='submit' value='添加为新" & ChannelShortName & "' onClick=""document.myform.Action.value='SaveAdd';document.myform.target='_self';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Preview' type='submit'  id='Preview' value=' 预 览 ' onClick=""document.myform.Action.value='Preview';document.myform.target='_blank';"" style='cursor:hand;'>"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
    rsPhoto.Close
    Set rsPhoto = Nothing
End Sub


Sub SavePhoto()
    Dim rsPhoto, sql, trs, i
    Dim PhotoID, ClassID, SpecialID, PhotoName, Keyword, Author, tAuthor, CopyFrom
    Dim PhotoIntro, PhotoThumb, PhotoUrl, Inputer, Editor, UpdateTime
    Dim BeginID, EndID, TempID, strTempID, strEndID
    Dim AddType
    Dim arrSpecialID
    AddType = PE_CLng(Request.Form("AddType"))
    If AddType = 0 And Action <> "SaveModify" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>所需参数" & ChannelShortName & "添加模式为空</li>"
    End If
    
    PhotoID = Trim(Request.Form("PhotoID"))
    ClassID = Trim(Request.Form("ClassID"))
    SpecialID = Trim(Request.Form("SpecialID"))

    PhotoName = Trim(Request.Form("PhotoName"))
    Keyword = Trim(Request.Form("Keyword"))
    Author = Trim(Request.Form("Author"))
    CopyFrom = Trim(Request.Form("CopyFrom"))
    PhotoIntro = Trim(Request.Form("PhotoIntro"))
    PhotoThumb = Trim(Request.Form("PhotoThumb"))
    PhotoUrl = Trim(Request.Form("PhotoUrls"))
    BeginID = PE_CLng(Trim(Request("BeginID")))
    EndID = PE_CLng(Trim(Request("EndID")))

    UpdateTime = PE_CDate(Trim(Request.Form("UpdateTime")))
    Status = PE_CLng(Trim(Request.Form("Status")))

    Inputer = UserName
    Editor = AdminName

    Call CheckClassPurview(Action, ClassID)
    If FoundErr = True Then Exit Sub
    
    If PhotoName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "名称不能为空</li>"
    End If
    Keyword = ReplaceBadChar(Keyword)
    If Keyword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入" & ChannelShortName & "关键字</li>"
    Else
        Call SaveKeyword(Keyword)
    End If
    If PhotoThumb = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>缩略图地址不能为空</li>"
    End If
    If PhotoUrl = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "地址不能为空</li>"
    End If
    If AddType = 3 Then
        If BeginID <= 0 Or EndID <= 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定序号开始ID和结束ID！</li>"
        End If
        If BeginID >= EndID Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>序号开始ID不能大于或等于结束ID！</li>"
        End If
    End If
    
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-3 or ChannelID=" & ChannelID & "")
    Do While Not rsField.EOF
        If rsField("EnableNull") = False Then
            If Trim(Request(rsField("FieldName"))) = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>请输入" & rsField("Title") & "！</li>"
            End If
        End If
        rsField.MoveNext
    Loop
    
    If FoundErr = True Then
        Exit Sub
    End If

    PhotoName = PE_HTMLEncode(PhotoName)
    Keyword = "|" & Keyword & "|"
    
    If Author = "" Then Author = XmlText("BaseText", "DefAuthor", "佚名")
    If CopyFrom = "" Then CopyFrom = XmlText("BaseText", "DefCopyFrom", "本站原创")

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

    Set rsPhoto = Server.CreateObject("adodb.recordset")
    If Action = "SaveAdd" Then
        If Session("PhotoName") = PhotoName And DateDiff("S", Session("AddTime"), Now()) < 100 Then
            FoundErr = True
            ErrMsg = "<li>请不要重复添加同一" & ChannelItemUnit & ChannelShortName & "</li>"
            Exit Sub
        Else
            Session("PhotoName") = PhotoName
            Session("AddTime") = Now()
            Dim iNum, rsblog, blogid
                   
            PhotoID = GetNewID("PE_Photo", "PhotoID")

            If UserID <> "" And UserID > 0 Then
                Set rsblog = Conn.Execute("select top 1 ID from PE_Space where Type=1 and UserID=" & UserID)
                If rsblog.BOF And rsblog.EOF Then
                    blogid = 0
                Else
                    blogid = rsblog("ID")
                End If
                Set rsblog = Nothing
            Else
                blogid = 0
            End If

            sql = "select top 1 * from PE_Photo"
            rsPhoto.Open sql, Conn, 1, 3
            If AddType < 3 Then
                For i = 0 To UBound(arrSpecialID)
                    Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (" & ModuleType & "," & PhotoID & "," & PE_CLng(arrSpecialID(i)) & ")")
                Next
                iNum = 1
                rsPhoto.addnew
                rsPhoto("PhotoID") = PhotoID
                rsPhoto("ChannelID") = ChannelID
                rsPhoto("Inputer") = Inputer
                rsPhoto("BlogID") = blogid
            Else
                iNum = EndID - BeginID + 1
                For TempID = BeginID To EndID
                    If IsNumeric(Trim(Request("EndID"))) Then
                        strEndID = Trim(Request("EndID"))
                    Else
                        strEndID = EndID
                    End If
                    strTempID = Right("00000" & TempID, Len(Trim(strEndID)))

                    For i = 0 To UBound(arrSpecialID)
                        Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (" & ModuleType & "," & PhotoID & "," & PE_CLng(arrSpecialID(i)) & ")")
                    Next
                    
                    rsPhoto.addnew
                    rsPhoto("PhotoID") = PhotoID
                    rsPhoto("ChannelID") = ChannelID
                    rsPhoto("ClassID") = ClassID
                    'rsPhoto("SpecialID") = SpecialID
                    rsPhoto("PhotoName") = Replace(PhotoName, "#", strTempID)
                    rsPhoto("Keyword") = Keyword
                    rsPhoto("Author") = Author
                    rsPhoto("CopyFrom") = CopyFrom
                    rsPhoto("PhotoIntro") = PhotoIntro
                    rsPhoto("PhotoThumb") = Replace(PhotoThumb, "#", strTempID)
                    rsPhoto("PhotoUrl") = Replace(PhotoUrl, "#", strTempID)
                    rsPhoto("Hits") = PE_CLng(Trim(Request.Form("Hits")))
                    rsPhoto("DayHits") = PE_CLng(Trim(Request.Form("DayHits")))
                    rsPhoto("WeekHits") = PE_CLng(Trim(Request.Form("WeekHits")))
                    rsPhoto("MonthHits") = PE_CLng(Trim(Request.Form("MonthHits")))
                    rsPhoto("Stars") = PE_CLng(Trim(Request.Form("Stars")))
                    rsPhoto("UpdateTime") = UpdateTime
                    rsPhoto("CreateTime") = UpdateTime
                    rsPhoto("Status") = Status
                    rsPhoto("OnTop") = PE_CBool(Trim(Request.Form("OnTop")))
                    rsPhoto("Elite") = PE_CBool(Trim(Request.Form("Elite")))
                    rsPhoto("Inputer") = Inputer
                    rsPhoto("BlogID") = blogid
                    rsPhoto("Editor") = Editor
                    rsPhoto("SkinID") = PE_CLng(Trim(Request.Form("SkinID")))
                    rsPhoto("TemplateID") = PE_CLng(Trim(Request.Form("TemplateID")))
                    rsPhoto("Deleted") = False
                    rsPhoto("PresentExp") = PresentExp
                    rsPhoto("InfoPurview") = PE_CLng(Trim(Request.Form("InfoPurview")))
                    rsPhoto("arrGroupID") = ReplaceBadChar(Trim(Request.Form("GroupID")))
                    rsPhoto("InfoPoint") = PE_CLng(Trim(Request.Form("InfoPoint")))
                    rsPhoto("ChargeType") = PE_CLng(Trim(Request.Form("ChargeType")))
                    rsPhoto("PitchTime") = PE_CLng(Trim(Request.Form("PitchTime")))
                    rsPhoto("ReadTimes") = PE_CLng(Trim(Request.Form("ReadTimes")))
                    rsPhoto("DividePercent") = PE_CLng(Trim(Request.Form("DividePercent")))
                    rsPhoto("VoteID") = VoteID
                    If Not (rsField.BOF And rsField.EOF) Then
                        rsField.MoveFirst
                        Do While Not rsField.EOF
                            If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                                rsPhoto(Trim(rsField("FieldName"))) = Trim(Request(rsField("FieldName")))
                            End If
                            rsField.MoveNext
                        Loop
                    End If
                    
                    rsPhoto.Update
                    PhotoID = PhotoID + 1
                Next
                Set rsField = Nothing
                PhotoID = PhotoID - 1
                rsPhoto.Close
            End If

            If Status = 3 Then
                Conn.Execute ("update PE_User set PassedItems=PassedItems+" & iNum & ",UserExp=UserExp+" & (iNum * PresentExp) & " where UserName='" & Inputer & "'")
            End If
        End If
    ElseIf Action = "SaveModify" Then
        If PhotoID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>不能确定PhotoID的值</li>"
        Else
            PhotoID = PE_CLng(PhotoID)
            sql = "select * from PE_Photo where PhotoID=" & PhotoID
            rsPhoto.Open sql, Conn, 1, 3
            If rsPhoto.BOF And rsPhoto.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>找不到此" & ChannelShortName & "，可能已经被其他人删除。</li>"
            Else
            
                '删除生成的文件，因为生成的文件可能会随着更新时间，游览权限等发生变化而产生多余的文件
                If UseCreateHTML > 0 And ObjInstalled_FSO = True Then
                    Dim tClass, PhotoPath
                    Set tClass = Conn.Execute("select ParentDir,ClassDir from PE_Class where ClassID=" & rsPhoto("ClassID") & "")
                    If tClass.BOF And tClass.EOF Then
                        ParentDir = "/"
                        ClassDir = ""
                    Else
                        ParentDir = tClass("ParentDir")
                        ClassDir = tClass("ClassDir")
                    End If
                    PhotoPath = HtmlDir & GetItemPath(StructureType, ParentDir, ClassDir, rsPhoto("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsPhoto("UpdateTime"), rsPhoto("PhotoID"))
                    If fso.FileExists(Server.MapPath(PhotoPath & FileExt_Item)) Then
                        DelSerialFiles Server.MapPath(PhotoPath & FileExt_Item)
                    End If
                End If
                If rsPhoto("Inputer") <> UserName And rsPhoto("Status") <> Status And (Status = -2 Or Status = 3) Then
                    Call SendEmailOfCheck(rsPhoto("Inputer"), rsPhoto)
                End If
                Call UpdateUserData(0, rsPhoto("Inputer"), 0, 0)

                If rsPhoto("Status") = 0 And Status = 3 Then
                    Conn.Execute ("update PE_User set UserExp=UserExp+" & rsPhoto("PresentExp") & " where UserName='" & rsPhoto("Inputer") & "'")
                End If
                If rsPhoto("Status") = 3 And Status <> 3 Then
                    Conn.Execute ("update PE_User set UserExp=UserExp-" & rsPhoto("PresentExp") & " where UserName='" & rsPhoto("Inputer") & "'")
                End If


                Dim rsInfo, sqlInfo, j
                i = 0
                sqlInfo = "select * from PE_InfoS where ModuleType=" & ModuleType & " and ItemID=" & PhotoID & " order by InfoID"
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
                            Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (" & ModuleType & "," & PhotoID & "," & PE_CLng(arrSpecialID(j)) & ")")
                        End If
                    Next
                End If
            End If
        End If
    End If
    If AddType <> 3 Then
        rsPhoto("ClassID") = ClassID
        rsPhoto("PhotoName") = PhotoName
        rsPhoto("Keyword") = Keyword
        rsPhoto("Author") = Author
        rsPhoto("CopyFrom") = CopyFrom
        rsPhoto("PhotoIntro") = PhotoIntro
        rsPhoto("PhotoThumb") = PhotoThumb
        rsPhoto("PhotoUrl") = PhotoUrl
        rsPhoto("Hits") = PE_CLng(Trim(Request.Form("Hits")))
        rsPhoto("DayHits") = PE_CLng(Trim(Request.Form("DayHits")))
        rsPhoto("WeekHits") = PE_CLng(Trim(Request.Form("WeekHits")))
        rsPhoto("MonthHits") = PE_CLng(Trim(Request.Form("MonthHits")))
        rsPhoto("Stars") = PE_CLng(Trim(Request.Form("Stars")))
        rsPhoto("UpdateTime") = UpdateTime
        rsPhoto("CreateTime") = UpdateTime
        rsPhoto("Status") = Status
        rsPhoto("OnTop") = PE_CBool(Trim(Request.Form("OnTop")))
        rsPhoto("Elite") = PE_CBool(Trim(Request.Form("Elite")))
        'rsPhoto("Inputer") = Inputer
        rsPhoto("Editor") = Editor
        rsPhoto("SkinID") = PE_CLng(Trim(Request.Form("SkinID")))
        rsPhoto("TemplateID") = PE_CLng(Trim(Request.Form("TemplateID")))
        rsPhoto("Deleted") = False
        rsPhoto("PresentExp") = PresentExp
        rsPhoto("InfoPurview") = PE_CLng(Trim(Request.Form("InfoPurview")))
        rsPhoto("arrGroupID") = ReplaceBadChar(Trim(Request.Form("GroupID")))
        rsPhoto("InfoPoint") = PE_CLng(Trim(Request.Form("InfoPoint")))
        rsPhoto("ChargeType") = PE_CLng(Trim(Request.Form("ChargeType")))
        rsPhoto("PitchTime") = PE_CLng(Trim(Request.Form("PitchTime")))
        rsPhoto("ReadTimes") = PE_CLng(Trim(Request.Form("ReadTimes")))
        rsPhoto("DividePercent") = PE_CLng(Trim(Request.Form("DividePercent")))

        rsPhoto("VoteID") = VoteID
        If Not (rsField.BOF And rsField.EOF) Then
            rsField.MoveFirst
            Do While Not rsField.EOF
                If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                    rsPhoto(Trim(rsField("FieldName"))) = Trim(Request(rsField("FieldName")))
                End If
                rsField.MoveNext
            Loop
        End If
        Set rsField = Nothing
        
        rsPhoto.Update
        rsPhoto.Close

    End If
    Set rsPhoto = Nothing
    Call UpdateChannelData(ChannelID)
    If Action = "SaveAdd" Then
        Call UpdateUserData(0, Inputer, 0, 0)
    End If

    Response.Write "<br><br>"
    Response.Write "<table class='border' align=center width='500' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "  <tr align=center>"
    Response.Write "    <td  height='22' colspan='3' align='center' class='title'> "
    If Action = "SaveAdd" Then
        Response.Write "<b>添加" & ChannelShortName & "成功</b>"
    Else
        Response.Write "<b>修改" & ChannelShortName & "成功</b>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='150' align='center' valign='top' rowspan='5'>"
    If AddType < 3 Then
        Response.Write "<img src='" & GetPhotoThumb(PhotoThumb) & "' width='150'>"
    End If
    Response.Write "    </td>"
    Response.Write "    <td width='100' align='right' class='tdbg5'><strong>所属栏目：</td>"
    Response.Write "    <td width='250'>" & ShowClassPath() & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='100' align='right' class='tdbg5'><strong>" & ChannelShortName & "名称：</td>"
    Response.Write "    <td width='250'>" & PE_HTMLEncode(PhotoName) & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='100' align='right' class='tdbg5'><strong>" & ChannelShortName & "作者：</td>"
    Response.Write "    <td width='250'>" & PE_HTMLEncode(Author) & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='100' align='right' class='tdbg5'><strong>关 键 字：</td>"
    Response.Write "    <td width='250'>" & Mid(Keyword, 2, Len(Keyword) - 2) & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='100' align='right' class='tdbg5'><strong>" & ChannelShortName & "状态：</strong></td>"
    Response.Write "    <td width='250'>"
    If Status = -1 Then
        Response.Write "草稿"
    ElseIf Status = -2 Then
        Response.Write "退稿"
    Else
        Response.Write arrStatus(Status)
    End If
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='40' colspan='4' align='center'>"
    Response.Write "【<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Modify&PhotoID=" & PhotoID & "'>修改此" & ChannelShortName & "</a>】&nbsp;"
    Response.Write "【<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Add&AddType=" & AddType & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>继续添加" & ChannelShortName & "</a>】&nbsp;"
    Response.Write "【<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Manage&ClassID=" & ClassID & "'>" & ChannelShortName & "管理</a>】&nbsp;"
    Response.Write "【<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Show&PhotoID=" & PhotoID & "'>预览" & ChannelShortName & "内容</a>】"
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
        Response.Write "<br><iframe id='CreatePhoto' width='100%' height='210' frameborder='0' src='Admin_CreatePhoto.asp?ChannelID=" & ChannelID & "&Action=CreatePhoto2&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&PhotoID=" & PhotoID & "&ShowBack=No'></iframe>"
    End If
End Sub

Sub Show()
    If PhotoID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定" & ChannelShortName & "ID！</li>"
        Exit Sub
    End If
    
    Dim rsPhoto, PurviewChecked, PurviewChecked2
    PurviewChecked = False
    PurviewChecked2 = False
    Set rsPhoto = Conn.Execute("select * from PE_Photo where PhotoID=" & PhotoID & "")
    If rsPhoto.BOF And rsPhoto.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到" & ChannelShortName & "！</li>"
        rsPhoto.Close
        Set rsPhoto = Nothing
        Exit Sub
    End If
    ClassID = rsPhoto("ClassID")

    Call CheckClassPurview(Action, ClassID)
    If FoundErr = True Then
        rsPhoto.Close
        Set rsPhoto = Nothing
        Exit Sub
    End If

    Dim arrSpecialID
    arrSpecialID = GetSpecialIDArr(ModuleType, PhotoID)

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

    Response.Write "<br>您现在的位置：&nbsp;<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "'>" & ChannelShortName & "管理</a>&nbsp;&gt;&gt;&nbsp;"
    If ParentID > 0 Then
        Dim sqlPath, rsPath
        sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
        Set rsPath = Conn.Execute(sqlPath)
        Do While Not rsPath.EOF
            Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    Response.Write "<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;查看" & ChannelShortName & "信息："
    Response.Write rsPhoto("PhotoName") & "<br><br>"

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>图片信息</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>所属专题</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>收费信息</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>自定义选项</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf

    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "名称：</td>"
    Response.Write "  <td width='200'><strong>" & PE_HTMLEncode(rsPhoto("PhotoName")) & "</strong></td>"
    Response.Write "  <td rowspan='6' align=center valign='middle'>"
    If rsPhoto("PhotoThumb") = "" Then
        Response.Write "无缩略图"
    Else
        Response.Write "<img src='" & GetPhotoThumb(rsPhoto("PhotoThumb")) & "' width='150'><br>" & PE_HTMLEncode(rsPhoto("PhotoName")) & ""
    End If
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "作者：</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(rsPhoto("Author")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>更新时间：</td>"
    Response.Write "  <td width='200'>" & rsPhoto("UpdateTime") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>推荐等级：</td>"
    Response.Write "  <td width='200'>" & String(rsPhoto("Stars"), "★") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "添加：</td>"
    Response.Write "  <td width='200'>" & rsPhoto("Inputer") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>责任编辑：</td>"
    Response.Write "  <td width='200'>"
    If rsPhoto("Status") = 3 Then
        Response.Write rsPhoto("Editor")
    Else
        Response.Write "无"
    End If
    Response.Write "</td></tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>查看次数：</td>"
    Response.Write "  <td colspan='3'>本日：" & rsPhoto("DayHits") & "&nbsp;&nbsp;&nbsp;&nbsp;本周：" & rsPhoto("WeekHits") & "&nbsp;&nbsp;&nbsp;&nbsp;本月：" & rsPhoto("MonthHits") & "&nbsp;&nbsp;&nbsp;&nbsp;总计：" & rsPhoto("Hits") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "地址：</td>"
    Response.Write "  <td colspan='3'>"
    Call ShowPhotoUrls(rsPhoto("PhotoUrl"))
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td align='right' class='tdbg5'>" & ChannelShortName & "简介：</td>"
    Response.Write "  <td colspan='3'>" & FilterBadTag(rsPhoto("PhotoIntro"), rsPhoto("Inputer")) & "</td>"
    Response.Write "</tr>"
    Response.Write "        </tbody>" & vbCrLf

    Call ShowTabs_Special(arrSpecialID, " disabled")

    Call ShowTabs_Purview_Modify("查看", rsPhoto, " disabled")

    Call ShowTabs_MyField_View(rsPhoto)

    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf



    Response.Write "<form name='formA' method='get' action='Admin_Photo.asp'><p align='center'>"
    Response.Write "<input type='hidden' name='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input type='hidden' name='PhotoID' value='" & PhotoID & "'>"
    Response.Write "<input type='hidden' name='Action' value=''>" & vbCrLf

    If rsPhoto("Deleted") = False Then
        PurviewChecked = CheckClassPurview("Manage", ClassID)
        PurviewChecked2 = CheckClassPurview("Check", ClassID)
        If (rsPhoto("Inputer") = UserName And rsPhoto("Status") = 0) Or PurviewChecked = True Then
            Response.Write "<input type='submit' name='submit' value='修改/审核' onclick=""document.formA.Action.value='Modify'"">&nbsp;&nbsp;"
            Response.Write "<input type='submit' name='submit' value=' 删 除 ' onclick=""document.formA.Action.value='Del'"">&nbsp;&nbsp;"
        End If
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
            Response.Write "<input type='submit' name='submit' value=' 移 动 ' onclick=""document.formA.Action.value='MoveToClass'"">&nbsp;&nbsp;"
        End If
        If PurviewChecked2 = True Then
            If rsPhoto("Status") > -1 Then
                Response.Write "<input type='submit' name='submit' value='直接退稿' onclick=""document.formA.Action.value='Reject'"">&nbsp;&nbsp;"
            End If
            If rsPhoto("Status") < MyStatus Then
                Response.Write "<input type='submit' name='submit' value='" & arrStatus(MyStatus) & "' onclick=""document.formA.Action.value='SetPassed'"">&nbsp;&nbsp;"
            End If
            If rsPhoto("Status") >= MyStatus Then
                Response.Write "<input type='submit' name='submit' value='取消审核' onclick=""document.formA.Action.value='CancelPassed'"">&nbsp;&nbsp;"
            End If
        End If
        If PurviewChecked = True Then
            If rsPhoto("OnTop") = False Then
                Response.Write "<input type='submit' name='submit' value='设为固顶' onclick=""document.formA.Action.value='SetOnTop'"">&nbsp;&nbsp;"
            Else
                Response.Write "<input type='submit' name='submit' value='取消固顶' onclick=""document.formA.Action.value='CancelOnTop'"">&nbsp;&nbsp;"
            End If
            If rsPhoto("Elite") = False Then
                Response.Write "<input type='submit' name='submit' value='设为推荐' onclick=""document.formA.Action.value='SetElite'"">"
            Else
                Response.Write "<input type='submit' name='submit' value='取消推荐' onclick=""document.formA.Action.value='CancelElite'"">"
            End If
        End If
    Else
        If AdminPurview = 1 Or AdminPurview_Channel = 1 Then
            Response.Write "<input type='submit' name='submit' value='彻底删除' onclick=""if(confirm('确定要彻底删除此" & ChannelShortName & "吗？彻底删除后将无法还原！')==true){document.formA.Action.value='ConfirmDel';}"">&nbsp;&nbsp;"
            Response.Write "<input type='submit' name='submit' value=' 还 原 ' onclick=""document.formA.Action.value='Restore'"">"
        End If
    End If
    Response.Write "</p></form>"

    rsPhoto.Close
    Set rsPhoto = Nothing

    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='0'><tr class='tdbg'><td>"
    Response.Write "<li>上一" & ChannelItemUnit & ChannelShortName & "："
    Dim rsPrev
    Set rsPrev = Conn.Execute("Select Top 1 P.PhotoID,P.PhotoName,C.ClassID,C.ClassName from PE_Photo P left join PE_Class C On P.ClassID=C.ClassID Where P.ChannelID=" & ChannelID & " and P.Deleted=" & PE_False & " and P.PhotoID<" & PhotoID & " order by P.PhotoID desc")
    If rsPrev.EOF Then
        Response.Write "没有了"
    Else
        Response.Write "[<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&ClassID=" & rsPrev("ClassID") & "'>" & rsPrev("ClassName") & "</a>] <a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Show&PhotoID=" & rsPrev("PhotoID") & "'>" & rsPrev("PhotoName") & "</a>"
    End If
    rsPrev.Close
    Set rsPrev = Nothing
    Response.Write "</li></td</tr><tr class='tdbg'><td><li>下一" & ChannelItemUnit & ChannelShortName & "："
    Dim rsNext
    Set rsNext = Conn.Execute("Select Top 1 P.PhotoID,P.PhotoName,C.ClassID,C.ClassName from PE_Photo P left join PE_Class C On P.ClassID=C.ClassID Where P.ChannelID=" & ChannelID & " and P.Deleted=" & PE_False & " and P.PhotoID>" & PhotoID & " order by P.PhotoID asc")
    If rsNext.EOF Then
        Response.Write "没有了"
    Else
        Response.Write "[<a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&ClassID=" & rsNext("ClassID") & "'>" & rsNext("ClassName") & "</a>] <a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Show&PhotoID=" & rsNext("PhotoID") & "'>" & rsNext("PhotoName") & "</a>"
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
        Response.Write " class='title5' onclick=""window.location.href='Admin_Photo.asp?Action=Show&ChannelID=" & ChannelID & "&PhotoID=" & PhotoID & "&InfoType=0'"""
    End If
    Response.Write ">相关评论</td><td"
    If InfoType = 1 Then
        Response.Write " class='title6'"
    Else
        Response.Write " class='title5' onclick=""window.location.href='Admin_Photo.asp?Action=Show&ChannelID=" & ChannelID & "&PhotoID=" & PhotoID & "&InfoType=1'"""
    End If
    Response.Write ">相关收费</td>"
    Response.Write "<td>&nbsp;</td></tr></table>"
    
    strFileName = "Admin_Photo.asp?Action=Show&ChannelID=" & ChannelID & "&PhotoID=" & PhotoID & "&InfoType=" & InfoType
    
    Select Case InfoType
    Case 0
        Call ShowComment(PhotoID)
    Case 1
        Call ShowConsumeLog(PhotoID)
    End Select
End Sub

Sub Preview()
    Response.Write "<br><table width='100%' border=0 align=center cellPadding=2 cellSpacing=1 bgcolor='#FFFFFF' class='border' style='WORD-BREAK: break-all'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='4'>"

    If ClassID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定所属栏目</li>"
        Exit Sub
    ElseIf ClassID > 0 Then
        Set tClass = Conn.Execute("select ClassName,ParentID,ParentPath from PE_Class where ClassID=" & ClassID)
        If tClass.BOF And tClass.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>找不到指定的栏目</li>"
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

    Response.Write PE_HTMLEncode(Request("PhotoName"))
    Response.Write " </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "名称：</td>"
    Response.Write "  <td colspan='3'><strong>" & PE_HTMLEncode(Request("PhotoName")) & "</strong></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "作者：</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(Request("Author")) & "</td>"
    Response.Write "  <td colspan='2' rowspan='4' align=center valign='middle'>"
    If Request("PhotoThumb") = "" Then
        Response.Write "无缩略图"
    Else
        Response.Write "<img src='" & GetPhotoThumb(Request("PhotoThumb")) & "' width='150'>"
    End If
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>添加时间：</td>"
    Response.Write "  <td width='200'>" & Now() & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>评分等级：</td>"
    Response.Write "  <td width='200'>" & String(Request("Stars"), "★") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>查看" & PointName & "数：</td>"
    Response.Write "  <td width='200'><font color=red> " & Request("InfoPoint") & "</font> " & PointUnit & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "地址：</td>"
    Response.Write "  <td colspan='3'>"
    Call ShowPhotoUrls(Request("PhotoUrl"))
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "简介：</td>"
    Response.Write "  <td height='100' colspan='3'>" & Request("PhotoIntro") & "</td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    Response.Write "<p align='center'>【<a href='javascript:window.close();'>关闭窗口</a>】</p>"
End Sub


'******************************************************************************************
'以下为批量设置属性使用的函数，各模块实现过程类似，修改时注意同时修改各模块内容。
'******************************************************************************************

Sub Batch()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
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

    
    PhotoID = Replace(PhotoID, " ", "")
    Response.Write "<form method='POST' name='myform' action='Admin_Photo.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><b>批量修改" & ChannelShortName & "属性</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='center'>"
    Response.Write "      <td class='tdbg' valign='top'>"
    Response.Write "        <table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <tr>"
    Response.Write "            <td align='center'><b>" & ChannelShortName & "范围</b></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <input type='radio' name='BatchType' value='1' checked>指定" & ChannelShortName & "ID：<br>"
    Response.Write "              <input type='text' name='BatchPhotoID' value='" & PhotoID & "' size='28'><br>"
    Response.Write "              <input type='radio' name='BatchType' value='2'>指定栏目的" & ChannelShortName & "：<br>"
    Response.Write "              <select name='BatchClassID' size='2' multiple style='height:280px;width:180px;'>" & GetClass_Option(0, 0) & "</select><br><div align='center'>"
    Response.Write "      <input type='button' name='Submit' value='  选定所有栏目  ' onclick='SelectAll()'><br>"
    Response.Write "      <input type='button' name='Submit' value='取消选定所有栏目' onclick='UnSelectAll()'></div></td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "      <td valign='top' align='left'><br>"
    
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>属性设置</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>收费选项</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>自定义选项</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyCommentLink' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "标题：</td>"
    Response.Write "            <td><input name='ShowCommentLink' type='checkbox' id='ShowCommentLink' value='Yes'> 列表显示时在标题旁显示“评论”链接"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyAuthor' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "作者：</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Author' type='text' id='Author' value='" & Trim(Session("Author")) & "' size='15' maxlength='30'> " & GetAuthorList("Admin", ChannelID, UserName)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyCopyFrom' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "来源：</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='CopyFrom' type='text' id='CopyFrom' value='" & Trim(Session("CopyFrom")) & "' size='15' maxlength='50'> " & GetCopyFromList("Admin", ChannelID)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Call ShowBatchCommon
    Response.Write "        </tbody>" & vbCrLf

    Call ShowTabs_Purview_Batch("查看")
    Call ShowTabs_MyField_Batch

    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "  <br><b>说明：</b><br>1、若要批量修改某个属性的值，请先选中其左侧的复选框，然后再设定属性值。<br>2、这里显示的属性值都是系统默认值，与所选" & ChannelShortName & "的已有属性无关<br>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "  <p align='center'>"
    Response.Write "    <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "    <input name='Action' type='hidden' id='Action' value='DoBatch'>"
    Response.Write "    <input name='add' type='submit'  id='Add' value=' 执行批处理 ' style='cursor:hand;'>&nbsp; "
    Response.Write "    <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p>"
    Response.Write "  <br>"
    Response.Write "</form>"
End Sub

Sub DoBatch()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    
    Dim rs, sql, BatchType, BatchPhotoID, BatchClassID, rsField
    Dim Author, CopyFrom
    Dim Keyword, OnTop, Elite, Stars, Hits, UpdateTime, SkinID, TemplateID
    Dim InfoPurview, arrGroupID, InfoPoint, ChargeType, PitchTime, ReadTimes, DividePercent
    
    BatchType = PE_CLng(Trim(Request("BatchType")))
    BatchPhotoID = Trim(Request.Form("BatchPhotoID"))
    BatchClassID = FilterArrNull(Trim(Request.Form("BatchClassID")), ",")
    Author = Trim(Request.Form("Author"))
    CopyFrom = Trim(Request.Form("CopyFrom"))
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
        If IsValidID(BatchPhotoID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要批量修改的" & ChannelShortName & "的ID</li>"
        End If
    Else
        If IsValidID(BatchClassID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要批量修改的" & ChannelShortName & "的栏目</li>"
        End If
    End If
    If Trim(Request("ModifyKeyword")) = "Yes" And Keyword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入" & ChannelShortName & "关键字</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    If Author = "" Then Author = XmlText("BaseText", "DefAuthor", "佚名")
    If CopyFrom = "" Then CopyFrom = XmlText("BaseText", "DefCopyFrom", "本站原创")

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
        sql = "select * from PE_Photo where ChannelID=" & ChannelID & " and PhotoID in (" & BatchPhotoID & ")"
    Else
        sql = "select * from PE_Photo where ChannelID=" & ChannelID & " and ClassID in (" & BatchClassID & ")"
    End If
    rs.Open sql, Conn, 1, 3
    Do While Not rs.EOF
        If Trim(Request("ModifyAuthor")) = "Yes" Then rs("Author") = Author
        If Trim(Request("ModifyCopyFrom")) = "Yes" Then rs("CopyFrom") = CopyFrom
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

        Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-3 or ChannelID=" & ChannelID & "")
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
    Call WriteSuccessMsg("批量修改" & ChannelShortName & "属性成功！", "Admin_Photo.asp?ChannelID=" & ChannelID)
    Call CreateAllJS
End Sub


'******************************************************************************************
'以下为设置固顶、推荐等属性使用的函数，各模块实现过程类似，修改时注意同时修改各模块内容。
'******************************************************************************************

Sub SetProperty()
    If PhotoID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先选定" & ChannelShortName & "！</li>"
        Exit Sub
    End If
    If Action = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
        Exit Sub
    End If
    
    Dim sqlProperty, rsProperty, arrUser
    arrUser = ""
    If InStr(PhotoID, ",") > 0 Then
        sqlProperty = "select * from PE_Photo where PhotoID in (" & PhotoID & ")"
    Else
        sqlProperty = "select * from PE_Photo where PhotoID=" & PhotoID
    End If
    Set rsProperty = Server.CreateObject("ADODB.Recordset")
    rsProperty.Open sqlProperty, Conn, 1, 3
    Do While Not rsProperty.EOF
        If CheckClassPurview(Action, rsProperty("ClassID")) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>对 " & rsProperty("PhotoName") & " 没有操作权限！</li>"
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
    Call WriteSuccessMsg("操作成功！", "Admin_Photo.asp?ChannelID=" & ChannelID)
    Call CreateAllJS
End Sub


'******************************************************************************************
'以下为移动至栏目、专题等操作使用的函数，各模块实现过程类似，修改时注意同时修改各模块内容。
'******************************************************************************************

Sub DoMoveToClass()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    
    Dim PhotoType, BatchPhotoID, BatchClassID
    Dim tChannelID, tClassID, tChannelDir, tUploadDir
    
    PhotoType = PE_CLng(Trim(Request("PhotoType")))
    BatchPhotoID = Trim(Request.Form("BatchPhotoID"))
    BatchClassID = FilterArrNull(Trim(Request.Form("BatchClassID")), ",")
    tChannelID = Trim(Request("tChannelID"))
    tClassID = Trim(Request("tClassID"))
    
    If PhotoType = 1 Then
        If IsValidID(BatchPhotoID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要批量移动的" & ChannelShortName & "的ID</li>"
        Else
            BatchPhotoID = ReplaceBadChar(BatchPhotoID)
        End If
    Else
        If IsValidID(BatchClassID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要批量移动的" & ChannelShortName & "的栏目</li>"
        Else
            BatchClassID = ReplaceBadChar(BatchClassID)
        End If
    End If
    If tChannelID = "" Then
        tChannelID = ChannelID
    Else
        tChannelID = PE_CLng(tChannelID)
        If tChannelID <> ChannelID Then
            If AdminPurview > 1 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<br><li>对不起，你的权限不够！</li>"
            Else
                Dim rsChannel
                Set rsChannel = Conn.Execute("select ChannelDir,UploadDir from PE_Channel where ChannelID=" & tChannelID & "")
                If rsChannel.BOF And rsChannel.EOF Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>找不到目标频道！</li>"
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
        ErrMsg = ErrMsg & "<li>请指定目标栏目！不能指定为外部栏目。</li>"
    Else
        tClassID = PE_CLng(tClassID)
        If tClassID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>目标栏目不允许添加" & ChannelShortName & "</li>"
        End If
    End If
    If FoundErr = True Then Exit Sub
    
    Dim rsBatchMove, sqlBatchMove, PhotoPath
    sqlBatchMove = "select P.PhotoID,PhotoThumb,PhotoUrl,P.UpdateTime,C.ParentDir,C.ClassDir  from PE_Photo P left join PE_Class C on P.ClassID=C.ClassID"
    If PhotoType = 1 Then
        sqlBatchMove = sqlBatchMove & " where P.ChannelID=" & ChannelID & " and P.PhotoID in (" & BatchPhotoID & ")"
    Else
        sqlBatchMove = sqlBatchMove & " where P.ChannelID=" & ChannelID & " and P.ClassID in (" & BatchClassID & ")"
    End If
    BatchPhotoID = ""
    Set rsBatchMove = Conn.Execute(sqlBatchMove)
    Do While Not rsBatchMove.EOF
        PhotoPath = HtmlDir & GetItemPath(StructureType, rsBatchMove("ParentDir"), rsBatchMove("ClassDir"), rsBatchMove("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsBatchMove("UpdateTime"), rsBatchMove("PhotoID"))
        If fso.FileExists(Server.MapPath(PhotoPath & FileExt_Item)) Then
            DelSerialFiles Server.MapPath(PhotoPath & FileExt_Item)
        End If
            
        If tChannelID <> ChannelID Then
            Call MoveUpFiles("缩略图|" & rsBatchMove("PhotoThumb") & "$$$" & rsBatchMove("PhotoUrl"), tChannelDir & "/" & tUploadDir)    '移动上传文件
        End If
        If BatchPhotoID = "" Then
            BatchPhotoID = rsBatchMove("PhotoID")
        Else
            BatchPhotoID = BatchPhotoID & "," & rsBatchMove("PhotoID")
        End If
        rsBatchMove.MoveNext
    Loop
    rsBatchMove.Close
    Set rsBatchMove = Nothing
    If BatchPhotoID <> "" Then
        Conn.Execute ("update PE_Photo set ChannelID=" & tChannelID & ",ClassID=" & tClassID & ",TemplateID=0,CreateTime=UpdateTime where PhotoID in (" & BatchPhotoID & ")")
    End If

    Call WriteSuccessMsg("成功将选定的" & ChannelShortName & "移动到目标频道的目标栏目中！", "Admin_Photo.asp?ChannelID=" & ChannelID & "")
    Call ClearSiteCache(0)
End Sub


Sub MoveUpFiles(strFiles, strTargetDir)
    On Error Resume Next
    Dim arrPhotoUrls, strTrueFile, arrUrls, strTrueDir, iTemp
    If strFiles = "" Or strTargetDir = "" Then Exit Sub
    
    If Not fso.FolderExists(Server.MapPath(InstallDir & strTargetDir)) Then fso.CreateFolder Server.MapPath(InstallDir & strTargetDir)
    
    arrPhotoUrls = Split(strFiles, "$$$")
    For iTemp = 0 To UBound(arrPhotoUrls)
        arrUrls = Split(arrPhotoUrls(iTemp), "|")
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
'以下为删除、清空、还原等操作使用的函数，各模块实现过程类似，修改时注意同时修改各模块内容。
'******************************************************************************************

Sub Del()
    If PhotoID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先选定" & ChannelShortName & "！</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel, PhotoPath, arrUser
    arrUser = ""
    sqlDel = "select P.PhotoID,P.PhotoName,P.UpdateTime,P.CreateTime,P.Inputer,P.Status,P.Deleted,P.PresentExp,P.ClassID,C.ParentDir,C.ClassDir from PE_Photo P left join PE_Class C on P.ClassID=C.ClassID"
    If InStr(PhotoID, ",") > 0 Then
        sqlDel = sqlDel & " where P.PhotoID in (" & PhotoID & ") order by P.PhotoID"
    Else
        sqlDel = sqlDel & " where P.PhotoID=" & PhotoID
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
            ErrMsg = ErrMsg & "<li>删除 <font color='red'>" & rsDel("PhotoName") & "</font> 失败！原因：没有操作权限！</li>"
        Else
            If FoundInArr(arrUser, rsDel("Inputer"), ",") = True Then
                If arrUser = "" Then
                    arrUser = rsDel("Inputer")
                Else
                    arrUser = arrUser & "," & rsDel("Inputer")
                End If
            End If
            PhotoPath = HtmlDir & GetItemPath(StructureType, rsDel("ParentDir"), rsDel("ClassDir"), rsDel("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsDel("UpdateTime"), rsDel("PhotoID"))
            If fso.FileExists(Server.MapPath(PhotoPath & FileExt_Item)) Then
                DelSerialFiles Server.MapPath(PhotoPath & FileExt_Item)
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
    Call WriteSuccessMsg("操作成功！", "Admin_Photo.asp?ChannelID=" & ChannelID)
    Call CreateAllJS
End Sub

Sub DelFile()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    If PhotoID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先选定" & ChannelShortName & "！</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel, PhotoPath
    PhotoID = ReplaceBadChar(PhotoID)
    sqlDel = "select P.PhotoID,P.UpdateTime,C.ParentDir,C.ClassDir from PE_Photo P left join PE_Class C on P.ClassID=C.ClassID where P.PhotoID in (" & PhotoID & ") order by P.PhotoID"
    Set rsDel = Conn.Execute(sqlDel)
    Do While Not rsDel.EOF
        PhotoPath = HtmlDir & GetItemPath(StructureType, rsDel("ParentDir"), rsDel("ClassDir"), rsDel("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsDel("UpdateTime"), rsDel("PhotoID"))
        If fso.FileExists(Server.MapPath(PhotoPath & FileExt_Item)) Then
            DelSerialFiles Server.MapPath(PhotoPath & FileExt_Item)
        End If
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    Conn.Execute ("update PE_Photo set CreateTime=UpdateTime where PhotoID in (" & PhotoID & ")")
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub ConfirmDel()
    If AdminPurview = 2 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    If PhotoID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先选定" & ChannelShortName & "！</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel
    sqlDel = "select PhotoThumb,PhotoUrl,VoteID from PE_Photo where PhotoID in (" & PhotoID & ")"
    Set rsDel = Conn.Execute(sqlDel)
    Do While Not rsDel.EOF
        Call DelUploadFiles(GetUploadFiles(rsDel("PhotoUrl"), rsDel("PhotoThumb")))
        Conn.Execute ("delete from PE_Vote where ID=" & rsDel("VoteID") & "")
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    Conn.Execute ("delete from PE_Photo where PhotoID in (" & PhotoID & ")")
    Conn.Execute ("delete from PE_Comment where ModuleType=" & ModuleType & " and InfoID in (" & PhotoID & ")")
    Conn.Execute ("delete from PE_ConsumeLog where ModuleType=" & ModuleType & " and InfoID in (" & PhotoID & ")")
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub ClearRecyclebin()
    If AdminPurview = 2 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel
    PhotoID = ""
    sqlDel = "select PhotoID,PhotoThumb,PhotoUrl,VoteID from PE_Photo where Deleted=" & PE_True & " and ChannelID=" & ChannelID
    Set rsDel = Conn.Execute(sqlDel)
    Do While Not rsDel.EOF
        If PhotoID = "" Then
            PhotoID = rsDel(0)
        Else
            PhotoID = PhotoID & "," & rsDel(0)
        End If
        Call DelUploadFiles(GetUploadFiles(rsDel("PhotoUrl"), rsDel("PhotoThumb")))
        Conn.Execute ("delete from PE_Vote where ID=" & rsDel("VoteID") & "")
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    If PhotoID <> "" Then
        Conn.Execute ("delete from PE_Photo where Deleted=" & PE_True & " and ChannelID=" & ChannelID & "")
        Conn.Execute ("delete from PE_Comment where ModuleType=" & ModuleType & " and InfoID in (" & PhotoID & ")")
        Conn.Execute ("delete from PE_ConsumeLog where ModuleType=" & ModuleType & " and InfoID in (" & PhotoID & ")")
    End If
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub Restore()
    If AdminPurview = 2 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    If PhotoID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先选定" & ChannelShortName & "！</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel, arrUser
    arrUser = ""
    If InStr(PhotoID, ",") > 0 Then
        sqlDel = "select * from PE_Photo where PhotoID in (" & PhotoID & ")"
    Else
        sqlDel = "select * from PE_Photo where PhotoID=" & PhotoID
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
    Call WriteSuccessMsg("操作成功！", "Admin_Photo.asp?ChannelID=" & ChannelID)
    Call CreateAllJS
End Sub

Sub RestoreAll()
    If AdminPurview = 2 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel, arrUser
    arrUser = ""
    sqlDel = "select * from PE_Photo where Deleted=" & PE_True & " and ChannelID=" & ChannelID
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
    Call WriteSuccessMsg("操作成功！", "Admin_Photo.asp?ChannelID=" & ChannelID)
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

Function GetUploadFiles(PhotoUrls, PhotoThumb)
    Dim arrPhotoUrls, arrUrls, iTemp, strUrls
    strUrls = ""
    If Left(PhotoThumb, 1) <> "/" And InStr(PhotoThumb, "://") <= 0 Then
        strUrls = strUrls & UploadDir & "/" & PhotoThumb
    End If
    arrPhotoUrls = Split(PhotoUrls, "$$$")
    For iTemp = 0 To UBound(arrPhotoUrls)
        arrUrls = Split(arrPhotoUrls(iTemp), "|")
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

'******************************************************************************************
'以下为各模块通用函数部分，仅个别几处不同（已在程序中注释），修改时注意同时修改各模块内容。
'******************************************************************************************

Function GetPath(RootName)
    Dim strPath
    strPath = "您现在的位置：&nbsp;" & ChannelName & "管理&nbsp;&gt;&gt;&nbsp;<a href='" & FileName & "'>" & RootName & "</a>&nbsp;&gt;&gt;&nbsp;"
    If ClassID > 0 Then
        If ParentID > 0 Then
            Dim sqlPath, rsPath
            sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
            Set rsPath = Conn.Execute(sqlPath)
            Do While Not rsPath.EOF
                strPath = strPath & "<a href='" & FileName & "&ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
                rsPath.MoveNext
            Loop
            rsPath.Close
            Set rsPath = Nothing
        End If
        strPath = strPath & "<a href='" & FileName & "&ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
    End If
    If ManageType = "My" Then   '此处各模块中不同
        strPath = strPath & "<font color=red>" & UserName & "</font> 添加的" & ChannelShortName & ""
    Else
        If Keyword = "" Then
            Select Case Status
            Case -2
                strPath = strPath & "退稿"
            Case -1
                strPath = strPath & "草稿"
            Case 0
                strPath = strPath & "待审核的" & ChannelShortName & "！"
            Case 1
                strPath = strPath & "已审核的" & ChannelShortName & "！"
            Case Else
                strPath = strPath & "所有" & ChannelShortName & "！"
            End Select
        Else
            Select Case strField
                Case "Title"
                    strPath = strPath & "标题中含有 <font color=red>" & Keyword & "</font> "
                Case "Content"
                    strPath = strPath & "内容中含有 <font color=red>" & Keyword & "</font> "
                Case "Author"
                    strPath = strPath & "作者姓名中含有 <font color=red>" & Keyword & "</font> "
                Case "Inputer"
                    strPath = strPath & "<font color=red>" & Keyword & "</font> 添加"
                Case "PhotoName", "PhotoName"
                    strPath = strPath & "名称中含有 <font color=red>" & Keyword & "</font> "
                Case "PhotoIntro", "PhotoIntro"
                    strPath = strPath & "内容中含有 <font color=red>" & Keyword & "</font> "
                Case Else
                    strPath = strPath & "名称中含有 <font color=red>" & Keyword & "</font> "
            End Select
            Select Case Status
            Case -2
                strPath = strPath & "的退稿"
            Case -1
                strPath = strPath & "的草稿"
            Case 0
                strPath = strPath & "并且未审核的" & ChannelShortName & "！"
            Case 1
                strPath = strPath & "并且已审核的" & ChannelShortName & "！"
            Case Else
                strPath = strPath & "的" & ChannelShortName & "！"
            End Select
        End If
    End If
    GetPath = strPath
End Function


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

Function GetPhotoThumb(PhotoThumb)
    If Left(PhotoThumb, 1) <> "/" And InStr(PhotoThumb, "://") <= 0 Then
        GetPhotoThumb = InstallDir & ChannelDir & "/" & UploadDir & "/" & PhotoThumb
    Else
        GetPhotoThumb = PhotoThumb
    End If
End Function

Sub ShowPhotoUrls(PhotoUrls)
    Dim arrPhotoUrls, arrUrls, iTemp
    If InStr(PhotoUrls, "$$$") > 1 Then
        arrPhotoUrls = Split(PhotoUrls, "$$$")
        For iTemp = 0 To UBound(arrPhotoUrls)
            arrUrls = Split(arrPhotoUrls(iTemp), "|")
            If UBound(arrUrls) = 1 Then
                If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                    Response.Write arrUrls(0) & "：<a href='" & InstallDir & ChannelDir & "/" & UploadDir & "/" & arrUrls(1) & "'>" & arrUrls(1) & "</a><br>"
                Else
                    Response.Write arrUrls(0) & "：<a href='" & arrUrls(1) & "'>" & arrUrls(1) & "</a><br>"
                End If
            End If
        Next
    Else
        arrUrls = Split(PhotoUrls, "|")
        If UBound(arrUrls) = 1 Then
            If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                Response.Write arrUrls(0) & "：<a href='" & InstallDir & ChannelDir & "/" & UploadDir & "/" & arrUrls(1) & "'>" & arrUrls(1) & "</a><br>"
            Else
                Response.Write arrUrls(0) & "：<a href='" & arrUrls(1) & "'>" & arrUrls(1) & "</a><br>"
            End If
        End If
    End If
End Sub
%>
