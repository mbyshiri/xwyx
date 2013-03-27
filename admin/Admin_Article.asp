<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_Content.asp"-->
<!--#include file="../Include/PowerEasy.CreateThumb.asp"-->
<!--#include file="../Include/PowerEasy.SendMail.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<!--#include file="../Include/PowerEasy.XmlHttp.asp"-->
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

Dim HtmlDir, PurviewChecked
Dim ManageType, Status, MyStatus, arrStatus
Dim ClassID, SpecialID, OnTop, IsElite, IsHot, Created
Dim tClass, ClassName, RootID, ParentID, Depth, ParentPath, Child, arrChildID, ParentDir, ClassDir, ClassPurview

Dim ArticleID

Dim PayStatus
Dim IncludePic, UploadFiles, DefaultPicUrl, IsThumb
Dim ArticlePro1, ArticlePro2, ArticlePro3, ArticlePro4
Dim VoteID

If ChannelID = 0 Then
    Response.Write "频道参数丢失！"
    Call CloseConn
    Response.End
End If
If ModuleType <> 1 Then
    Response.Write "<li>指定的频道ID不对！</li>"
    Call CloseConn
    Response.End
End If
ModuleName = "Article"
SheetName = "PE_Article"


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
ArticleID = Trim(Request("ArticleID"))
PayStatus = Trim(Request("PayStatus")) '文章支付状态

If Action = "" Then
    Action = "Manage"
End If
If Status = "" Then
    Status = 9
Else
    Status = PE_CLng(Status) '文章状态   9－－所有文章，-1－－草稿，0－－待审核，1－－已审核，-2－－退稿
End If
If IsValidID(ArticleID) = False Then
    ArticleID = ""
End If
If PayStatus = "" Then
    PayStatus = "False"
End If

FileName = "Admin_Article.asp?ChannelID=" & ChannelID & "&Action=" & Action & "&ManageType=" & ManageType
strFileName = FileName & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&Field=" & strField & "&keyword=" & Keyword

If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))
ArticlePro1 = XmlText("Article", "ArticlePro1", "[图文]")
ArticlePro2 = XmlText("Article", "ArticlePro2", "[组图]")
ArticlePro3 = XmlText("Article", "ArticlePro3", "[推荐]")
ArticlePro4 = XmlText("Article", "ArticlePro4", "[注意]")


If Action = "ExportExcel" Then
    Call ExportExcel
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
    Response.End
End If
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
Case "SaveAdd", "SaveModify", "SaveModifyAsAdd"
    strTitle = strTitle & "保存" & ChannelShortName
Case "Move"
    strTitle = strTitle & "移动" & ChannelShortName
Case "Preview", "Show"
    strTitle = strTitle & "预览" & ChannelShortName
Case "Batch", "DoBatch"
    strTitle = strTitle & "批量修改" & ChannelShortName & "属性"
Case "MoveToClass"
    strTitle = strTitle & "批量移动" & ChannelShortName
Case "BatchReplace"
    strTitle = strTitle & "批量替换" & ChannelShortName
Case "AddToSpecial"
    strTitle = strTitle & "添加" & ChannelShortName & "到专题"
Case "MoveToSpecial"
    strTitle = strTitle & "移动" & ChannelShortName & "到专题"
Case "Manage"
    Select Case ManageType
    Case "Check"
        strTitle = strTitle & ChannelShortName & "审核"
    Case "PayMoney"
        strTitle = strTitle & ChannelShortName & "稿费管理"
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
Call ShowPageTitle(strTitle, 10111)

Response.Write "  <tr class='tdbg'>"
Response.Write "    <td width='70' height='30' ><strong>管理导航：</strong></td><td colspan='5'>"
Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Status=9'>" & ChannelShortName & "管理首页</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>添加" & ChannelShortName & "</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ManageType=Check&Status=0'>审核" & ChannelShortName & "</a>"
If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ManageType=Special'>专题" & ChannelShortName & "管理</a>"
End If
If AdminPurview = 1 Or AdminPurview_Channel = 1 Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ManageType=Recyclebin&Status=9'>" & ChannelShortName & "回收站管理</a>"
End If
If FoundInArr(arrEnabledTabs, "Copyfee", ",") = True Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&ManageType=PayMoney&PayStatus=False' target=main>稿费管理</a>"
End If
If UseCreateHTML > 0 And ObjInstalled_FSO = True Then
    Response.Write "&nbsp;|&nbsp;<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ManageType=HTML&Status=1'><b>生成HTML管理</b></a>"
End If
Response.Write "</td></tr>" & vbCrLf

If Action = "Manage" Then
    Response.Write "<form name='form3' method='Post' action='" & strFileName & "'><tr class='tdbg'>"
    Response.Write "  <td width='70' height='30' ><strong>" & ChannelShortName & "选项：</strong></td><td>"
    If ManageType = "PayMoney" Then
        Response.Write "<input name='PayStatus' type='radio' onclick='submit();' " & RadioValue(PayStatus, "False") & ">未支付稿费的" & ChannelShortName & "&nbsp;&nbsp;&nbsp"
        Response.Write "<input name='PayStatus' type='radio' onclick='submit();' " & RadioValue(PayStatus, "True") & ">已支付稿费的" & ChannelShortName
    ElseIf ManageType = "HTML" Then
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

strFileName = strFileName & "&Status=" & Status & "&Created=" & Created & "&PayStatus=" & PayStatus & "&OnTop=" & OnTop & "&IsElite=" & IsElite & "&IsHot=" & IsHot

Select Case Action
Case "Add"
    Call Add
Case "Modify", "Check"
    Call Modify
Case "SaveAdd", "SaveModify", "SaveModifyAsAdd"
    Call SaveArticle
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
Case "BatchReplace"
    Call BatchReplace
Case "DoBatchReplace"
    Call DoBatchReplace
Case "Manage"
    Call main
Case "ConfirmPay"
    Call ConfirmPay
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

    Dim arrUser, i, NotReceiveUser
    Dim rsArticleList, sql, Querysql
    PurviewChecked = False
    If ClassID = 0 Then
        If strField = "" And AdminPurview = 2 And AdminPurview_Channel = 3 And ManageType <> "My" Then
            If ManageType = "Check" Or ManageType = "PayMoney" Then
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
        If ManageType = "Check" Or ManageType = "PayMoney" Then
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
    Case "PayMoney"
        Call ShowContentManagePath("稿费管理")
    Case "Receive"
        Call ShowContentManagePath(ChannelShortName & "签收管理")
    Case "HTML"
        Call ShowContentManagePath(ChannelShortName & "生成")
    Case "Recyclebin"
        Call ShowContentManagePath(ChannelShortName & "回收站管理")
    Case "Special"
        Call ShowContentManagePath("专题" & ChannelShortName & "管理")
    Case Else
        Call ShowContentManagePath(ChannelShortName & "管理")
    End Select

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "    <form name='myform' method='Post' action='Admin_Article.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "     <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "          <tr class='title' height='22'> "
    Response.Write "            <td height='22' width='30' align='center'><strong>选中</strong></td>"
    Response.Write "            <td width='25' align='center'><strong>ID</strong></td>"
    If ManageType = "Special" Then
        Response.Write "        <td width='120' align='center'><strong>所属专题</strong></td>"
    End If
    Response.Write "            <td align='center' ><strong>" & ChannelShortName & "标题</strong></td>"
    Response.Write "            <td width='60' align='center' ><strong>录入者</strong></td>"
      '添加稿费管理界面
    If ManageType = "PayMoney" Then
        Response.Write "        <td width='60' align='center'><strong>作者</strong></td>"
        Response.Write "        <td width='80' align='center'><strong>稿费受益者</strong></td>"
        Response.Write "        <td width='60' align='center'><strong>稿费</strong></td>"
        Response.Write "        <td width='40' align='center'><strong>已支付</strong></td>"
        If PayStatus = "True" Then
            Response.Write "        <td width='60' align='center' ><strong>支付日期</strong></td>"
        Else
            Response.Write "        <td width='60' align='center' ><strong>录入日期</strong></td>"
        End If
    Else
        Response.Write "            <td width='40' align='center' ><strong>点击数</strong></td>"
        Response.Write "            <td width='80' align='center' ><strong>" & ChannelShortName & "属性</strong></td>"
        Response.Write "            <td width='60' align='center' ><strong>审核状态</strong></td>"
    End If
    If UseCreateHTML > 0 And ObjInstalled_FSO = True And ManageType <> "Special" Then
        Response.Write "            <td width='40' align='center' ><strong>已生成</strong></td>"
    End If
    If ManageType = "Check" Then
        Response.Write "            <td width='120' align='center' ><strong>审核操作</strong></td>"
    ElseIf ManageType = "PayMoney" Then
        If PayStatus = "False" Then
            Response.Write "            <td width='60' align='center' ><strong>稿费操作</strong></td>"
        End If
    ElseIf ManageType = "HTML" Then
        Response.Write "            <td width='180' align='center' ><strong>生成HTML操作</strong></td>"
    ElseIf ManageType = "Recyclebin" Then
        Response.Write "            <td width='100' align='center' ><strong>回收站操作</strong></td>"
    ElseIf ManageType = "Special" Then
        Response.Write "            <td width='100' align='center' ><strong>专题管理操作</strong></td>"
    Else
        Response.Write "            <td width='150' align='center' ><strong>常规管理操作</strong></td>"
    End If
    Response.Write "          </tr>"

    If ManageType = "Special" Then
        sql = "select top " & MaxPerPage & " I.InfoID,I.SpecialID,A.ArticleID,SP.SpecialName,A.Title,A.Keyword,A.Author,A.UpdateTime,A.Inputer,"
        sql = sql & "A.CopyFrom,A.DefaultPicUrl,A.IncludePic,A.PaginationType,A.Receive,A.ReceiveUser,A.Received,"
        sql = sql & "A.Hits,A.OnTop,A.Elite,A.Status,A.Stars,A.InfoPoint,A.VoteID"
        sql = sql & " from PE_Article A right join (PE_InfoS I left join PE_Special SP on I.SpecialID=SP.SpecialID) on A.ArticleID=I.ItemID "
    Else
        If ClassID = -1 Or (ClassID > 0 And Child = 0) Then
            sql = "select top " & MaxPerPage & " A.ClassID,A.ArticleID,A.Title,A.Keyword,A.Author,A.UpdateTime,A.Inputer,"
            sql = sql & "A.CopyFrom,A.IncludePic,A.DefaultPicUrl,A.PaginationType,A.Receive,A.ReceiveUser,A.Received,"
            sql = sql & "A.Hits,A.OnTop,A.Elite,A.Status,A.Stars,A.InfoPoint,A.Beneficiary,A.IsPayed,A.CopyMoney,A.PayDate,A.VoteID"
            sql = sql & " from PE_Article A "
        Else
            sql = "select top " & MaxPerPage & " A.ClassID,A.ArticleID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview,A.Title,A.Keyword,A.Author,A.UpdateTime,A.Inputer,"
            sql = sql & "A.CopyFrom,A.IncludePic,A.DefaultPicUrl,A.PaginationType,A.Receive,A.ReceiveUser,A.Received,"
            sql = sql & "A.Hits,A.OnTop,A.Elite,A.Status,A.Stars,A.InfoPoint,A.Beneficiary,A.IsPayed,A.CopyMoney,A.PayDate,A.VoteID"
            sql = sql & " from PE_Article A left join PE_Class C on A.ClassID=C.ClassID "
        End If
    End If
    
    Querysql = " where A.ChannelID=" & ChannelID
    If ManageType = "Special" Then
        Querysql = Querysql & " and I.ModuleType=" & ModuleType
    End If
    If ManageType = "Receive" Then
        Querysql = Querysql & " and A.Receive=" & PE_True & ""
    End If
    If ManageType = "Recyclebin" Then
        Querysql = Querysql & " and A.Deleted=" & PE_True & ""
    Else
        Querysql = Querysql & " and A.Deleted=" & PE_False & ""
    End If
    If ManageType = "HTML" Then
        If Created = "False" Then
            Querysql = Querysql & " and (A.CreateTime<=A.UpdateTime or A.CreateTime is Null)"
        ElseIf Created = "True" Then
            Querysql = Querysql & " and A.CreateTime>A.UpdateTime"
        End If
        Querysql = Querysql & " and A.Status=3" '当文章为已审核时，才在生成管理中出现
    ElseIf ManageType = "PayMoney" Then
       '如果是稿费管理，则查出所有的以通过审查但是却没有被删除和计算稿费的文章
        If PayStatus = "False" Then
            Querysql = Querysql & " and A.Status=3 and A.CopyMoney>0 and  A.IsPayed=" & PE_False & "" '查询出计算过，但是没有被支付的文章
        ElseIf PayStatus = "True" Then
            Querysql = Querysql & " and A.Status=3 and A.CopyMoney>0 and A.IsPayed=" & PE_True & "" '查询出被计算了，并且支付过的文章
        End If
    Else
        Select Case Status
        Case -2 '退稿
            Querysql = Querysql & " and A.Status=-2"
        Case -1 '草稿
            Querysql = Querysql & " and A.Status=-1"
        Case 0  '待审核
            Querysql = Querysql & " and A.Status>=0 and A.Status<" & MyStatus
        Case 1  '已审核
            Querysql = Querysql & " and A.Status>=" & MyStatus
        Case Else
            Querysql = Querysql & " and A.Status>-1"
        End Select
        If OnTop = "True" Then
            Querysql = Querysql & " and A.OnTop=" & PE_True & ""
        End If
        If IsElite = "True" Then
            Querysql = Querysql & " and A.Elite=" & PE_True & ""
        End If
        If IsHot = "True" Then
            Querysql = Querysql & " and A.Hits>=" & HitsOfHot & ""
        End If
    End If

    If ClassID <> 0 Then
        If Child > 0 Then
            Querysql = Querysql & " and A.ClassID in (" & arrChildID & ")"
        Else
            Querysql = Querysql & " and A.ClassID=" & ClassID
        End If
    End If
    If SpecialID > 0 Then
        Querysql = Querysql & " and I.SpecialID=" & SpecialID
    End If
    If ManageType = "My" Then
        Querysql = Querysql & " and A.Inputer='" & UserName & "' "
    End If
    If Keyword <> "" Then
        Select Case strField
        Case "Title"
            Querysql = Querysql & " and A.Title like '%" & Keyword & "%' "
        Case "Content"
            Querysql = Querysql & " and A.Content like '%" & Keyword & "%' "
        Case "Author"
            Querysql = Querysql & " and A.Author like '%" & Keyword & "%' "
        Case "Inputer"
            Querysql = Querysql & " and A.Inputer='" & Keyword & "' "
        Case "Editor"
            Querysql = Querysql & " and A.Editor='" & Keyword & "' "
        Case "UpdateTime"
            Querysql = Querysql & " and DateDiff(" & PE_DatePart_D & ",A.UpdateTime,'" & Keyword & "')=0 "
        Case "Keyword"
            Querysql = Querysql & " and A.Keyword like '%|" & Keyword & "|%' "
        Case "ID"
            Querysql = Querysql & " and A.ArticleID=" & PE_Clng(Keyword) & " "
        Case Else
            Querysql = Querysql & " and A.Title like '%" & Keyword & "%' "
        End Select
    End If
    If ManageType = "Special" Then
        totalPut = PE_CLng(Conn.Execute("select Count(*) from PE_InfoS I inner join PE_Article A on I.ItemID=A.ArticleID " & Querysql)(0))
    Else
        totalPut = PE_CLng(Conn.Execute("select Count(*) from PE_Article A " & Querysql)(0))
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
            Querysql = Querysql & " and I.InfoID < (select min(InfoID) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " I.InfoID from PE_InfoS I inner join PE_Article A on I.ItemID=A.ArticleID " & Querysql & " order by I.InfoID desc) as QueryArticle)"
        Else
            Querysql = Querysql & " and A.ArticleID < (select min(ArticleID) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " A.ArticleID from PE_Article A " & Querysql & " order by A.ArticleID desc) as QueryArticle)"
        End If
    End If
    If ManageType = "Special" Then
        sql = sql & Querysql & " order by I.InfoID desc"
    Else
        sql = sql & Querysql & " order by A.ArticleID desc"
    End If

    Set rsArticleList = Server.CreateObject("ADODB.Recordset")
    rsArticleList.Open sql, Conn, 1, 1
    If rsArticleList.BOF And rsArticleList.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>"
        If ClassID > 0 Then
            Response.Write "此栏目及其子栏目中没有任何"
        Else
            Response.Write "没有任何"
        End If
        If ManageType = "PayMoney" Then
            Select Case PayStatus
            Case "True"
                Response.Write "<font color=blue>已付稿费</font>的" & ChannelShortName & "！"
            Case "False"
                Response.Write "<font color=green>未支付稿费</font>" & ChannelShortName & "！"
            'Case Else
              '  Response.Write "需要支付稿费的" & ChannelShortName & "！"
            End Select
        Else
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
        End If
        Response.Write "<br><br></td></tr>"
    Else
        Dim ArticleNum, ArticlePath
        ArticleNum = 0
        Do While Not rsArticleList.EOF
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            If ManageType = "Special" Then
                Response.Write "        <td width='30' align='center'><input name='InfoID' type='checkbox' onclick='CheckItem(this)' id='InfoID' value='" & rsArticleList("InfoID") & "'></td>"
                Response.Write "        <td width='25' align='center'>" & rsArticleList("InfoID") & "</td>"
                Response.Write "        <td width='120' align='center'>"
                If rsArticleList("SpecialID") > 0 Then
                    Response.Write "<a href='" & FileName & "&SpecialID=" & rsArticleList("SpecialID") & "'>" & rsArticleList("SpecialName") & "</a>"
                Else
                    Response.Write "&nbsp;"
                End If
                Response.Write "</td>"
            Else
                Response.Write "        <td width='30' align='center'><input name='ArticleID' type='checkbox' onclick='CheckItem(this)' id='ArticleID' value='" & rsArticleList("ArticleID") & "'></td>"
                Response.Write "        <td width='25' align='center'>" & rsArticleList("ArticleID") & "</td>"
            End If
            Response.Write "        <td>"
            If ManageType <> "Special" Then
                If rsArticleList("ClassID") <> ClassID And ClassID <> -1 Then
                    Response.Write "<a href='" & FileName & "&ClassID=" & rsArticleList("ClassID") & "'>["
                    If rsArticleList("ClassName") <> "" Then
                        Response.Write rsArticleList("ClassName")
                    Else
                        Response.Write "<font color='gray'>不属于任何栏目</font>"
                    End If
                    Response.Write "]</a>&nbsp;"
                End If
            End If
            
            Select Case rsArticleList("IncludePic")
                Case 1
                    Response.Write "<font color=blue>" & ArticlePro1 & "</font>"
                Case 2
                    Response.Write "<font color=blue>" & ArticlePro2 & "</font>"
                Case 3
                    Response.Write "<font color=blue>" & ArticlePro3 & "</font>"
                Case 4
                    Response.Write "<font color=blue>" & ArticlePro4 & "</font>"
            End Select
            
            Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & rsArticleList("ArticleID") & "'"
            Response.Write " title='标&nbsp;&nbsp;&nbsp;&nbsp;题：" & rsArticleList("Title") & vbCrLf & "作&nbsp;&nbsp;&nbsp;&nbsp;者：" & rsArticleList("Author") & vbCrLf & "转 贴 自：" & rsArticleList("CopyFrom") & vbCrLf & "更新时间：" & rsArticleList("UpdateTime") & vbCrLf
            Response.Write "点 击 数：" & rsArticleList("Hits") & vbCrLf & "关 键 字：" & Mid(rsArticleList("Keyword"), 2, Len(rsArticleList("Keyword")) - 2) & vbCrLf & "推荐等级："
            If rsArticleList("Stars") = 0 Then
                Response.Write "无"
            Else
                Response.Write String(rsArticleList("Stars"), "★")
            End If
            Response.Write vbCrLf & "分页方式："
            If rsArticleList("PaginationType") = 0 Then
                Response.Write "不分页"
            ElseIf rsArticleList("PaginationType") = 1 Then
                Response.Write "自动分页"
            ElseIf rsArticleList("PaginationType") = 2 Then
                Response.Write "手动分页"
            End If
            Response.Write vbCrLf & "阅读点数：" & rsArticleList("InfoPoint")
            Response.Write "'>" & rsArticleList("title") & "</a>"
            If ManageType = "Receive" And rsArticleList("Receive") = True Then
                Response.Write "&nbsp;"
                
                If rsArticleList("Received") = "" Then
                    NotReceiveUser = rsArticleList("ReceiveUser")
                Else
                    NotReceiveUser = ""
                    arrUser = Split(rsArticleList("ReceiveUser"), ",")
                    For i = 0 To UBound(arrUser)
                        If FoundInArr(rsArticleList("Received"), arrUser(i), "|") = False Then
                            If NotReceiveUser = "" Then
                                NotReceiveUser = arrUser(i)
                            Else
                                NotReceiveUser = NotReceiveUser & "," & arrUser(i)
                            End If
                        End If
                    Next
                End If
                Response.Write "<a href='' onclick='return false' title='"
                Response.Write "要求签收用户：" & rsArticleList("ReceiveUser") & vbCrLf
                Response.Write "已经签收用户：" & rsArticleList("Received") & vbCrLf
                Response.Write "尚未签收用户：" & NotReceiveUser
                If NotReceiveUser <> "" Then
                    Response.Write "'><font color=red>[签收中]</font></a>"
                Else
                    Response.Write "'><font color=green>[已签收]</font></a>"
                End If
            End If
            Response.Write "</td>"
            Response.Write "      <td width='60' align='center'><a href='" & FileName & "&field=Inputer&keyword=" & rsArticleList("Inputer") & "' title='点击将查看此用户录入的所有" & ChannelShortName & "'>" & rsArticleList("Inputer") & "</a></td>"
               '修改审核管理界面
            If ManageType = "PayMoney" Then
                Response.Write "      <td width='60' align='center'>" & rsArticleList("Author") & "</td>"
                Response.Write "      <td width='80' align='center'>" & rsArticleList("Beneficiary") & "</td>"
                Response.Write "      <td width='60' align='center'>" & FormatNumber(rsArticleList("CopyMoney"), 2, vbTrue, vbFalse, vbTrue) & "</td>"
                If rsArticleList("Ispayed") = True Then
                    Response.Write "      <td width='40' align='center'><b>√</b></td>"
                Else
                    Response.Write "      <td width='40' align='center'><font color=red><b>×</b></font></td>"
                End If
                If PayStatus Then
                    Response.Write "      <td width='60' align='center'>" & rsArticleList("payDate") & "</td>"
                Else
                    Response.Write "      <td width='60' align='center'>" & rsArticleList("UpdateTime") & "</td>"
                End If
            Else
                Response.Write "      <td width='40' align='center'>" & rsArticleList("Hits") & "</td>"
                Response.Write "      <td width='80' align='center'>"
                If rsArticleList("OnTop") = True Then
                    Response.Write "<font color=blue>顶</font> "
                Else
                    Response.Write "&nbsp;&nbsp;&nbsp;"
                End If
                If rsArticleList("Hits") >= HitsOfHot Then
                    Response.Write "<font color=red>热</font> "
                Else
                    Response.Write "&nbsp;&nbsp;&nbsp;"
                End If
                If rsArticleList("Elite") = True Then
                    Response.Write "<font color=green>荐</font> "
                Else
                    Response.Write "&nbsp;&nbsp;&nbsp;"
                End If
                If Trim(rsArticleList("DefaultPicUrl")) <> "" Then
                    Response.Write "<font color=blue>图</font>"
                Else
                    Response.Write "&nbsp;&nbsp;"
                End If
                If rsArticleList("VoteID") > 0 Then
                    Response.Write "<a href='" & InstallDir & "Vote.asp?ID=" & rsArticleList("VoteID") & "&Action=Show' target='_blank'>调</a>"
                Else
                    Response.Write "&nbsp;&nbsp;"
                End If
                Response.Write "    </td>"
                Response.Write "    <td width='60' align='center'>"
                Select Case rsArticleList("Status")
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
                Response.Write "</td>"
            End If
        
            Dim iClassPurview
            If UseCreateHTML > 0 And ObjInstalled_FSO = True And ManageType <> "Special" Then
                Response.Write "    <td width='40' align='center'>"
                If ClassID = -1 Or (ClassID > 0 And Child = 0) Then
                    iClassPurview = ClassPurview
                    ArticlePath = HtmlDir & GetItemPath(StructureType, ParentDir, ClassDir, rsArticleList("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsArticleList("UpdateTime"), rsArticleList("ArticleID")) & FileExt_Item
                Else
                    iClassPurview = PE_CLng(rsArticleList("ClassPurview"))
                    ArticlePath = HtmlDir & GetItemPath(StructureType, rsArticleList("ParentDir"), rsArticleList("ClassDir"), rsArticleList("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsArticleList("UpdateTime"), rsArticleList("ArticleID")) & FileExt_Item
                End If
                If iClassPurview > 0 Or rsArticleList("InfoPoint") > 0 Then
                    Response.Write "<a href='#' title='因为设置了阅读权限，所以不用生成HTML'><font color=green><b>¤</b></font></a>"
                Else
                    If fso.FileExists(Server.MapPath(ArticlePath)) Then
                        Response.Write "<a href='#' title='文件位置：" & ArticlePath & "'><b>√</b></a>"
                    Else
                        Response.Write "<font color=red><b>×</b></font>"
                    End If
                End If
                Response.Write "</td>"
            End If
            Select Case ManageType
            Case "Check"
                Response.Write "<td width='120' align='center'>"
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
                    If rsArticleList("Status") <= MyStatus Then
                        If rsArticleList("Status") > -1 Then
                            Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Reject&ArticleID=" & rsArticleList("ArticleID") & "'>直接退稿</a>&nbsp;&nbsp;"
                        End If
                        If rsArticleList("Status") < MyStatus Then
                            Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Check&ArticleID=" & rsArticleList("ArticleID") & "'>审核</a>&nbsp;&nbsp;"
                            Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=SetPassed&ArticleID=" & rsArticleList("ArticleID") & "'>通过</a>&nbsp;&nbsp;"
                        Else
                            Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=CancelPassed&ArticleID=" & rsArticleList("ArticleID") & "'>取消审核</a>"
                        End If
                    End If
                End If
                Response.Write "</td>"
            Case "PayMoney"
                If rsArticleList("IsPayed") = False And PayStatus = "False" Then
                    Response.Write "<td width='60' align='center'><a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=ConfirmPay&ArticleID=" & rsArticleList("ArticleID") & "'>支付稿费</a></td>"
                End If
            Case "HTML"
                Response.Write "    <td width='180' align='left'>&nbsp;"
                If iClassPurview = 0 And rsArticleList("InfoPoint") = 0 And rsArticleList("Status") = 3 And (AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True) Then
                    Response.Write "<a href='Admin_CreateArticle.asp?ChannelID=" & ChannelID & "&Action=CreateArticle&ArticleID=" & rsArticleList("ArticleID") & "' title='生成本" & ChannelShortName & "的HTML页面'>生成文件</a>&nbsp;"
                    If fso.FileExists(Server.MapPath(ArticlePath)) Then
                        Response.Write "<a href='" & ArticlePath & "' target='_blank' title='查看本" & ChannelShortName & "的HTML页面'>查看文件</a>&nbsp;"
                        Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=DelFile&ArticleID=" & rsArticleList("ArticleID") & "' title='删除本" & ChannelShortName & "的HTML页面' onclick=""return confirm('确定要删除此" & ChannelShortName & "的HTML页面吗？');"">删除文件</a>&nbsp;"
                    End If
                End If
                Response.Write "</td>"
            Case "Recyclebin"
                Response.Write "<td width='100' align='center'>"
                Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=ConfirmDel&ArticleID=" & rsArticleList("ArticleID") & "' onclick=""return confirm('确定要彻底删除此" & ChannelShortName & "吗？彻底删除后将无法还原！');"">彻底删除</a> "
                Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Restore&ArticleID=" & rsArticleList("ArticleID") & "'>还原</a>"
                Response.Write "</td>"
            Case "Special"
                Response.Write "<td width='100' align='center'>"
                If rsArticleList("SpecialID") > 0 Then
                    Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=DelFromSpecial&InfoID=" & rsArticleList("InfoID") & "' onclick=""return confirm('确定要将此" & ChannelShortName & "从其所属专题中删除吗？');"">从所属专题中删除</a> "
                End If
                Response.Write "</td>"
            Case Else
                Response.Write "    <td width='150' align='left'>&nbsp;"
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or CheckPurview_Class(arrClass_Input, ParentPath & "," & ClassID) Or UserName = rsArticleList("Inputer") Then
                    Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Modify&ArticleID=" & rsArticleList("ArticleID") & "'>修改</a>&nbsp;"
                Else
                    Response.Write "<font color='gray'>修改&nbsp;</font>"
                End If
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Or UserName = rsArticleList("Inputer") Then
                    Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Del&ArticleID=" & rsArticleList("ArticleID") & "' onclick=""return confirm('确定要删除此" & ChannelShortName & "吗？删除后你还可以从回收站中还原。');"">删除</a>&nbsp;"
                Else
                    Response.Write "<font color='gray'>删除&nbsp;</font>"
                End If
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
                    If rsArticleList("OnTop") = False Then
                        Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=SetOnTop&ArticleID=" & rsArticleList("ArticleID") & "'>固顶</a>&nbsp;"
                    Else
                        Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=CancelOnTop&ArticleID=" & rsArticleList("ArticleID") & "'>解固</a>&nbsp;"
                    End If
                    If rsArticleList("Elite") = False Then
                        Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=SetElite&ArticleID=" & rsArticleList("ArticleID") & "'>设为推荐</a>"
                    Else
                        Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=CancelElite&ArticleID=" & rsArticleList("ArticleID") & "'>取消推荐</a>"
                    End If
                End If
                Response.Write "</td>"
            End Select
            Response.Write "</tr>"

            ArticleNum = ArticleNum + 1
            If ArticleNum >= MaxPerPage Then Exit Do
            rsArticleList.MoveNext
        Loop
    End If
    rsArticleList.Close
    Set rsArticleList = Nothing
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>选中本页显示的所有" & ChannelShortName & "</td><td>"
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
                    Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateClass';document.myform.CreateType.value='1';document.myform.action='Admin_CreateArticle.asp';"" value='生成当前栏目列表页'>&nbsp;&nbsp;"
                End If
                If ClassPurview = 0 Then
                    Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateArticle';document.myform.CreateType.value='2';document.myform.action='Admin_CreateArticle.asp';"" value='生成当前栏目的" & ChannelShortName & "'>&nbsp;&nbsp;"
                End If
            Else
                Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateIndex';document.myform.CreateType.value='1';document.myform.action='Admin_CreateArticle.asp';"" value='生成首页'>&nbsp;&nbsp;"
                If UseCreateHTML = 1 Or UseCreateHTML = 3 Then
                    Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateClass';document.myform.CreateType.value='2';document.myform.action='Admin_CreateArticle.asp';"" value='生成所有栏目列表页'>&nbsp;&nbsp;"
                End If
                Response.Write "<input name='submit1' type='submit' id='submit1' onClick=""document.myform.Action.value='CreateArticle';document.myform.CreateType.value='3';document.myform.action='Admin_CreateArticle.asp';"" value='生成所有" & ChannelShortName & "'>&nbsp;&nbsp;"
            End If
            Response.Write "<input name='submit2' type='submit' id='submit2' onClick=""document.myform.Action.value='CreateArticle';document.myform.action='Admin_CreateArticle.asp';"" value='生成选定的" & ChannelShortName & "'>&nbsp;&nbsp;&nbsp;&nbsp;"
            Response.Write "<input name='Submit3' type='submit' id='submit3' onClick=""document.myform.Action.value='DelFile';document.myform.action='Admin_Article.asp'"" value='删除选定" & ChannelShortName & "的HTML文件'>"
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
    Case "PayMoney"
        Response.Write "<Script Language='JavaScript'>"
        Response.Write "function SetBtStatPayValue()"
        Response.Write "{"
        Response.Write "document.myform.Action.value='ConfirmPay';"
        Response.Write "document.myform.submit();"
        Response.Write "}"
        Response.Write "function SetExportExcelValue()"
        Response.Write "{"
        Response.Write "document.myform.Action.value='ExportExcel';"
        Response.Write "document.myform.submit();"
        Response.Write "}"
        Response.Write "</Script>"
        Response.Write "<Input Type='Hidden' Name='ManageType' Value='" & ManageType & "'>"
        Response.Write "<Input Type='Hidden' Name='PayStatus' Value='" & PayStatus & "'>"
        Response.Write "<table border=0>"
        Response.Write "<tr>"
        If PayStatus = "False" Then
            Response.Write "<td><Input name='BtStatPay' type='Button' id='BtStatPay' value='批量支付稿费' onClick=""SetBtStatPayValue()""></td>"
        End If
        Response.Write "</tr>"
        Response.Write "<tr>"
        Response.Write "<td>"
        Call PopCalendarInit
        Response.Write "<input name='SelectType' type='radio' value='ID' >按ＩＤ范围选择："
        Response.Write "起始ＩＤ<input type='text' name='BeginID'  size='10' value='1'>"
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "终止ＩＤ<input type='text' name='EndID'  size='10' value='1000'>"
        Response.Write "&nbsp;&nbsp;&nbsp;<br>"
        Response.Write "<input name='SelectType' type='radio' value='Date'>按日期范围选择："
        Response.Write "起始日期<input type='text' name='BeginDate' id='BeginDate' size='10' maxlength='10'><a style='cursor:hand;' onClick='PopCalendar.show(document.myform.BeginDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>&nbsp;结束日期<input type='text' name='EndDate' id='EndDate' size='10' maxlength='10'><a style='cursor:hand;' onClick='PopCalendar.show(document.myform.EndDate, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>"
        If PayStatus = "False" Then
            Response.Write "<input type='button' name='btExportExcel'  value='导出未支付的" & ChannelShortName & "到EXCEL' onClick=""SetExportExcelValue()"">"
            
        Else
            Response.Write "<input type='submit' name='btExportExcel'  value='导出已支付的" & ChannelShortName & "到EXCEL' onClick=""SetExportExcelValue()"">"
        End If
        
        Response.Write "</td>"
        Response.Write "</tr>"
        Response.Write "</table>"
    Case Else
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Or PurviewChecked = True Then
            Response.Write "<input name='submit1' type='submit' value=' 批量删除 ' onClick=""document.myform.Action.value='Del'""> "
            If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
                Response.Write "<input type='submit' name='Submit4' value=' 批量移动 ' onClick=""document.myform.Action.value='MoveToClass'""> "
                Response.Write "<input type='submit' name='Submit3' value=' 批量设置 ' onClick=""document.myform.Action.value='Batch'""> "
                Response.Write "<input name='submit1' type='submit' value=' 审核通过 ' onClick=""document.myform.Action.value='SetPassed'""> "
                Response.Write "<input name='submit2' type='submit' value=' 取消审核 ' onClick=""document.myform.Action.value='CancelPassed'""> "
                Response.Write "<input name='submit3' type='submit' value=' 批量替换 ' onClick=""document.myform.Action.value='BatchReplace'""> "
            End If
        End If
    End Select
    
    Response.Write "<input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
'    If SystemDatabaseType = "SQL" Then
'        totalPut = Cmd.Parameters("RETURN_VALUE").Value
'        CurrentPage = Cmd.Parameters("@ActualPage").Value
'    End If
    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName & "", True)
    End If

    Response.Write "<form method='Get' name='SearchForm' action='" & FileName & "'>"
    Response.Write "<Input Type='Hidden' Name='PayStatus' Value='" & PayStatus & "'>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='80' align='right'><strong>" & ChannelShortName & "搜索：</strong></td>"
    Response.Write "   <td>"
    Response.Write "<select name='Field' size='1'>"
    Response.Write "<option value='Title' selected>" & ChannelShortName & "标题</option>"
    Response.Write "<option value='Content'>" & ChannelShortName & "内容</option>"
    Response.Write "<option value='Author'>" & ChannelShortName & "作者</option>"
    If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
        Response.Write "<option value='Inputer'>录入者</option>"
        Response.Write "<option value='Editor'>审核人</option>"
    End If
    Response.Write "<option value='UpdateTime'>更新时间</option>"
    Response.Write "<option value='Keyword'>关键字</option>"
    Response.Write "<option value='ID'>" & ChannelShortName & "ID</option>"
    Response.Write "</select>"
    Response.Write "<select name='ClassID'><option value=''>所有栏目</option>" & GetClass_Option(1, 0) & "</select>"
    Response.Write "<input type='text' name='keyword'  size='20' value='关键字' maxlength='50' onFocus='this.select();'>"
    Response.Write "<input type='submit' name='Submit'  value='搜索'>"
    Response.Write "<input name='ManageType' type='hidden' id='ManageType' value='" & ManageType & "'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "</td></tr></table></form>"
    Response.Write "<br><b>说明：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "属性中的各项含义：<font color=blue>顶</font>----固顶" & ChannelShortName & "，<font color=red>热</font>----热门" & ChannelShortName & "，<font color=green>荐</font>----推荐" & ChannelShortName & "，<font color=blue>图</font>----首页图片" & ChannelShortName & "，<font color=black>调</font>----有调查<br><br>"
End Sub

Sub ShowJS_Article()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function AddItem(strFileName){" & vbCrLf
    Response.Write "    var arrName=strFileName.split('.');" & vbCrLf
    Response.Write "    var FileExt=arrName[1];" & vbCrLf
    Response.Write "    if (FileExt=='gif'||FileExt=='jpg'||FileExt=='jpeg'||FileExt=='jpe'||FileExt=='bmp'||FileExt=='png'){" & vbCrLf
    Response.Write "        if (arrName[0].substr(arrName[0].length-2,arrName[0].length)!='_S'){" & vbCrLf
    Response.Write "            if(document.myform.IncludePic.selectedIndex<2){" & vbCrLf
    Response.Write "                document.myform.IncludePic.selectedIndex+=1;" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        frmPreview.img.src='" & InstallDir & ChannelDir & "/" & UploadDir & "/" & "'+strFileName;" & vbCrLf
    Response.Write "        document.myform.DefaultPicUrl.value=strFileName;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    document.myform.DefaultPicList.options[document.myform.DefaultPicList.length]=new Option(strFileName,strFileName);" & vbCrLf
    Response.Write "    document.myform.DefaultPicList.selectedIndex+=1;" & vbCrLf
    Response.Write "    if(document.myform.UploadFiles.value==''){" & vbCrLf
    Response.Write "        document.myform.UploadFiles.value=strFileName;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        document.myform.UploadFiles.value=document.myform.UploadFiles.value+'|'+strFileName;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function selectPaginationType(){" & vbCrLf
    Response.Write "  document.myform.PaginationType.value=2;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function rUseLinkUrl(){" & vbCrLf
    Response.Write "  if(document.myform.UseLinkUrl.checked==true){" & vbCrLf
    Response.Write "    document.myform.LinkUrl.disabled=false;" & vbCrLf
    Response.Write "     ArticleContent.style.display='none';" & vbCrLf
    Response.Write "     ArticleContent2.style.display='none';" & vbCrLf
    Response.Write "     ArticleContent3.style.display='none';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    document.myform.LinkUrl.disabled=true;" & vbCrLf
    Response.Write "    ArticleContent.style.display='';" & vbCrLf
    Response.Write "    ArticleContent2.style.display='';" & vbCrLf
    Response.Write "    ArticleContent3.style.display='';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  var CurrentMode=editor.CurrentMode;" & vbCrLf
    Response.Write "  if (CurrentMode==0){" & vbCrLf
    Response.Write "    document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else if(CurrentMode==1){" & vbCrLf
    Response.Write "    document.myform.Content.value=editor.HtmlEdit.document.body.innerText;" & vbCrLf
    Response.Write "  }" & vbCrLf

    Response.Write "  else{" & vbCrLf
    Response.Write "    alert('预览状态不能保存！请先回到编辑状态后再保存');" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf

    Response.Write "  if (document.myform.Title.value==''){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "标题不能为空！');" & vbCrLf
    Response.Write "    document.myform.Title.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Keyword.value==''){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    alert('关键字不能为空！');" & vbCrLf
    Response.Write "    document.myform.Keyword.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.UseLinkUrl.checked==true){" & vbCrLf
    Response.Write "    if (document.myform.LinkUrl.value==''||document.myform.LinkUrl.value=='http://'){" & vbCrLf
    Response.Write "      ShowTabs(0);" & vbCrLf
    Response.Write "      alert('请输入转向链接的地址！');" & vbCrLf
    Response.Write "      document.myform.LinkUrl.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    if (document.myform.Content.value==''){" & vbCrLf
    Response.Write "      ShowTabs(0);" & vbCrLf
    Response.Write "      alert('" & ChannelShortName & "内容不能为空！');" & vbCrLf
    Response.Write "      editor.HtmlEdit.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
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
    Response.Write "  return true;  " & vbCrLf
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
    
    Response.Write "function SelectUser(){" & vbCrLf
    Response.Write "    var arr=showModalDialog('Admin_SourceList.asp?TypeSelect=UserList&DefaultValue='+document.myform.InceptUser.value,'','dialogWidth:600px; dialogHeight:450px; help: no; scroll: yes; status: no');" & vbCrLf
    Response.Write "    if (arr != null){" & vbCrLf
    Response.Write "        document.myform.InceptUser.value=arr;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    
    Response.Write "function getPayMoney(){" & vbCrLf
    'Response.Write "alert(document.myform.PerWordMoney.value*document.myform.WordNumber.value/1000);"
    Response.Write "  document.myform.CopyMoney1.value=document.myform.PerWordMoney.value*document.myform.WordNumber.value/1000;" & vbCrLf
    Response.Write "}" & vbCrLf
    
    Response.Write "function IsDigit(){" & vbCrLf
    Response.Write "  return ((event.keyCode >= 48) && (event.keyCode <= 57));" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function CopyTitle(){" & vbCrLf
    Response.Write "  if (document.myform.VoteTitle.value==''){" & vbCrLf
    Response.Write "     document.myform.VoteTitle.value = document.myform.Title.value;" & vbCrLf
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
    'Response.Write "   ShowTabs(5);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==54) {" & vbCrLf
    'Response.Write "   ShowTabs(6);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==55) {" & vbCrLf
    'Response.Write "   ShowTabs(7);CopyTitle();" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write " if(window.event.keyCode==56) {" & vbCrLf
    'Response.Write "   ShowTabs(4);" & vbCrLf
    'Response.Write " }" & vbCrLf
    'Response.Write "}" & vbCrLf
    'Response.Write "document.onkeypress = getKey;" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Function ReplaceJSImage(ByVal Content)
    If Content = "" Then
        ReplaceJSImage = Content
        Exit Function
    End If
    Dim strTemp
    '图片替换JS
    regEx.Pattern = "(\<Script)(.[^\<]*)(\<\/Script\>)"
    Set Matches = regEx.Execute(Content)

    For Each Match In Matches
        strTemp = Replace(Match.value, "<", "[!")
        strTemp = Replace(strTemp, ">", "!]")
        strTemp = Replace(strTemp, "'", """")
        strTemp = "<IMG alt='#" & strTemp & "#' src=""" & InstallDir & "editor/images/jscript.gif"" border=0 $>"
        Content = Replace(Content, Match.value, strTemp)
    Next
    ReplaceJSImage = Content
End Function

Sub ShowTabs_Title()
    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>基本信息</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>所属专题</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>属性设置</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'"
    If FoundInArr(arrEnabledTabs, "Charge", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">收费选项</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(4)'"
    If FoundInArr(arrEnabledTabs, "Recieve", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">签收设置</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(5)'"
    If FoundInArr(arrEnabledTabs, "Copyfee", ",") = False Or Action = "Add" Then Response.Write " style='display:none'"
    Response.Write ">稿费设置</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(6);CopyTitle()'"
    If FoundInArr(arrEnabledTabs, "Vote", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">调查设置</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(7)'>自定义选项</td>" & vbCrLf
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
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(4)'"
    If FoundInArr(arrEnabledTabs, "Recieve", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">签收设置</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(5)'"
    If FoundInArr(arrEnabledTabs, "Copyfee", ",") = False Or Action = "Add" Then Response.Write " style='display:none'"
    Response.Write ">稿费设置</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(6);CopyTitle()'"
    If FoundInArr(arrEnabledTabs, "Vote", ",") = False Then Response.Write " style='display:none'"
    Response.Write ">调查设置</td>" & vbCrLf
    Response.Write "    <td id='TabBottom' class='title3' onclick='ShowTabs(7)'>自定义选项</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub

Sub Add()
    Call ShowJS_Article
    
    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_Article.asp?ChannelID=" & ChannelID & "'>" & ChannelName & "管理</a>&nbsp;&gt;&gt;&nbsp;添加" & ChannelShortName & "</td></tr></table>"
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Article.asp' target='_self'>"

    Call ShowTabs_Title

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf

    Call ShowTr_Class

    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "标题：</td>"
    Response.Write "            <td>"
    Response.Write "              <table width='100%' border='0' cellpadding='0' cellspacing='2'>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='64' class='tdbg5'>简短标题：</td>"
    Response.Write "                  <td>"
    
    Response.Write "                    <select name='IncludePic'>"
    Response.Write "                      <option  value='0' selected> </option>"
    Response.Write "                      <option value='1'>" & ArticlePro1 & "</option>"
    Response.Write "                      <option value='2'>" & ArticlePro2 & "</option>"
    Response.Write "                      <option value='3'>" & ArticlePro3 & "</option>"
    Response.Write "                      <option value='4'>" & ArticlePro4 & "</option>"
    Response.Write "                    </select>"
    
   
    Response.Write "                    <input name='Title' type='text' id='Title' value='' size='56' autocomplete='off' maxlength='255' class='bginput' onPropertyChange=""moreitem('Title',10," & ChannelID & ",'satitle');"" onBlur=""setTimeout('Element.hide(satitle)',500);"">"
    Response.Write "                    <select name='TitleFontColor' id='TitleFontColor'>"
    Response.Write "                      <option value='' selected>颜色</option>"
    Response.Write "                      <option value=''>默认</option>"
    Response.Write "                      <option value='#000000' style='background-color:#000000'></option>"
    Response.Write "                      <option value='#FFFFFF' style='background-color:#FFFFFF'></option>"
    Response.Write "                      <option value='#008000' style='background-color:#008000'></option>"
    Response.Write "                      <option value='#800000' style='background-color:#800000'></option>"
    Response.Write "                      <option value='#808000' style='background-color:#808000'></option>"
    Response.Write "                      <option value='#000080' style='background-color:#000080'></option>"
    Response.Write "                      <option value='#800080' style='background-color:#800080'></option>"
    Response.Write "                      <option value='#808080' style='background-color:#808080'></option>"
    Response.Write "                      <option value='#FFFF00' style='background-color:#FFFF00'></option>"
    Response.Write "                      <option value='#00FF00' style='background-color:#00FF00'></option>"
    Response.Write "                      <option value='#00FFFF' style='background-color:#00FFFF'></option>"
    Response.Write "                      <option value='#FF00FF' style='background-color:#FF00FF'></option>"
    Response.Write "                      <option value='#FF0000' style='background-color:#FF0000'></option>"
    Response.Write "                      <option value='#0000FF' style='background-color:#0000FF'></option>"
    Response.Write "                      <option value='#008080' style='background-color:#008080'></option>"
    Response.Write "                    </select>"
    Response.Write "                    <select name='TitleFontType' id='TitleFontType'>"
    Response.Write "                      <option value='0' selected>字形</option>"
    Response.Write "                      <option value='1'>粗体</option>"
    Response.Write "                      <option value='2'>斜体</option>"
    Response.Write "                      <option value='3'>粗+斜</option>"
    Response.Write "                      <option value='0'>规则</option>"
    Response.Write "                    </select>"
    Response.Write "                    <div id=""satitle"" style='display:none'></div>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='64' class='tdbg5'>完整标题：</td>"
    Response.Write "                  <td><input name='TitleIntact' type='text' id='TitleIntact' value='' size='80' maxlength='500'></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='64' class='tdbg5'>副 标 题：</td>"
    Response.Write "                  <td><input name='Subheading' type='text' id='Subheading' value='' size='80' maxlength='500'></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr>"
    Response.Write "                  <td></td><td><input name='ShowCommentLink' type='checkbox' id='ShowCommentLink' value='Yes'>显示" & ChannelShortName & "列表时在标题旁显示评论链接&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='checksame' value='检查重复标题' onclick=""showModalDialog('Admin_CheckSameTitle.asp?ModuleType=" & ModuleType & "&Title='+document.myform.Title.value,'checksame','dialogWidth:350px; dialogHeight:250px; help: no; scroll: no; status: no');""></td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>关键字：</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='Keyword' type='text' style=""clear:both"" id='Keyword' value='" & Trim(Session("Keyword")) & "' autocomplete='off' size='50' maxlength='255' onPropertyChange=""moreitem('Keyword',10," & ChannelID & ",'skey');"" onBlur=""setTimeout('Element.hide(skey)',500);""> <font color='#FF0000'>*</font> " & GetKeywordList("Admin", ChannelID)
    Response.Write "              </div><div id=""skey"" style='display:none'></div><font color='#0000FF'>用来查找相关" & ChannelShortName & "，可输入多个关键字，中间用<font color='#FF0000'>“|”</font>隔开。不能出现&quot;'&?;:()等字符。</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "作者：</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='Author' type='text' id='Author' value='" & Trim(Session("Author")) & "' autocomplete='off' size='50' maxlength='100' onPropertyChange=""moreitem('Author',10," & ChannelID & ",'sauthor');"" onBlur=""setTimeout('Element.hide(sauthor)',500);"">" & GetAuthorList("Admin", ChannelID, UserName)
    Response.Write "              </div><div id=""sauthor"" style='display:none'></div>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "来源：</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='CopyFrom' type='text' id='CopyFrom' value='" & Trim(Session("CopyFrom")) & "' autocomplete='off' size='50' maxlength='100' onPropertyChange=""moreitem('CopyFrom',10," & ChannelID & ",'scopyfrom');"" onBlur=""setTimeout('Element.hide(scopyfrom)',500);"">" & GetCopyFromList("Admin", ChannelID)
    Response.Write "              </div><div id=""scopyfrom"" style='display:none'></div>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><font color='#FF0000'>转向链接：</font></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='LinkUrl' type='text' id='LinkUrl' value='http://' size='50' maxlength='255' disabled>"
    Response.Write "              <input name='UseLinkUrl' type='checkbox' id='UseLinkUrl' value='Yes' onClick='rUseLinkUrl();'> <font color='#FF0000'>使用转向链接</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "简介：</td>"
    Response.Write "            <td><textarea name='Intro' cols='80' rows='4'></textarea></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg' id='ArticleContent' style=""display:''"">"
    Response.Write "            <td width='120' align='right' valign='bottom' class='tdbg5'><p>" & ChannelShortName & "内容：</p>"
    If EnableSaveRemote = True And IsObjInstalled("Microsoft.XMLHTTP") = True Then
        Response.Write "<table><tr><td><input type='checkbox' name='SaveRemotePic' value='Yes' checked></td><td>自动下载" & ChannelShortName & "内容里的图片</td>"
        If PhotoObject = 1 Then
            Response.Write "<tr><td><input type='checkbox' name='AddWatermark' value='Yes' checked></td><td>是否给" & ChannelShortName & "内容里远程获得的图片加水印</td></tr>"
            Response.Write "<tr><td><input type='checkbox' name='AddThumb' value='Yes' checked></td><td>是否给" & ChannelShortName & "内容里远程获得的第一张图片加缩略图</td></tr>"
        End If
        Response.Write "</tr></table>"
        Response.Write "<div align='left'><font color='#006600'>&nbsp;&nbsp;&nbsp;&nbsp;启用此功能后，如果从其它网站上复制内容到右边的编辑器中，并且内容中包含有图片，本系统会在保存" & ChannelShortName & "时自动把相关图片复制到本站服务器上。"
        Response.Write "<br>&nbsp;&nbsp;&nbsp;&nbsp;系统会因所下载图片的大小而影响速度，建议图片较多时不要使用此功能。</font>"
    End If
    Response.Write "<br><br><font color='red'>换行请按Shift+Enter<br><br>另起一段请按Enter</font></div><br><br><br><br><iframe id='frmPreview' width='120' height='150' frameborder='1' src='Admin_imgPreview.asp'></iframe>"
    Response.Write "            </td>"
    Response.Write "            <td>"
    Response.Write "              <textarea name='Content' style='display:none'>" & XmlText("Article", "DefaultAddTemplate", "") & "</textarea>"
    Response.Write "              <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=0&tContentid=Content' frameborder='1' scrolling='no' width='600' height='600' ></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><font color='#FF0000'>首页图片：</font></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='DefaultPicUrl' type='text' id='DefaultPicUrl' size='56' maxlength='200'>"
    Response.Write "              用于在首页的图片" & ChannelShortName & "处显示 <br>直接从上传图片中选择："
    Response.Write "              <select name='DefaultPicList' id='DefaultPicList' onChange=""DefaultPicUrl.value=this.value;frmPreview.img.src=((this.value == '') ? '../images/nopic.gif' : '" & InstallDir & ChannelDir & "/" & UploadDir & "/'+this.value);"">"
    Response.Write "                <option selected>不指定首页图片</option>"
    Response.Write "              </select>"
    Response.Write "              <input name='UploadFiles' type='hidden' id='UploadFiles'>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg' id='ArticleContent2' style=""display:''""> "
    Response.Write "            <td width='120' align='right' class='tdbg5'>内容分页方式：</td>"
    Response.Write "            <td>"
    Response.Write "              <select name='PaginationType' id='PaginationType'>"
    Response.Write "                <option value='0' selected>不分页</option>"
    Response.Write "                <option value='1'>自动分页</option>"
    Response.Write "                <option value='2'>手动分页</option>"
    Response.Write "              </select>"
    Response.Write "              &nbsp;&nbsp;&nbsp;&nbsp;<strong><font color='#0000FF'>注：</font></strong><font color='#0000FF'>手动分页符标记为“</font><font color='#FF0000'>[NextPage]</font><font color='#0000FF'>”，注意大小写</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg' id='ArticleContent3' style=""display:''"">"
    Response.Write "            <td width='120' align='right' class='tdbg5'>&nbsp;</td>"
    Response.Write "            <td>自动分页时的每页大约字符数（包含HTML标记且必须大于100）：<input name='MaxCharPerPage' type='text' id='MaxCharPerPage' value='10000' size='8' maxlength='8'></td>"
    Response.Write "          </tr>"
    
    Call ShowTabs_Status_Add
    
    Response.Write "        </tbody>" & vbCrLf
    
    
    Call ShowTabs_Special(SpecialID, "")

    Call ShowTabs_Property_Add
    
    Call ShowTabs_Purview_Add("阅读")
    
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>签收用户：</td>"
    Response.Write "            <td><textarea name='InceptUser' cols='72' rows='5' readonly></textarea><br>"
    Response.Write "                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='bt_choose' value='选择用户' onClick='SelectUser();'>"
    Response.Write "                    &nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='bt_cancel' value='清除用户' onClick=""myform.InceptUser.value=''"">"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>签收方式：</td>"
    Response.Write "            <td>"
    Response.Write "                    <select name='AutoReceiveTime'>"
    Response.Write "                      <option value='0' selected>手动签收</option>"
    Response.Write "                      <option value='5'>5秒钟后</option>"
    Response.Write "                      <option value='10'>10秒钟后</option>"
    Response.Write "                      <option value='30'>30秒钟后</option>"
    Response.Write "                      <option value='60'>1分钟后</option>"
    Response.Write "                      <option value='120'>2分钟后</option>"
    Response.Write "                      <option value='300'>5分钟后</option>"
    Response.Write "                    </select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>文档类型：</td>"
    Response.Write "            <td><select name='ReceiveType' id='ReceiveType'><option value='0' selected>公众文档</option><option value='1'>专属文档</option></select></td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf

    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
   
    Response.Write "        </tbody>" & vbCrLf

    Call ShowTabs_Vote_Add
    
    Call ShowTabs_MyField_Add
    
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"

    Call ShowTabs_Bottom
    
    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='add' type='submit'  id='Add' value=' 添 加 ' onClick=""document.myform.Action.value='SaveAdd';document.myform.target='_self';"" style='cursor:hand;'>&nbsp; "
    Response.Write "   <input name='Preview' type='submit'  id='Preview' value=' 预 览 ' onClick=""document.myform.Action.value='Preview';document.myform.target='_blank';"" style='cursor:hand;'>"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
End Sub


Sub Modify()
    Dim rsArticle, sql, tmpAuthor, tmpCopyFrom

    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的" & ChannelShortName & "ID</li>"
        Exit Sub
    Else
        ArticleID = PE_CLng(ArticleID)
    End If
    sql = "select * from PE_Article where ArticleID=" & ArticleID & ""
    Set rsArticle = Conn.Execute(sql)
    If rsArticle.BOF And rsArticle.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到" & ChannelShortName & "</li>"
        rsArticle.Close
        Set rsArticle = Nothing
        Exit Sub
    End If

    ClassID = rsArticle("ClassID")
    Dim arrSpecialID
    arrSpecialID = GetSpecialIDArr(ModuleType, ArticleID)

    If rsArticle("Inputer") <> UserName Then
        Call CheckClassPurview(Action, ClassID)
    End If
    If FoundErr = True Then
        rsArticle.Close
        Set rsArticle = Nothing
        Exit Sub
    End If
    tmpAuthor = rsArticle("Author")
    tmpCopyFrom = rsArticle("CopyFrom")
    EmailOfReject = Replace(EmailOfReject, "{$Title}", rsArticle("Title"))
    EmailOfPassed = Replace(EmailOfPassed, "{$Title}", rsArticle("Title"))

    Call ShowJS_Article


    Response.Write "<br><table width='100%'><tr><td align='left'>您现在的位置：<a href='Admin_Article.asp?ChannelID=" & ChannelID & "'>" & ChannelName & "管理</a>&nbsp;&gt;&gt;&nbsp;修改" & ChannelShortName & "</td></tr></table>"
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Article.asp' target='_self'>"

    Call ShowTabs_Title

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf

    Call ShowTr_Class

    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "标题：</td>"
    Response.Write "            <td>"
    Response.Write "              <table width='100%'  border='0' cellspacing='2' cellpadding='0'>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='64'>简短标题：</td>"
    Response.Write "                  <td>"

    Response.Write "                    <select name='IncludePic'>"
    Response.Write "                      <option  value='0'"
    If rsArticle("IncludePic") = 0 Then Response.Write " selected"
    Response.Write "> </option>"
    Response.Write "                      <option value='1'"
    If rsArticle("IncludePic") = 1 Then Response.Write " selected"
    Response.Write ">" & ArticlePro1 & "</option>"
    Response.Write "                      <option value='2'"
    If rsArticle("IncludePic") = 2 Then Response.Write " selected"
    Response.Write ">" & ArticlePro2 & "</option>"
    Response.Write "                      <option value='3'"
    If rsArticle("IncludePic") = 3 Then Response.Write " selected"
    Response.Write ">" & ArticlePro3 & "</option>"
    Response.Write "                      <option value='4'"
    If rsArticle("IncludePic") = 4 Then Response.Write " selected"
    Response.Write ">" & ArticlePro4 & "</option>"
    Response.Write "                    </select>"
        
    Response.Write "                    <input name='Title' type='text' id='Title' value='" & rsArticle("Title") & "' autocomplete='off' size='56' maxlength='255' class='bginput' onPropertyChange=""moreitem('Title',10," & ChannelID & ",'satitle');"" onBlur=""setTimeout('Element.hide(satitle)',500);"">"
    Response.Write "                    <select name='TitleFontColor' id='TitleFontColor'>"
    Response.Write "                      <option value=''"
    If rsArticle("TitleFontColor") = "" Then Response.Write " selected"
    Response.Write ">颜色</option>"
    Response.Write "                      <option value=''>默认</option>"
    Response.Write "                      <option value='#000000' style='background-color:#000000'"
    If rsArticle("TitleFontColor") = "#000000" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#FFFFFF' style='background-color:#FFFFFF'"
    If rsArticle("TitleFontColor") = "#FFFFFF" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#008000' style='background-color:#008000'"
    If rsArticle("TitleFontColor") = "#008000" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#800000' style='background-color:#800000'"
    If rsArticle("TitleFontColor") = "#800000" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#808000' style='background-color:#808000'"
    If rsArticle("TitleFontColor") = "#808000" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#000080' style='background-color:#000080'"
    If rsArticle("TitleFontColor") = "#000080" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#800080' style='background-color:#800080'"
    If rsArticle("TitleFontColor") = "#800080" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#808080' style='background-color:#808080'"
    If rsArticle("TitleFontColor") = "#808080" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#FFFF00' style='background-color:#FFFF00'"
    If rsArticle("TitleFontColor") = "#FFFF00" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#00FF00' style='background-color:#00FF00'"
    If rsArticle("TitleFontColor") = "#00FF00" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#00FFFF' style='background-color:#00FFFF'"
    If rsArticle("TitleFontColor") = "#00FFFF" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#FF00FF' style='background-color:#FF00FF'"
    If rsArticle("TitleFontColor") = "#FF00FF" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#FF0000' style='background-color:#FF0000'"
    If rsArticle("TitleFontColor") = "#FF0000" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#0000FF' style='background-color:#0000FF'"
    If rsArticle("TitleFontColor") = "#0000FF" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                      <option value='#008080' style='background-color:#008080'"
    If rsArticle("TitleFontColor") = "#008080" Then Response.Write " selected"
    Response.Write "></option>"
    Response.Write "                    </select>"
    Response.Write "                    <select name='TitleFontType' id='TitleFontType'>"
    Response.Write "                      <option value='0'"
    If rsArticle("TitleFontType") = 0 Then Response.Write " selected"
    Response.Write ">字形</option>"
    Response.Write "                      <option value='1'"
    If rsArticle("TitleFontType") = 1 Then Response.Write " selected"
    Response.Write ">粗体</option>"
    Response.Write "                      <option value='2'"
    If rsArticle("TitleFontType") = 2 Then Response.Write " selected"
    Response.Write ">斜体</option>"
    Response.Write "                      <option value='3'"
    If rsArticle("TitleFontType") = 3 Then Response.Write " selected"
    Response.Write ">粗+斜</option>"
    Response.Write "                      <option value='0'"
    If rsArticle("TitleFontType") = 4 Then Response.Write " selected"
    Response.Write ">规则</option>"
    Response.Write "                    </select>"
    Response.Write "                    <div id=""satitle"" style='display:none'></div>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='64'>完整标题：</td>"
    Response.Write "                  <td><input name='TitleIntact' type='text' id='TitleIntact' value='" & rsArticle("TitleIntact") & "' size='80' maxlength='500'></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='64'>副 标 题：</td>"
    Response.Write "                  <td><input name='Subheading' type='text' id='Subheading' value='" & rsArticle("Subheading") & "' size='80' maxlength='500'></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr>"
    Response.Write "                  <td width='64'>&nbsp;</td>"
    Response.Write "                  <td>"
    Response.Write "                    <input name='ShowCommentLink' type='checkbox' id='ShowCommentLink' value='Yes'"
    If rsArticle("ShowCommentLink") = True Then Response.Write "checked"
    Response.Write ">显示" & ChannelShortName & "列表时在标题旁显示评论链接</td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>关键字：</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='Keyword' type='text' id='Keyword' value='" & Mid(rsArticle("Keyword"), 2, Len(rsArticle("Keyword")) - 2) & "' autocomplete='off' size='50' maxlength='255' onPropertyChange=""moreitem('Keyword',10," & ChannelID & ",'skey');"" onBlur=""setTimeout('Element.hide(skey)',500);""> <font color='#FF0000'>*</font> " & GetKeywordList("Admin", ChannelID)
    Response.Write "              </div><div id=""skey"" style='display:none'></div>"
    Response.Write "              <font color='#0000FF'>用来查找相关" & ChannelShortName & "，可输入多个关键字，中间用<font color='#FF0000'>“|”</font>隔开。不能出现&quot;'&?;:()等字符。</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "作者：</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='Author' type='text' id='Author' value='" & tmpAuthor & "' autocomplete='off' size='50' maxlength='100' onPropertyChange=""moreitem('Author',10," & ChannelID & ",'sauthor');"" onBlur=""setTimeout('Element.hide(sauthor)',500);"">" & GetAuthorList("Admin", ChannelID, UserName)
    Response.Write "              </div><div id=""sauthor"" style='display:none'></div>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "来源：</td>"
    Response.Write "            <td>"
    Response.Write "              <div style=""clear: both;""><input name='CopyFrom' type='text' id='CopyFrom' value='" & tmpCopyFrom & "' autocomplete='off' size='50' maxlength='100' onPropertyChange=""moreitem('CopyFrom',10," & ChannelID & ",'scopyfrom');"" onBlur=""setTimeout('Element.hide(scopyfrom)',500);"">" & GetCopyFromList("Admin", ChannelID)
    Response.Write "              </div><div id=""scopyfrom"" style='display:none'></div>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><font color='#FF0000'>转向链接：</font></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='LinkUrl' type='text' id='LinkUrl' value='" & rsArticle("LinkUrl") & "' size='50' maxlength='255'"
    If rsArticle("LinkUrl") = "" Or rsArticle("LinkUrl") = "http://" Then Response.Write " disabled"
    Response.Write "> <input name='UseLinkUrl' type='checkbox' id='UseLinkUrl' value='Yes' onClick='rUseLinkUrl();'"
    If rsArticle("LinkUrl") <> "" And rsArticle("LinkUrl") <> "http://" Then Response.Write " checked"
    Response.Write "><font color='#FF0000'>使用转向链接</font></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "简介：</td>"
    Response.Write "            <td><textarea name='Intro' cols='80' rows='4'>" & rsArticle("Intro") & "</textarea></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg' id='ArticleContent' style=""display:'"
    If rsArticle("LinkUrl") <> "" And rsArticle("LinkUrl") <> "http://" Then Response.Write "none"
    Response.Write "'"">"
    Response.Write "            <td width='120' align='right' class='tdbg5'><p>" & ChannelShortName & "内容：</p>"
    If EnableSaveRemote = True And IsObjInstalled("Microsoft.XMLHTTP") = True Then
        Response.Write "<table><tr><td><input type='checkbox' name='SaveRemotePic' value='Yes' checked></td><td>自动下载" & ChannelShortName & "内容里的图片</td>"
        If PhotoObject = 1 Then
            Response.Write "<tr><td><input type='checkbox' name='AddWatermark' value='Yes' checked></td><td>是否给" & ChannelShortName & "内容里远程获得的图片加水印</td></tr>"
            Response.Write "<tr><td><input type='checkbox' name='AddThumb' value='Yes' checked></td><td>是否给" & ChannelShortName & "内容里远程获得的第一张图片加缩略图</td></tr>"
        End If
        Response.Write "</table>"
        Response.Write "<div align='left'><font color='#006600'>&nbsp;&nbsp;&nbsp;&nbsp;启用此功能后，如果从其它网站上复制内容到右边的编辑器中，并且内容中包含有图片，本系统会在保存" & ChannelShortName & "时自动把相关图片复制到本站服务器上。"
        Response.Write "<br>&nbsp;&nbsp;&nbsp;&nbsp;系统会因所下载图片的大小而影响速度，建议图片较多时不要使用此功能。</font>"
    End If
    Response.Write "<br><br><font color='red'>换行请按Shift+Enter<br><br>另起一段请按Enter</font></div><br><br><br><br><iframe id='frmPreview' width='120' height='150' frameborder='1' src='Admin_imgPreview.asp'></iframe>"
    Response.Write "            </td>"
    Response.Write "            <td>"
    Response.Write "              <textarea name='Content' style='display:none'>" & ReplaceJSImage(Replace(Replace(Server.HTMLEncode(FilterBadTag(rsArticle("Content"), rsArticle("Inputer"))), "[InstallDir_ChannelDir]", InstallDir & ChannelDir & "/"), "{$UploadDir}", UploadDir)) & "</textarea>"
    Response.Write "              <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=0&tContentid=Content' frameborder='1' scrolling='no' width='600' height='600' ></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>首页图片：</td>"
    Response.Write "            <td><input name='DefaultPicUrl' type='text' id='DefaultPicUrl' value='" & rsArticle("DefaultPicUrl") & "' size='56' maxlength='200'>"
    Response.Write "              用于在首页的图片" & ChannelShortName & "处显示 <br>直接从上传图片中选择："
    Response.Write "              <select name='DefaultPicList' id='DefaultPicList' onChange=""DefaultPicUrl.value=this.value;frmPreview.img.src=((this.value == '') ? '../images/nopic.gif' : '" & InstallDir & ChannelDir & "/" & UploadDir & "/'+this.value);"">"
    Response.Write "                <option value=''"
    If rsArticle("DefaultPicUrl") = "" Then Response.Write "selected"
    Response.Write ">不指定首页图片</option>"
    If rsArticle("UploadFiles") <> "" Then
        Dim IsOtherUrl
        IsOtherUrl = True
        If InStr(rsArticle("UploadFiles"), "|") > 1 Then
            Dim arrUploadFiles, intTemp
            arrUploadFiles = Split(rsArticle("UploadFiles"), "|")
            For intTemp = 0 To UBound(arrUploadFiles)
                If rsArticle("DefaultPicUrl") = arrUploadFiles(intTemp) Then
                    Response.Write "<option value='" & arrUploadFiles(intTemp) & "' selected>" & arrUploadFiles(intTemp) & "</option>"
                    IsOtherUrl = False
                Else
                    Response.Write "<option value='" & arrUploadFiles(intTemp) & "'>" & arrUploadFiles(intTemp) & "</option>"
                End If
            Next
        Else
            If rsArticle("UploadFiles") = rsArticle("DefaultPicUrl") Then
                Response.Write "<option value='" & rsArticle("UploadFiles") & "' selected>" & rsArticle("UploadFiles") & "</option>"
                IsOtherUrl = False
            Else
                Response.Write "<option value='" & rsArticle("UploadFiles") & "'>" & rsArticle("UploadFiles") & "</option>"
            End If
        End If
        If IsOtherUrl = True And rsArticle("DefaultPicUrl") <> "" Then
            Response.Write "<option value='" & rsArticle("DefaultPicUrl") & "' selected>" & rsArticle("DefaultPicUrl") & "</option>"
        End If
    End If
    Response.Write "              </select>"
    Response.Write "              <input name='UploadFiles' type='hidden' id='UploadFiles' value='" & rsArticle("UploadFiles") & "'> "
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg' id='ArticleContent2' style=""display:'"
    If rsArticle("LinkUrl") <> "" And rsArticle("LinkUrl") <> "http://" Then Response.Write "none"
    Response.Write "'"">"
    Response.Write "            <td width='120' align='right' class='tdbg5'>内容分页方式：</td>"
    Response.Write "            <td>"
    Response.Write "              <select name='PaginationType' id='PaginationType'>"
    Response.Write "                <option value='0'"
    If rsArticle("PaginationType") = 0 Then Response.Write " selected"
    Response.Write ">不分页</option>"
    Response.Write "                <option value='1'"
    If rsArticle("PaginationType") = 1 Then Response.Write " selected"
    Response.Write ">自动分页</option>"
    Response.Write "                <option value='2'"
    If rsArticle("PaginationType") = 2 Then Response.Write " selected"
    Response.Write ">手动分页</option>"
    Response.Write "              </select>"
    Response.Write "              &nbsp;&nbsp;&nbsp;&nbsp;<strong><font color='#0000FF'>注：</font></strong><font color='#0000FF'>手动分页符标记为“</font><font color='#FF0000'>[NextPage]</font><font color='#0000FF'>”，注意大小写</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg' id='ArticleContent3' style=""display:'"
    If rsArticle("LinkUrl") <> "" And rsArticle("LinkUrl") <> "http://" Then Response.Write "none"
    Response.Write "'"">"
    Response.Write "            <td width='120' align='right' class='tdbg5'>&nbsp;</td>"
    Response.Write "            <td>自动分页时的每页大约字符数（包含HTML标记且必须大于100）：<input name='MaxCharPerPage' type='text' id='MaxCharPerPage' value='" & rsArticle("MaxCharPerPage") & "' size='8' maxlength='8'></td>"
    Response.Write "          </tr>"
    Call ShowTabs_Status_Modify(rsArticle)
    Response.Write "        </tbody>" & vbCrLf

    Call ShowTabs_Special(arrSpecialID, "")
    
    Call ShowTabs_Property_Modify(rsArticle)
    
    Call ShowTabs_Purview_Modify("阅读", rsArticle, "")
    
    
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>签收用户：</td>"
    Response.Write "            <td><textarea name='InceptUser' cols='72' rows='3' readonly>" & rsArticle("ReceiveUser") & "</textarea><br>"
    Response.Write "                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='bt_choose' value='选择用户' onClick='SelectUser();'>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='bt_cancel' value='清除用户' onClick=""myform.InceptUser.value=''""></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>签收方式：</td>"
    Response.Write "            <td><select name='AutoReceiveTime'>"
    Response.Write "                      <option value='0'"
    If rsArticle("AutoReceiveTime") = "0" Then Response.Write " selected"
    Response.Write ">手动签收</option>"
    Response.Write "                      <option value='5'"
    If rsArticle("AutoReceiveTime") = "5" Then Response.Write " selected"
    Response.Write ">5秒钟后</option>"
    Response.Write "                      <option value='10'"
    If rsArticle("AutoReceiveTime") = "10" Then Response.Write " selected"
    Response.Write ">10秒钟后</option>"
    Response.Write "                      <option value='30'"
    If rsArticle("AutoReceiveTime") = "30" Then Response.Write " selected"
    Response.Write ">30秒钟后</option>"
    Response.Write "                      <option value='60'"
    If rsArticle("AutoReceiveTime") = "60" Then Response.Write " selected"
    Response.Write ">1分钟后</option>"
    Response.Write "                      <option value='120'"
    If rsArticle("AutoReceiveTime") = "120" Then Response.Write " selected"
    Response.Write ">2分钟后</option>"
    Response.Write "                      <option value='300'"
    If rsArticle("AutoReceiveTime") = "300" Then Response.Write " selected"
    Response.Write ">5分钟后</option>"
    Response.Write "                    </select>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>文档类型：</td>"
    Response.Write "            <td><select name='ReceiveType' id='ReceiveType'>"
    Response.Write "                      <option value='0'"
    If rsArticle("ReceiveType") = "0" Then Response.Write " selected"
    Response.Write ">公众文档</option>"
    Response.Write "                      <OPTION value='1'"
    If rsArticle("ReceiveType") = "1" Then Response.Write " selected"
    Response.Write ">专属文档</OPTION>"
    Response.Write "                    </select>"
    Response.Write "            </td>"
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Dim WordNum, strArticle, strSql, rsChannel, payNum
    '过滤空格，HTML字符，计算字数
    WordNum = getWordNumber(rsArticle("Content"))
    
    If MoneyPerKw <= 0 Then
       payNum = 0
    Else
       payNum = MoneyPerKw / 1000 * WordNum
    End If
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>文章字数：</td><td><input type='text' name='WordNumber' MaxLength='10' size=6 disabled value='" & WordNum & "'>&nbsp;字</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>支付标准：</td><td><input type='text' name='PerWordMoney' MaxLength='10'size=6 value=' " & MoneyPerKw & "'ONKEYPRESS=""event.returnValue=IsDigit();"">&nbsp;元/千字</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>估计稿费：</td><td><input type ='text' name='CopyMoney1' MaxLength='10' size='6' value='" & payNum & "' disabled> 元&nbsp;&nbsp;&nbsp;&nbsp;<input type=button name=payCalculate value=' 计算 ' onclick='getPayMoney();'></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>确认支付稿费：</td><td><input type ='text' name='CopyMoney' ONKEYPRESS=""event.returnValue=IsDigit();"" MaxLength='10' size='6'"
    If rsArticle("IsPayed") = True Then
        Response.Write " disabled"
    End If
    Response.Write " value='" & rsArticle("CopyMoney") & "'> 元</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>稿费受益者：</td><td><input type='text' name='Beneficiary' size='20' value='" & rsArticle("Inputer") & "'>&nbsp;&nbsp;<font color=blue>多个受益者之间用"",""隔开</font></td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf
    
    Call ShowTabs_Vote_Modify(rsArticle)

    Call ShowTabs_MyField_Modify(rsArticle)
        
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>" & vbCrLf

    Call ShowTabs_Bottom

    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "   <input name='ArticleID' type='hidden' id='ArticleID' value='" & rsArticle("ArticleID") & "'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='Save' type='submit' value='保存修改结果' onClick=""document.myform.Action.value='SaveModify';document.myform.target='_self';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Save' type='submit' value='添加为新" & ChannelShortName & "' onClick=""document.myform.Action.value='SaveModifyAsAdd';document.myform.target='_self';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Preview' type='submit'  id='Preview' value=' 预 览 ' onClick=""document.myform.Action.value='Preview';document.myform.target='_blank';"" style='cursor:hand;'>"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'><br>"
    Response.Write "  </p><br>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    Response.Write "<script language='javascript'>setTimeout('setpic()',1000);" & vbCrLf
    Response.Write "function setpic(){" & vbCrLf
    If rsArticle("DefaultPicUrl") <> "" Then
        If Left(rsArticle("DefaultPicUrl"), 1) <> "/" And InStr(rsArticle("DefaultPicUrl"), "://") <= 0 Then
            Response.Write "frmPreview.img.src='" & InstallDir & ChannelDir & "/" & UploadDir & "/" & rsArticle("DefaultPicUrl") & "';"
        Else
            Response.Write "frmPreview.img.src='" & rsArticle("DefaultPicUrl") & "';"
        End If
    End If
    Response.Write "}" & vbCrLf
    Response.Write "</script>"
    
    rsArticle.Close
    Set rsArticle = Nothing


End Sub

Sub SaveArticle()
    Dim rsArticle, sql, trs, i
    Dim ArticleID, ClassID, SpecialID, Title, Content
    Dim Keyword, Author, tAuthor, CopyFrom, Inputer, Editor, UpdateTime
    Dim arrUploadFiles, LinkUrl, UseLinkUrl
    Dim ReceiveUser
    Dim arrSpecialID

    ArticleID = Trim(Request.Form("ArticleID"))
    ClassID = Trim(Request.Form("ClassID"))
    SpecialID = Trim(Request.Form("SpecialID"))

    Title = Trim(Request.Form("Title"))
    Keyword = Trim(Request.Form("Keyword"))
    UseLinkUrl = Trim(Request.Form("UseLinkUrl"))
    LinkUrl = Trim(Request.Form("LinkUrl"))
    For i = 1 To Request.Form("Content").Count
        Content = Content & Request.Form("Content")(i)
    Next
    Author = Trim(Request.Form("Author"))
    CopyFrom = Trim(Request.Form("CopyFrom"))
    DefaultPicUrl = Trim(Request.Form("DefaultPicUrl"))
    UploadFiles = Trim(Request.Form("UploadFiles"))
    UpdateTime = PE_CDate(Trim(Request.Form("UpdateTime")))

    '注意检验这里的值
    Status = PE_CLng(Trim(Request.Form("Status")))
    ReceiveUser = ReplaceBadChar(Trim(Request("InceptUser")))
    IncludePic = PE_CLng(Trim(Request.Form("IncludePic")))
    
    Inputer = UserName
    Editor = AdminName


    Call CheckClassPurview(Action, ClassID)
    If FoundErr = True Then Exit Sub
    
    If Title = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "标题不能为空</li>"
    Else
        Title = ReplaceText(Title, 2)
    End If
    Keyword = ReplaceBadChar(Keyword)
    If Keyword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入" & ChannelShortName & "关键字</li>"
    Else
        Call SaveKeyword(Keyword)
    End If
    If UseLinkUrl = "Yes" Then
        If LinkUrl = "" Or LCase(LinkUrl) = "http://" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>链接地址不能为空</li>"
        Else
            If InStr(LinkUrl, "://") <= 0 And Left(LinkUrl, 1) <> "/" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>本站地址请以 / 开头。</li>"
            End If
        End If
    Else
        If Content = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & ChannelShortName & "内容不能为空</li>"
        End If
    End If
    
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-1 or ChannelID=" & ChannelID & "")
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
    
    '处理图片JS标签代码
    Dim strTemp, strTemp2
    regEx.Pattern = "\<IMG(.[^\<]*)\$\>"
    Set Matches = regEx.Execute(Content)

    For Each Match In Matches
        regEx.Pattern = "\#(.*)\#"
        Set strTemp = regEx.Execute(Match.value)

        For Each Match2 In strTemp
            strTemp2 = Replace(Match2.value, "&amp;", "&")
            strTemp2 = Replace(strTemp2, "#", "")
            strTemp2 = Replace(strTemp2, "&13;&10;", vbCrLf)
            strTemp2 = Replace(strTemp2, "&9;", "vbTab")
            strTemp2 = Replace(strTemp2, "[!", "<")
            strTemp2 = Replace(strTemp2, "!]", ">")
            Content = Replace(Content, Match.value, strTemp2)
        Next
    Next

    Title = PE_HTMLEncode(Title)
    Keyword = Replace("|" & Keyword & "|","||","|")

    '将绝对地址转化为相对地址
    Dim strSiteUrl
    strSiteUrl = Request.ServerVariables("HTTP_REFERER")
    strSiteUrl = LCase(Left(strSiteUrl, InStrRev(strSiteUrl, "/") - 1))
    strSiteUrl = Left(strSiteUrl, InStrRev(strSiteUrl, "/")) & ChannelDir & "/"
    Content = Replace(Content, strSiteUrl, "[InstallDir_ChannelDir]")
    Content = Replace(Content, UploadDir, "{$UploadDir}")

    If Trim(Request.Form("SaveRemotePic")) = "Yes" And EnableSaveRemote = True Then
        Content = ReplaceRemoteUrl(Content)
    End If

    '以下这段代码是为了解决内容中有如下情况下的频道目录的替换问题。就是说，只替换以频道目录开头的地址，如果是外部地址中含有频道目录，就不替换
    '<a href="/aaa/999.rar">
    '<a href="http://www.baidu.com/aaa/999.rar">
    '<img src="/aaa/999.rar">
    '<img src=/aaa/999.rar>
    '<img src='/aaa/999.rar>

    strSiteUrl = InstallDir & ChannelDir & "/"
    Content = Replace(Content, "'" & strSiteUrl, "'" & "[InstallDir_ChannelDir]")
    Content = Replace(Content, """" & strSiteUrl, """" & "[InstallDir_ChannelDir]")
    Content = Replace(Content, "=" & strSiteUrl, "=" & "[InstallDir_ChannelDir]")

    
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

    Set rsArticle = Server.CreateObject("adodb.recordset")
    If Action = "SaveAdd" Or Action = "SaveModifyAsAdd" Then
        If Session("Title") = Title And DateDiff("S", Session("AddTime"), Now()) < 100 Then
            FoundErr = True
            ErrMsg = "<li>请不要重复添加同一" & ChannelItemUnit & ChannelShortName & "</li>"
            Exit Sub
        Else
            Session("Title") = Title
            Session("AddTime") = Now()
            ArticleID = GetNewID("PE_Article", "ArticleID")
            
            For i = 0 To UBound(arrSpecialID)
                Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (" & ModuleType & "," & ArticleID & "," & PE_CLng(arrSpecialID(i)) & ")")
            Next
            sql = "select top 1 * from PE_Article"
            rsArticle.Open sql, Conn, 1, 3
            rsArticle.addnew
            rsArticle("ArticleID") = ArticleID
            rsArticle("ChannelID") = ChannelID
            rsArticle("Inputer") = Inputer

            If UserID <> "" And UserID > 0 Then
                Dim blogid
                Set blogid = Conn.Execute("select top 1 ID from PE_Space where Type=1 and UserID=" & UserID)
                If blogid.BOF And blogid.EOF Then
                    rsArticle("BlogID") = 0
                Else
                    rsArticle("BlogID") = blogid("ID")
                End If
                Set blogid = Nothing
            End If
            
            If ReceiveUser <> "" And Status = 3 Then
                rsArticle("Receive") = True
                Call Add_User_UnsignedItems(ArticleID, ReceiveUser)
            Else
                rsArticle("Receive") = False
            End If

            If Status = 3 Then
                Conn.Execute ("update PE_User set UserExp=UserExp+" & PresentExp & " where UserName='" & Inputer & "'")
            End If
        End If
        
    ElseIf Action = "SaveModify" Then
        If ArticleID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>不能确定ArticleID的值</li>"
        Else
            ArticleID = PE_CLng(ArticleID)
            sql = "select * from PE_Article where ArticleID=" & ArticleID
            rsArticle.Open sql, Conn, 1, 3
            If rsArticle.BOF And rsArticle.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>找不到此" & ChannelShortName & "，可能已经被其他人删除。</li>"
            Else
            
                '删除生成的文件，因为生成的文件可能会随着更新时间，游览权限等发生变化而产生多余的文件
                If UseCreateHTML > 0 And ObjInstalled_FSO = True Then
                    Dim tClass, ArticlePath
                    Set tClass = Conn.Execute("select ParentDir,ClassDir from PE_Class where ClassID=" & rsArticle("ClassID") & "")
                    If tClass.BOF And tClass.EOF Then
                        ParentDir = "/"
                        ClassDir = ""
                    Else
                        ParentDir = tClass("ParentDir")
                        ClassDir = tClass("ClassDir")
                    End If
                    ArticlePath = HtmlDir & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsArticle("UpdateTime"), rsArticle("ArticleID"))
                    If fso.FileExists(Server.MapPath(ArticlePath & FileExt_Item)) Then
                        DelSerialFiles Server.MapPath(ArticlePath & FileExt_Item)
                    End If
                    If rsArticle("PaginationType") > 0 Then
                        DelSerialFiles Server.MapPath(ArticlePath) & "_*.*"
                    End If
                End If
                If rsArticle("Inputer") <> UserName And rsArticle("Status") <> Status And (Status = -2 Or Status = 3) Then
                    Call SendEmailOfCheck(rsArticle("Inputer"), rsArticle)
                End If

                Call UpdateUserData(0, rsArticle("Inputer"), 0, 0)
            
                If rsArticle("Status") < 3 And Status = 3 Then
                    Conn.Execute ("update PE_User set UserExp=UserExp+" & rsArticle("PresentExp") & " where UserName='" & rsArticle("Inputer") & "'")
                End If
                If rsArticle("Status") = 3 And Status < 3 Then
                    Conn.Execute ("update PE_User set UserExp=UserExp-" & rsArticle("PresentExp") & " where UserName='" & rsArticle("Inputer") & "'")
                End If
    
                Dim rsInfo, sqlInfo, j
                i = 0
                sqlInfo = "select * from PE_InfoS where ModuleType=" & ModuleType & " and ItemID=" & ArticleID & " order by InfoID"
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
                            Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (" & ModuleType & "," & ArticleID & "," & PE_CLng(arrSpecialID(j)) & ")")
                        End If
                    Next
                End If
                
                If ReceiveUser = "" Or Status <> 3 Then
                    rsArticle("Receive") = False
                    If rsArticle("ReceiveUser") <> "" Then
                        Call Del_User_UnsignedItems(ArticleID, rsArticle("ReceiveUser"))
                    End If
                Else
                    rsArticle("Receive") = True
                    If rsArticle("ReceiveUser") <> "" Then
                        Call Del_User_UnsignedItems(ArticleID, rsArticle("ReceiveUser"))
                    End If
                    Call Add_User_UnsignedItems(ArticleID, ReceiveUser)
                End If
                        
            End If
        End If
    End If

    rsArticle("ClassID") = ClassID
    rsArticle("Title") = Title
    rsArticle("TitleIntact") = Trim(Request.Form("TitleIntact"))
    rsArticle("Subheading") = Trim(Request.Form("Subheading"))
    rsArticle("TitleFontColor") = Trim(Request.Form("TitleFontColor"))
    rsArticle("TitleFontType") = PE_CLng(Trim(Request.Form("TitleFontType")))
    rsArticle("Intro") = Trim(Request.Form("Intro"))
    rsArticle("Content") = Content
    rsArticle("Keyword") = Keyword
    rsArticle("Author") = Author
    rsArticle("CopyFrom") = CopyFrom
    rsArticle("LinkUrl") = LinkUrl
    rsArticle("Editor") = Editor
    rsArticle("IncludePic") = IncludePic
    rsArticle("ShowCommentLink") = PE_CBool(Trim(Request.Form("ShowCommentLink")))
    rsArticle("Status") = Status
    rsArticle("OnTop") = PE_CBool(Trim(Request.Form("OnTop")))
    rsArticle("Elite") = PE_CBool(Trim(Request.Form("Elite")))
    If Action = "SaveModifyAsAdd" Then 
        rsArticle("Hits") = 0
	Else 
	    rsArticle("Hits") = PE_CLng(Trim(Request.Form("Hits")))
    End IF	
    rsArticle("Stars") = PE_CLng(Trim(Request.Form("Stars")))
    rsArticle("UpdateTime") = UpdateTime
    rsArticle("CreateTime") = UpdateTime
    rsArticle("PaginationType") = PE_CLng(Trim(Request.Form("PaginationType")))
    rsArticle("MaxCharPerPage") = PE_CLng(Trim(Request.Form("MaxCharPerPage")))
    rsArticle("SkinID") = PE_CLng(Trim(Request.Form("SkinID")))
    rsArticle("TemplateID") = PE_CLng(Trim(Request.Form("TemplateID")))
    rsArticle("DefaultPicUrl") = DefaultPicUrl
    rsArticle("UploadFiles") = UploadFiles
    rsArticle("Deleted") = False
    rsArticle("PresentExp") = PresentExp

    rsArticle("Copymoney") = PE_CDbl(Trim(Request.Form("CopyMoney"))) '稿费
    rsArticle("IsPayed") = False
    rsArticle("Beneficiary") = Trim(Request.Form("Beneficiary"))    '稿费受益者 多个受益者之间用“，”隔开

    rsArticle("InfoPurview") = PE_CLng(Trim(Request.Form("InfoPurview")))
    rsArticle("arrGroupID") = ReplaceBadChar(Trim(Request.Form("GroupID")))
    rsArticle("InfoPoint") = PE_CLng(Trim(Request.Form("InfoPoint")))
    rsArticle("ChargeType") = PE_CLng(Trim(Request.Form("ChargeType")))
    rsArticle("PitchTime") = PE_CLng(Trim(Request.Form("PitchTime")))
    rsArticle("ReadTimes") = PE_CLng(Trim(Request.Form("ReadTimes")))
    rsArticle("DividePercent") = PE_CLng(Trim(Request.Form("DividePercent")))

    rsArticle("ReceiveUser") = ReceiveUser
    rsArticle("Received") = ""
    rsArticle("AutoReceiveTime") = PE_CLng(Request("AutoReceiveTime"))
    rsArticle("ReceiveType") = PE_CLng(Request("ReceiveType")) '获得签收文章的类型，0 为公有，1为私有
    rsArticle("VoteID") = VoteID

    If Not (rsField.BOF And rsField.EOF) Then
        rsField.MoveFirst
        Do While Not rsField.EOF
            If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                rsArticle(Trim(rsField("FieldName"))) = Trim(Request(rsField("FieldName")))
            End If
            rsField.MoveNext
        Loop
    End If
    Set rsField = Nothing

    rsArticle.Update
    rsArticle.Close
    Set rsArticle = Nothing
    Call UpdateChannelData(ChannelID)
    If Action = "SaveAdd" Or Action = "SaveModifyAsAdd" Then
        Call UpdateUserData(0, Inputer, 0, 0)
    End If
    Response.Write "<br><br>"
    Response.Write "<table class='border' align='center' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "  <tr class='title'> "
    Response.Write "    <td  height='22' align='center' colspan='2'> "
    If Action = "SaveAdd" Or Action = "SaveModifyAsAdd" Then
        Response.Write "<b>添加" & ChannelShortName & "成功</b>"
    Else
        Response.Write "<b>修改" & ChannelShortName & "成功</b>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right' class='tdbg5'><strong>所属栏目：</strong></td>"
    Response.Write "          <td width='400'>" & ShowClassPath() & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right' class='tdbg5'><strong>" & ChannelShortName & "标题：</strong></td>"
    Response.Write "          <td width='400'>" & Title & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right' class='tdbg5'><strong>作&nbsp;&nbsp;&nbsp;&nbsp;者：</strong></td>"
    Response.Write "          <td width='400'>" & Author & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right' class='tdbg5'><strong>" & ChannelShortName & "来源：</strong></td>"
    Response.Write "          <td width='400'>" & CopyFrom & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right' class='tdbg5'><strong>关 键 字：</strong></td>"
    Response.Write "          <td width='400'>" & Mid(Keyword, 2, Len(Keyword) - 2) & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right' class='tdbg5'><strong>" & ChannelShortName & "状态：</strong></td>"
    Response.Write "          <td width='400'>"
    If Status = -1 Then
        Response.Write "草稿"
    ElseIf Status = -2 Then
        Response.Write "退稿"
    Else
        Response.Write arrStatus(Status)
    End If
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg' align='center'>"
    Response.Write "    <td height='30' colspan='2'>"
    Response.Write "【<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Modify&ArticleID=" & ArticleID & "'>修改本文</a>】&nbsp;"
    Response.Write "【<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>继续添加" & ChannelShortName & "</a>】&nbsp;"
    Response.Write "【<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Manage&ClassID=" & ClassID & "'>" & ChannelShortName & "管理</a>】&nbsp;"
    Response.Write "【<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & ArticleID & "'>查看" & ChannelShortName & "内容</a>】"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf

    Session("Keyword") = Trim(Request("Keyword"))
    Session("Author") = Author
    Session("CopyFrom") = CopyFrom
    Session("PaginationType") = PE_CLng(Trim(Request("PaginationType")))
    Session("SkinID") = PE_CLng(Trim(Request("SkinID")))
    Session("TemplateID") = PE_CLng(Trim(Request("TemplateID")))
    
    Call ClearSiteCache(0)
    Call CreateAllJS

    If Status = 3 And PE_CLng(Trim(Request("ReceiveType"))) = 0 And UseCreateHTML > 0 And ObjInstalled_FSO = True And Trim(Request.Form("CreateImmediate")) = "Yes" Then
        Response.Write "<br><iframe id='CreateArticle' width='100%' height='210' frameborder='0' src='Admin_CreateArticle.asp?ChannelID=" & ChannelID & "&Action=CreateArticle2&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&ArticleID=" & ArticleID & "&ShowBack=No'></iframe>"
    End If
End Sub

Sub Show()
    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定" & ChannelShortName & "ID！</li>"
        Exit Sub
    End If
    
    Dim rsArticle, PurviewChecked, PurviewChecked2
    PurviewChecked = False
    PurviewChecked2 = False
    Set rsArticle = Conn.Execute("select * from PE_Article where ArticleID=" & ArticleID & "")
    If rsArticle.BOF And rsArticle.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到" & ChannelShortName & "！</li>"
        rsArticle.Close
        Set rsArticle = Nothing
        Exit Sub
    End If
    ClassID = rsArticle("ClassID")

    Call CheckClassPurview(Action, ClassID)
    If FoundErr = True Then
        rsArticle.Close
        Set rsArticle = Nothing
        Exit Sub
    End If

    Dim arrSpecialID
    arrSpecialID = GetSpecialIDArr(ModuleType, ArticleID)

    Call WriteJS_Show

    Response.Write "<br>您现在的位置：&nbsp;<a href='Admin_Article.asp?ChannelID=" & ChannelID & "'>" & ChannelShortName & "管理</a>&nbsp;&gt;&gt;&nbsp;"
    If ParentID > 0 Then
        Dim sqlPath, rsPath
        sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
        Set rsPath = Conn.Execute(sqlPath)
        Do While Not rsPath.EOF
            Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    Response.Write "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;查看" & ChannelShortName & "内容："
    Select Case rsArticle("IncludePic")
        Case 1
            Response.Write "<font color=blue>" & ArticlePro1 & "</font>"
        Case 2
            Response.Write "<font color=blue>" & ArticlePro2 & "</font>"
        Case 3
            Response.Write "<font color=blue>" & ArticlePro3 & "</font>"
        Case 4
            Response.Write "<font color=blue>" & ArticlePro4 & "</font>"
    End Select
    
    Response.Write rsArticle("Title") & "<br><br>"



    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>文章信息</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>所属专题</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>收费信息</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>签收信息</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(4)'>稿费信息</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(5)'>自定义选项</td>" & vbCrLf
    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf
    
    Response.Write "  <tr align='center' class='tdbg'>"
    Response.Write "    <td height='40' colspan='2'>"
    If Trim(rsArticle("TitleIntact")) <> "" Then
        Response.Write "<font size='4'><b>" & rsArticle("TitleIntact") & "</b></font>"
    Else
        Response.Write "<font size='4'><b>" & rsArticle("Title") & "</b></font>"
    End If
    If Trim(rsArticle("Subheading")) <> "" Then
        Response.Write "<br>" & rsArticle("Subheading")
    End If
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr align='center' class='tdbg'>"
    Response.Write "    <td colspan='2'>"
    Dim Author, CopyFrom
    Author = rsArticle("Author")
    CopyFrom = rsArticle("CopyFrom")
    Response.Write "作者：" & Author & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "来源：" & CopyFrom
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;点击数：" & rsArticle("Hits") & "&nbsp;&nbsp;&nbsp;&nbsp;更新时间：" & FormatDateTime(rsArticle("UpdateTime"), 2) & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "属性："
    If rsArticle("OnTop") = True Then
        Response.Write "<font color=blue>顶</font>&nbsp;"
    Else
        Response.Write "&nbsp;&nbsp;&nbsp;"
    End If
    If rsArticle("Hits") >= HitsOfHot Then
        Response.Write "<font color=red>热</font>&nbsp;"
    Else
        Response.Write "&nbsp;&nbsp;&nbsp;"
    End If
    If rsArticle("Elite") = True Then
        Response.Write "<font color=green>荐</font>"
    Else
        Response.Write "&nbsp;&nbsp;"
    End If
    Response.Write "&nbsp;&nbsp;<font color='#009900'>" & String(rsArticle("Stars"), "★") & "</font>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td colspan='2'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='5'>"
    Response.Write "        <tr>"
    Response.Write "          <td height='200' valign='top'>"
    If Trim(rsArticle("LinkUrl")) <> "" Then
        Response.Write "<p align='center'><br><br><br><font color=red>本" & ChannelShortName & "是链接外部" & ChannelShortName & "内容。链接地址为：<a href='" & rsArticle("LinkUrl") & "' target='_blank'>" & rsArticle("LinkUrl") & "</a></font></p>"
    Else
        Response.Write "<p>" & Replace(Replace(FilterBadTag(rsArticle("Content"), rsArticle("Inputer")), "[InstallDir_ChannelDir]", InstallDir & ChannelDir & "/"), "{$UploadDir}", UploadDir) & "</p>"
    End If
    Response.Write "       </td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr  align='right' class='tdbg'>"
    Response.Write "    <td colspan='2'>"
    Response.Write ChannelShortName & "录入：<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Field=Inputer&Keyword=" & rsArticle("Inputer") & "'>" & rsArticle("Inputer") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;责任编辑："
    If rsArticle("Status") = 3 Then
        Response.Write rsArticle("Editor")
    Else
        Response.Write "无"
    End If
    Response.Write " </td>"
    Response.Write "  </tr>"
    Response.Write "        </tbody>" & vbCrLf

    Call ShowTabs_Special(arrSpecialID, " disabled")

    Call ShowTabs_Purview_Modify("阅读", rsArticle, " disabled")
    
    
    Dim NotReceiveUser, arrUser, i
    If rsArticle("Receive") = True Then
        If rsArticle("Received") = "" Then
            NotReceiveUser = rsArticle("ReceiveUser")
        Else
            NotReceiveUser = ""
            arrUser = Split(rsArticle("ReceiveUser"), ",")
            For i = 0 To UBound(arrUser)
                If FoundInArr(rsArticle("Received"), arrUser(i), "|") = False Then
                    If NotReceiveUser = "" Then
                        NotReceiveUser = arrUser(i)
                    Else
                        NotReceiveUser = NotReceiveUser & "," & arrUser(i)
                    End If
                End If
            Next
        End If
    End If
    
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>要求签收用户：</td>"
    Response.Write "            <td style='width:600; word-wrap:break-word;'>" & rsArticle("ReceiveUser") & "</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>已经签收用户：</td>"
    Response.Write "            <td style='width:600; word-wrap:break-word;'>" & rsArticle("Received") & "</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>尚未签收用户：</td>"
    Response.Write "            <td style='width:600; word-wrap:break-word;'>" & NotReceiveUser & "</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>签收方式：</td>"
    Response.Write "            <td><select name='AutoReceiveTime' disabled>"
    Response.Write "                      <option value='0'"
    If rsArticle("AutoReceiveTime") = "0" Then Response.Write " selected"
    Response.Write ">手动签收</option>"
    Response.Write "                      <option value='5'"
    If rsArticle("AutoReceiveTime") = "5" Then Response.Write " selected"
    Response.Write ">5秒钟后</option>"
    Response.Write "                      <option value='10'"
    If rsArticle("AutoReceiveTime") = "10" Then Response.Write " selected"
    Response.Write ">10秒钟后</option>"
    Response.Write "                      <option value='30'"
    If rsArticle("AutoReceiveTime") = "30" Then Response.Write " selected"
    Response.Write ">30秒钟后</option>"
    Response.Write "                      <option value='60'"
    If rsArticle("AutoReceiveTime") = "60" Then Response.Write " selected"
    Response.Write ">1分钟后</option>"
    Response.Write "                      <option value='120'"
    If rsArticle("AutoReceiveTime") = "120" Then Response.Write " selected"
    Response.Write ">2分钟后</option>"
    Response.Write "                      <option value='300'"
    If rsArticle("AutoReceiveTime") = "300" Then Response.Write " selected"
    Response.Write ">5分钟后</option>"
    Response.Write "                    </select>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>文档类型：</td>"
    Response.Write "            <td><select name='ReceiveType' id='ReceiveType' disabled>"
    Response.Write "                      <option value='0'"
    If rsArticle("ReceiveType") = "0" Then Response.Write " selected"
    Response.Write ">公众文档</option>"
    Response.Write "                      <OPTION value='1'"
    If rsArticle("ReceiveType") = "1" Then Response.Write " selected"
    Response.Write ">专属文档</OPTION>"
    Response.Write "                    </select>"
    Response.Write "            </td>"
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    
    Dim WordNum, strArticle, strSql, rsChannel, payNum
    '过滤空格，HTML字符，计算字数
    WordNum = getWordNumber(rsArticle("Content"))
    
    If MoneyPerKw <= 0 Then
       payNum = 0
    Else
       payNum = MoneyPerKw / 1000 * WordNum
    End If
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>文章字数：</td><td>" & WordNum & " 字</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>支付标准：</td><td>" & MoneyPerKw & " 元/千字</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>估计稿费：</td><td>" & payNum & " 元</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>确认支付稿费：</td><td>" & rsArticle("CopyMoney") & " 元</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>稿费受益者：</td><td>" & rsArticle("Inputer") & "</td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf
    

    Call ShowTabs_MyField_View(rsArticle)

    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf



    Response.Write "<form name='formA' method='get' action='Admin_Article.asp'><p align='center'>"
    Response.Write "<input type='hidden' name='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input type='hidden' name='ArticleID' value='" & ArticleID & "'>"
    Response.Write "<input type='hidden' name='Action' value=''>" & vbCrLf

    If rsArticle("Deleted") = False Then
        PurviewChecked = CheckClassPurview("Manage", ClassID)
        PurviewChecked2 = CheckClassPurview("Check", ClassID)
        If (rsArticle("Inputer") = UserName And rsArticle("Status") = 0) Or PurviewChecked = True Then
            Response.Write "<input type='submit' name='submit' value='修改/审核' onclick=""document.formA.Action.value='Modify'"">&nbsp;&nbsp;"
            Response.Write "<input type='submit' name='submit' value=' 删 除 ' onclick=""document.formA.Action.value='Del'"">&nbsp;&nbsp;"
        End If
        If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
            Response.Write "<input type='submit' name='submit' value=' 移 动 ' onclick=""document.formA.Action.value='MoveToClass'"">&nbsp;&nbsp;"
        End If
        If PurviewChecked2 = True Then
            If rsArticle("Status") > -1 Then
                Response.Write "<input type='submit' name='submit' value='直接退稿' onclick=""document.formA.Action.value='Reject'"">&nbsp;&nbsp;"
            End If
            If rsArticle("Status") < MyStatus Then
                Response.Write "<input type='submit' name='submit' value='" & arrStatus(MyStatus) & "' onclick=""document.formA.Action.value='SetPassed'"">&nbsp;&nbsp;"
            End If
            If rsArticle("Status") >= MyStatus Then
                Response.Write "<input type='submit' name='submit' value='取消审核' onclick=""document.formA.Action.value='CancelPassed'"">&nbsp;&nbsp;"
            End If
        End If
        If PurviewChecked = True Then
            If rsArticle("OnTop") = False Then
                Response.Write "<input type='submit' name='submit' value='设为固顶' onclick=""document.formA.Action.value='SetOnTop'"">&nbsp;&nbsp;"
            Else
                Response.Write "<input type='submit' name='submit' value='取消固顶' onclick=""document.formA.Action.value='CancelOnTop'"">&nbsp;&nbsp;"
            End If
            If rsArticle("Elite") = False Then
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

    rsArticle.Close
    Set rsArticle = Nothing

    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='0'><tr class='tdbg'><td>"
    Response.Write "<li>上一" & ChannelItemUnit & ChannelShortName & "："
    Dim rsPrev
    Set rsPrev = Conn.Execute("Select Top 1 A.ArticleID,A.Title,C.ClassID,C.ClassName from PE_Article A left join PE_Class C On A.ClassID=C.ClassID Where A.ChannelID=" & ChannelID & " and A.Deleted=" & PE_False & " and A.ArticleID<" & ArticleID & " order by A.ArticleID desc")
    If rsPrev.EOF Then
        Response.Write "没有了"
    Else
        Response.Write "[<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ClassID=" & rsPrev("ClassID") & "'>" & rsPrev("ClassName") & "</a>] <a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & rsPrev("ArticleID") & "'>" & rsPrev("Title") & "</a>"
    End If
    rsPrev.Close
    Set rsPrev = Nothing
    Response.Write "</li></td</tr><tr class='tdbg'><td><li>下一" & ChannelItemUnit & ChannelShortName & "："
    Dim rsNext
    Set rsNext = Conn.Execute("Select Top 1 A.ArticleID,A.Title,C.ClassID,C.ClassName from PE_Article A left join PE_Class C On A.ClassID=C.ClassID Where A.ChannelID=" & ChannelID & " and A.Deleted=" & PE_False & " and A.ArticleID>" & ArticleID & " order by A.ArticleID asc")
    If rsNext.EOF Then
        Response.Write "没有了"
    Else
        Response.Write "[<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&ClassID=" & rsNext("ClassID") & "'>" & rsNext("ClassName") & "</a>] <a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & rsNext("ArticleID") & "'>" & rsNext("Title") & "</a>"
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
        Response.Write " class='title5' onclick=""window.location.href='Admin_Article.asp?Action=Show&ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & "&InfoType=0'"""
    End If
    Response.Write ">相关评论</td><td"
    If InfoType = 1 Then
        Response.Write " class='title6'"
    Else
        Response.Write " class='title5' onclick=""window.location.href='Admin_Article.asp?Action=Show&ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & "&InfoType=1'"""
    End If
    Response.Write ">相关收费</td>"
    Response.Write "<td>&nbsp;</td></tr></table>"
    
    strFileName = "Admin_Article.asp?Action=Show&ChannelID=" & ChannelID & "&ArticleID=" & ArticleID & "&InfoType=" & InfoType
    
    Select Case InfoType
    Case 0
        Call ShowComment(ArticleID)
    Case 1
        Call ShowConsumeLog(ArticleID)
    End Select
End Sub

Sub WriteJS_Show()
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
End Sub

Sub Preview()
    Call WriteJS_Show
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td width='400' height='22'>"

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

    Select Case Trim(Request("IncludePic"))
        Case 1
            Response.Write "<font color=blue>" & ArticlePro1 & "</font>"
        Case 2
            Response.Write "<font color=blue>" & ArticlePro2 & "</font>"
        Case 3
            Response.Write "<font color=blue>" & ArticlePro3 & "</font>"
        Case 4
            Response.Write "<font color=blue>" & ArticlePro4 & "</font>"
    End Select
    
    Response.Write PE_HTMLEncode(Request("Title"))
    Response.Write " </td>"
    Response.Write "    <td width='50' height='22' align='right'>"
    If LCase(Request("OnTop")) = "yes" Then
        Response.Write "顶&nbsp;"
    Else
        Response.Write "&nbsp;&nbsp;&nbsp;"
    End If
    If LCase(Request("Hot")) = "yes" Then
        Response.Write "热&nbsp;"
    Else
        Response.Write "&nbsp;&nbsp;&nbsp;"
    End If
    If LCase(Request("Elite")) = "yes" Then
        Response.Write "荐"
    Else
        Response.Write "&nbsp;&nbsp;"
    End If
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'><td colspan='3' align='center'><font size='4'>"
    If Trim(Request("TitleIntact")) <> "" Then
        Response.Write "<font size='4'><b>" & PE_HTMLEncode(Request("TitleIntact")) & "</b></font>"
    Else
        Response.Write "<font size='4'><b>" & PE_HTMLEncode(Request("Title")) & "</b></font>"
    End If
    If Trim(Request("Subheading")) <> "" Then
        Response.Write "<br>" & PE_HTMLEncode(Request("Subheading"))
    End If

    Response.Write "</font></td></tr>"
    Response.Write "  <tr class='tdbg'><td colspan='3' align='center'>作者：" & PE_HTMLEncode(Request("Author")) & "&nbsp;&nbsp;&nbsp;&nbsp;转贴自：" & PE_HTMLEncode(Request("CopyFrom")) & "&nbsp;&nbsp;&nbsp;&nbsp;点击数：0&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "录入：" & UserName & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td colspan='3'><p>" & Request("Content") & "</p></td></tr>"
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

    
    ArticleID = Replace(ArticleID, " ", "")
    Response.Write "<form method='POST' name='myform' action='Admin_Article.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
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
    Response.Write "              <input type='text' name='BatchArticleID' value='" & ArticleID & "' size='28'><br>"
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
    Response.Write "            <td><input name='Author' type='text' id='Author' value='" & Trim(Session("Author")) & "' size='15' maxlength='30'> " & GetAuthorList("Admin", ChannelID, UserName)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyCopyFrom' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "来源：</td>"
    Response.Write "            <td><input name='CopyFrom' type='text' id='CopyFrom' value='" & Trim(Session("CopyFrom")) & "' size='15' maxlength='50'> " & GetCopyFromList("Admin", ChannelID)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyPaginationType' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>内容分页方式：</td>"
    Response.Write "            <td><select name='PaginationType' id='PaginationType'>"
    Response.Write "                <option value='0' selected>不分页</option>"
    Response.Write "                <option value='1'>自动分页</option>"
    Response.Write "                <option value='2'>手动分页</option>"
    Response.Write "              </select>"
    Response.Write "              自动分页时的每页大约字符数（包含HTML标记且必须大于100）：<input name='MaxCharPerPage' type='text' id='MaxCharPerPage' value='10000' size='8' maxlength='8'>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Call ShowBatchCommon
    Response.Write "        </tbody>" & vbCrLf

    Call ShowTabs_Purview_Batch("阅读")
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
    Response.Write "    <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
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
    
    Dim rs, sql, BatchType, BatchArticleID, BatchClassID, rsField
    Dim Author, ShowCommentLink, CopyFrom, PaginationType, MaxCharPerPage
    Dim Keyword, OnTop, Elite, Stars, Hits, UpdateTime, SkinID, TemplateID
    Dim InfoPurview, arrGroupID, InfoPoint, ChargeType, PitchTime, ReadTimes, DividePercent
    
    BatchType = PE_CLng(Trim(Request("BatchType")))
    BatchArticleID = Trim(Request.Form("BatchArticleID"))
    BatchClassID = FilterArrNull(Trim(Request.Form("BatchClassID")), ",")
    Author = Trim(Request.Form("Author"))
    ShowCommentLink = Trim(Request.Form("ShowCommentLink"))
    CopyFrom = Trim(Request.Form("CopyFrom"))
    PaginationType = PE_CLng(Trim(Request.Form("PaginationType")))
    MaxCharPerPage = PE_CLng(Trim(Request.Form("MaxCharPerPage")))
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
        If IsValidID(BatchArticleID) = False Then
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
    If Trim(Request("ModifyPaginationType")) = "Yes" And PaginationType = 1 And MaxCharPerPage = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定自动分页时的每页大约字符数，必须大于0</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    If Author = "" Then Author = XmlText("BaseText", "DefAuthor", "佚名")
    If CopyFrom = "" Then CopyFrom = XmlText("BaseText", "DefCopyFrom", "本站原创")

    If ShowCommentLink = "Yes" Then
        ShowCommentLink = True
    Else
        ShowCommentLink = False
    End If
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
        sql = "select * from PE_Article where ChannelID=" & ChannelID & " and ArticleID in (" & BatchArticleID & ")"
    Else
        sql = "select * from PE_Article where ChannelID=" & ChannelID & " and ClassID in (" & BatchClassID & ")"
    End If
    rs.Open sql, Conn, 1, 3
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=-1 or ChannelID=" & ChannelID & "")
    Do While Not rs.EOF
        If Trim(Request("ModifyAuthor")) = "Yes" Then rs("Author") = Author
        If Trim(Request("ModifyCopyFrom")) = "Yes" Then rs("CopyFrom") = CopyFrom
        If Trim(Request("ModifyCommentLink")) = "Yes" Then rs("ShowCommentLink") = ShowCommentLink
        If Trim(Request("ModifyPaginationType")) = "Yes" Then
            rs("PaginationType") = PaginationType
            rs("MaxCharPerPage") = MaxCharPerPage
        End If
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

        If Not (rsField.BOF And rsField.EOF) Then
            rsField.MoveFirst
            Do While Not rsField.EOF
                If Trim(Request("Modify" & Trim(rsField("FieldName")))) = "Yes" Then
                    If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                        rs(Trim(rsField("FieldName"))) = Trim(Request(rsField("FieldName")))
                    End If
                End If
                rsField.MoveNext
            Loop
        End If
        rs("CreateTime") = rs("UpdateTime")

        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    rsField.Close
    Set rsField = Nothing
    Call ClearSiteCache(0)

    Call WriteSuccessMsg("批量修改" & ChannelShortName & "属性成功！", "Admin_Article.asp?ChannelID=" & ChannelID)
    Call CreateAllJS
End Sub
'=================================================
'过程名：BatchReplace
'作  用：批量替换
'=================================================
Sub BatchReplace()

    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If

    Response.Write "<SCRIPT language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.BatchClassID.length;i++){" & vbCrLf
    Response.Write "    document.myform.BatchClassID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.BatchClassID.length;i++){" & vbCrLf
    Response.Write "    document.myform.BatchClassID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

    ArticleID = Replace(ArticleID, " ", "")
    Response.Write "<form method='POST' name='myform' action='Admin_Article.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><b>批量替换" & ChannelShortName & "内容</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='center'>" & vbCrLf
    Response.Write "       <td class='tdbg' valign='top'>"
    Response.Write "        <table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <tr>"
    Response.Write "            <td align='center'><b>" & ChannelShortName & "范围</b></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <input type='radio' name='BatchType' value='1' checked>指定" & ChannelShortName & "ID：<br>"
    Response.Write "              <input type='text' name='BatchArticleID' value='" & ArticleID & "' size='28'><br>"
    Response.Write "              <input type='radio' name='BatchType' value='2'>指定栏目的" & ChannelShortName & "：<br>"
    Response.Write "              <select name='BatchClassID' size='2' multiple style='height:280px;width:180px;'>" & GetClass_Option(0, 0) & "</select><br><div align='center'>"
    Response.Write "              <input type='button' name='Submit' value='  选定所有栏目  ' onclick='SelectAll()'><br>"
    Response.Write "              <input type='button' name='Submit' value='取消选定所有栏目' onclick='UnSelectAll()'></div></td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "     </td>" & vbCrLf
    Response.Write "      <td valign='top'>" & vbCrLf
    Response.Write "       <table width='100%' height='400' border='0' cellpadding='0' cellspacing='1'>"
    Response.Write "         <tr class='tdbg'>" & vbCrLf
    Response.Write "           <td width='150' align=""right"" class='tdbg5'><strong>替换内容：&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td align='left'>" & vbCrLf
    Response.Write "            <INPUT TYPE='checkbox' NAME='ItemBatchTitle'  value='yes' >" & ChannelShortName & "标题&nbsp;&nbsp;"
    Response.Write "            <INPUT TYPE='checkbox' NAME='ItemBatchContent'  value='yes' checked>" & ChannelShortName & "内容</td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class='tdbg'>" & vbCrLf
    Response.Write "           <td width='150' align=""right"" class='tdbg5'><strong>替换类型：&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td align='left'>" & vbCrLf
    Response.Write "            <INPUT TYPE='radio' NAME='ItemBatchType' onClick=""javascript:PE_ItemReplaceStart.style.display='none';PE_ItemReplaceEnd.style.display='none';PE_ItemReplace.style.display='';"" value='1' checked>简单替换&nbsp;&nbsp;"
    Response.Write "            <INPUT TYPE='radio' NAME='ItemBatchType' onClick=""javascript:PE_ItemReplaceStart.style.display='';PE_ItemReplaceEnd.style.display='';PE_ItemReplace.style.display='none';"" value='2' >高级替换</td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_ItemReplace' style='display:'> " & vbCrLf
    Response.Write "           <td width=""150"" class='tdbg5' align=""right""><strong> 要替换的字符：&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='ItemReplace' ROWS='' COLS='' style='width:400px;height:100px'></TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_ItemReplaceStart' style='display:none'> " & vbCrLf
    Response.Write "           <td width=""150"" class='tdbg5' align=""right"" ><strong> 要替换的开始字符：&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='ItemReplaceStart' ROWS='' COLS='' style='width:400px;height:100px'></TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_ItemReplaceEnd' style='display:none'> " & vbCrLf
    Response.Write "           <td width=""150"" class='tdbg5' align=""right"" ><strong> 要替换的结束字符：&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='ItemReplaceEnd' ROWS='' COLS='' style='width:400px;height:100px'></TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_ItemReplaceResult' style='display:'>" & vbCrLf
    Response.Write "           <td width=""150"" class='tdbg5'  align=""right""><strong> 要替换后的字符：&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='ItemReplaceResult' ROWS='' COLS='' style='width:400px;height:100px'></TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class='tdbg'>" & vbCrLf
    Response.Write "           <td width='150' align=""right"" class='tdbg5'><strong>是否标题加前缀：&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td align='left'>" & vbCrLf
    Response.Write "            <INPUT TYPE='radio' NAME='IsTitlePrefix' onClick=""javascript:PE_TitlePrefix.style.display='';"" value='1' >是&nbsp;&nbsp;"
    Response.Write "            <INPUT TYPE='radio' NAME='IsTitlePrefix' onClick=""javascript:PE_TitlePrefix.style.display='none';"" value='0' checked>否</td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_TitlePrefix' style='display:none'>" & vbCrLf
    Response.Write "           <td width=""150"" class='tdbg5'  align=""right""><strong> 给标题前缀字符：&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='ItemTitlePrefix' ROWS='' COLS='' style='width:400px;height:100px'></TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class='tdbg'>" & vbCrLf
    Response.Write "           <td width='150' align=""right"" class='tdbg5'><strong>是否内容加前缀：&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td align='left'>" & vbCrLf
    Response.Write "            <INPUT TYPE='radio' NAME='IsContentPrefix' onClick=""javascript:PE_ContentPrefix.style.display='';"" value='1' >是&nbsp;&nbsp;"
    Response.Write "            <INPUT TYPE='radio' NAME='IsContentPrefix' onClick=""javascript:PE_ContentPrefix.style.display='none';"" value='0' checked>否</td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg"" id='PE_ContentPrefix' style='display:none'>" & vbCrLf
    Response.Write "           <td width=""150"" class='tdbg5'  align=""right""><strong> 给内容加前缀字符：&nbsp;</strong></td>" & vbCrLf
    Response.Write "           <td class=""tdbg""  valign='top'> <TEXTAREA NAME='ItemContentPrefix' ROWS='' COLS='' style='width:400px;height:100px'></TEXTAREA></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "         <tr class=""tdbg""> " & vbCrLf
    Response.Write "           <td colspan=""2"" align=""center"" class=""tdbg"" height=""50"">" & vbCrLf
    Response.Write "            <input name=""Action"" type=""hidden"" id=""Action"" value=""BatchReplace"">" & vbCrLf
    Response.Write "            <input name=""ChannelID"" type=""hidden"" id=""ChannelID"" value=" & ChannelID & ">" & vbCrLf
    Response.Write "            <input name=""Cancel"" type=""button"" id=""Cancel"" value=""返回上一步"" onClick=""window.location.href='javascript:history.go(-1)'"" >&nbsp;&nbsp;" & vbCrLf
    Response.Write "            <input  type=""submit"" name=""Submit"" value="" 开始替换 "" onClick=""document.myform.Action.value='DoBatchReplace';"" >" & vbCrLf
    Response.Write "           </td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "       </table>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

'=================================================
'过程名：DoBatchReplace
'作  用：批量替换处理
'=================================================
Sub DoBatchReplace()

    Dim rs, sql, BatchType, BatchArticleID, BatchClassID, ChannelID
    Dim ItemBatchType, ItemReplace, ItemReplaceStart, ItemReplaceEnd, ItemReplaceResult
    Dim ItemBatchTitle, ItemBatchContent, IsTitlePrefix, ItemTitlePrefix, IsContentPrefix, ItemContentPrefix
    Dim FoundErr, ErrMsg

    BatchType = PE_CLng(Trim(Request("BatchType")))
    BatchArticleID = Trim(Request.Form("BatchArticleID"))
    BatchClassID = FilterArrNull(Trim(Request.Form("BatchClassID")), ",")
    ChannelID = PE_CLng(Trim(Request.Form("ChannelID")))

    ItemBatchType = PE_CLng(Trim(Request.Form("ItemBatchType")))
    ItemBatchTitle = Trim(Request.Form("ItemBatchTitle"))
    ItemBatchContent = Trim(Request.Form("ItemBatchContent"))
    ItemReplace = Trim(Request.Form("ItemReplace"))
    ItemReplaceStart = Trim(Request.Form("ItemReplaceStart"))
    ItemReplaceEnd = Trim(Request.Form("ItemReplaceEnd"))
    ItemReplaceResult = Trim(Request.Form("ItemReplaceResult"))

    IsTitlePrefix = PE_CLng(Trim(Request.Form("IsTitlePrefix")))
    ItemTitlePrefix = Trim(Request.Form("ItemTitlePrefix"))
    IsContentPrefix = PE_CLng(Trim(Request.Form("IsContentPrefix")))
    ItemContentPrefix = Trim(Request.Form("ItemContentPrefix"))

    If IsTitlePrefix = 0 Then
        ItemTitlePrefix = ""
    Else
        If Len(ItemTitlePrefix) > 100 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>标题前缀，不能过长</li>"
        End If
    End If

    If IsContentPrefix = 0 Then
        ItemContentPrefix = ""
    End If

    If BatchType = 1 Then
        If IsValidID(BatchArticleID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要批量修改的" & ChannelShortName & "的ID</li>"
        End If
    Else
        If IsValidID(BatchClassID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要批量修改的" & ChannelShortName & "的栏目</li>"
        End If
    End If

    If ItemBatchTitle = "yes" Then
        ItemBatchTitle = True
    End If
    If ItemBatchContent = "yes" Then
        ItemBatchContent = True
    End If

    If ItemBatchTitle = False And ItemBatchContent = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>至少要选择一个要替换的类型" & ChannelShortName & "标题或" & ChannelShortName & "内容</li>"
    End If

    If ItemBatchType = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>没有选择" & ChannelShortName & "替换字符类型</li>"
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    If ItemBatchType = 1 Then
        If ItemReplace = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>输入要替换的代码不能为空</li>"
        End If
    ElseIf ItemBatchType = 2 Then
        If ItemReplaceStart = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>输入要替换的开始代码不能为空</li>"
        End If
        If ItemReplaceEnd = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>输入要替换后的结束代码不能为空</li>"
        End If
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>选择" & ChannelShortName & "替换字符类型不对</li>"
    End If

    If ItemReplaceResult = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>输入要替换后的代码不能为空</li>"
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    If ItemBatchTitle = True Then
        If PE_CLng(Conn.Execute("Select count(*) From PE_Article Where Title='" & ReplaceBadChar(ItemReplaceResult) & "'")(0)) > 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>要替换的标题与数据库已有的标题重复</li>"
        End If
        If FoundErr = True Then
            Call WriteErrMsg(ErrMsg, ComeUrl)
            Exit Sub
        End If
    End If

    Response.Write "<li>正在替换数据请稍后..</li>&nbsp;&nbsp;<br>"

    Set rs = Server.CreateObject("ADODB.Recordset")
    If BatchType = 1 Then
        sql = "select * from PE_Article where ChannelID=" & ChannelID & " and ArticleID in (" & BatchArticleID & ")"
    Else
        sql = "select * from PE_Article where ChannelID=" & ChannelID & " and ClassID in (" & BatchClassID & ")"
    End If
    rs.Open sql, Conn, 1, 3
    If rs.BOF And rs.EOF Then
        Response.Write "没有可替换的标题或正文！"
    Else
        Do While Not rs.EOF
            If ItemBatchType = 1 Then
                If ItemBatchTitle = True Then
                    If InStr(rs("title"), ItemReplace) <> 0 Then
                        rs("title") = ItemTitlePrefix & Replace(rs("title"), ItemReplace, ItemReplaceResult)
                        Response.Write "<br>&nbsp;&nbsp;" & ChannelShortName & "ID：" & rs("ArticleID") & "&nbsp;&nbsp;" & rs("title") & "..<font color='#009900'>标题替换成功！</font>"
                    End If
                End If
                If ItemBatchContent = True Then
                    If InStr(rs("Content"), ItemReplace) <> 0 Then
                        rs("Content") = ItemContentPrefix & Replace(rs("Content"), ItemReplace, ItemReplaceResult)
                        Response.Write "<br>&nbsp;&nbsp;" & ChannelShortName & "ID：" & rs("ArticleID") & "&nbsp;&nbsp;" & rs("title") & "..<font color='#009900'>内容替换成功！</font>"
                    End If
                End If
            ElseIf ItemBatchType = 2 Then
                If ItemBatchTitle = True Then
                    rs("title") = ItemTitlePrefix & BatchReplaceString(rs("title"), ItemReplaceStart, ItemReplaceEnd, ItemReplaceResult, "标题", rs("ArticleID"), rs("title"))
                End If
                If ItemBatchContent = True Then
                    rs("Content") = ItemContentPrefix & BatchReplaceString(rs("Content"), ItemReplaceStart, ItemReplaceEnd, ItemReplaceResult, "内容", rs("ArticleID"), rs("title"))
                End If
            End If
            rs("CreateTime") = rs("UpdateTime")

            rs.Update
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Call ClearSiteCache(0)
    Response.Write "<br>&nbsp;&nbsp;<font color='red'>" & ChannelShortName & "替换操作完成</font>"
    Response.Write "<br><center>&nbsp;&nbsp;<a href='Admin_Article.asp?ChannelID=" & ChannelID & "'>返回" & ChannelShortName & "管理</a> </center>"
End Sub

'**************************************************
'函数名：BatchReplaceString
'作  用：批量替换处理函数
'参  数：ItemContent ----数据
'参  数：ItemReplaceStart ----获得要替换的开头代码
'参  数：ItemReplaceEnd ----获得要替换的结束代码
'参  数：ItemReplaceResult ----要替换的代码
'参  数：ItemName ----类型名称
'返回值：True  ----已创建
'**************************************************
Function BatchReplaceString(ItemContent, ItemReplaceStart, ItemReplaceEnd, ItemReplaceResult, ItemName, ArticleID, Title)
    If InStr(ItemContent, ItemReplaceStart) = 0 Or InStr(ItemContent, ItemReplaceEnd) = 0 Then
        BatchReplaceString = ItemContent
        Exit Function
    End If
    If GetBody(ItemContent, ItemReplaceStart, ItemReplaceEnd, True, True) = "" Then
        BatchReplaceString = ItemContent
        Exit Function
    End If
    BatchReplaceString = Replace(ItemContent, GetBody(ItemContent, ItemReplaceStart, ItemReplaceEnd, True, True), ItemReplaceResult)
    Response.Write "<br>&nbsp;&nbsp;" & ChannelShortName & "ID：" & ArticleID & "&nbsp;&nbsp;" & Title & "..<font color='#009900'>" & ItemName & "替换成功！</font>"
End Function


'******************************************************************************************
'以下为设置固顶、推荐等属性使用的函数，各模块实现过程类似，修改时注意同时修改各模块内容。
'******************************************************************************************

Sub SetProperty()
    If ArticleID = "" Then
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
    If InStr(ArticleID, ",") > 0 Then
        sqlProperty = "select * from PE_Article where ArticleID in (" & ArticleID & ")"
    Else
        sqlProperty = "select * from PE_Article where ArticleID=" & ArticleID
    End If
    Set rsProperty = Server.CreateObject("ADODB.Recordset")
    rsProperty.Open sqlProperty, Conn, 1, 3
    Do While Not rsProperty.EOF
        If CheckClassPurview(Action, rsProperty("ClassID")) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>对 " & rsProperty("Title") & " 没有操作权限！</li>"
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
    Call WriteSuccessMsg("操作成功！", "Admin_Article.asp?ChannelID=" & ChannelID)

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
    
    Dim ArticleType, BatchArticleID, BatchClassID
    Dim tChannelID, tClassID, tChannelDir, tUploadDir
    
    ArticleType = PE_CLng(Trim(Request("ArticleType")))
    BatchArticleID = Trim(Request.Form("BatchArticleID"))
    BatchClassID = FilterArrNull(Trim(Request.Form("BatchClassID")), ",")
    tChannelID = Trim(Request("tChannelID"))
    tClassID = Trim(Request("tClassID"))
    
    If ArticleType = 1 Then
        If IsValidID(BatchArticleID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要批量移动的" & ChannelShortName & "的ID</li>"
        End If
    Else
        If IsValidID(BatchClassID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要批量移动的" & ChannelShortName & "的栏目</li>"
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
    
    Dim rsBatchMove, sqlBatchMove, ArticlePath
    sqlBatchMove = "select A.ArticleID,A.UploadFiles,A.UpdateTime,A.PaginationType,C.ParentDir,C.ClassDir from PE_Article A left join PE_Class C on A.ClassID=C.ClassID"
    If ArticleType = 1 Then
        sqlBatchMove = sqlBatchMove & " where A.ChannelID=" & ChannelID & " and A.ArticleID in (" & BatchArticleID & ")"
    Else
        sqlBatchMove = sqlBatchMove & " where A.ChannelID=" & ChannelID & " and A.ClassID in (" & BatchClassID & ")"
    End If
    BatchArticleID = ""
    Set rsBatchMove = Conn.Execute(sqlBatchMove)
    Do While Not rsBatchMove.EOF
        ArticlePath = HtmlDir & GetItemPath(StructureType, rsBatchMove("ParentDir"), rsBatchMove("ClassDir"), rsBatchMove("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsBatchMove("UpdateTime"), rsBatchMove("ArticleID"))
        If fso.FileExists(Server.MapPath(ArticlePath & FileExt_Item)) Then
            DelSerialFiles Server.MapPath(ArticlePath & FileExt_Item)
        End If
        If rsBatchMove("PaginationType") > 0 Then
            DelSerialFiles Server.MapPath(ArticlePath) & "_*" & FileExt_Item
        End If
        If tChannelID <> ChannelID Then
            Call MoveUpFiles(rsBatchMove("UploadFiles") & "", tChannelDir & "/" & tUploadDir)    '移动上传文件
        End If
        If BatchArticleID = "" Then
            BatchArticleID = rsBatchMove("ArticleID")
        Else
            BatchArticleID = BatchArticleID & "," & rsBatchMove("ArticleID")
        End If
        rsBatchMove.MoveNext
    Loop
    rsBatchMove.Close
    Set rsBatchMove = Nothing
    If BatchArticleID <> "" Then
        Conn.Execute ("update PE_Article set ChannelID=" & tChannelID & ",ClassID=" & tClassID & ",TemplateID=0,CreateTime=UpdateTime where ArticleID in (" & BatchArticleID & ")")
    End If

    Call WriteSuccessMsg("成功将选定的" & ChannelShortName & "移动到目标频道的目标栏目中！", "Admin_Article.asp?ChannelID=" & ChannelID & "")
    Call ClearSiteCache(0)
End Sub


Sub MoveUpFiles(strFiles, strTargetDir)
    On Error Resume Next
    Dim strTrueFile, arrFiles, strTrueDir, i
    If IsNull(strFiles) Or strFiles = "" Or strTargetDir = "" Then Exit Sub
    
    If Not fso.FolderExists(Server.MapPath(InstallDir & strTargetDir)) Then fso.CreateFolder Server.MapPath(InstallDir & strTargetDir)
    
    arrFiles = Split(strFiles, "|")
    For i = 0 To UBound(arrFiles)
        strTrueDir = Server.MapPath(InstallDir & strTargetDir & "/" & Left(arrFiles(i), InStr(arrFiles(i), "/")))
        If Not fso.FolderExists(strTrueDir) Then fso.CreateFolder strTrueDir
        strTrueFile = Server.MapPath(InstallDir & ChannelDir & "/" & UploadDir & "/" & arrFiles(i))
        If fso.FileExists(strTrueFile) Then
            fso.MoveFile strTrueFile, Server.MapPath(InstallDir & strTargetDir & "/" & arrFiles(i))
        End If
    Next
End Sub

'******************************************************************************************
'以下为删除、清空、还原等操作使用的函数，各模块实现过程类似，修改时注意同时修改各模块内容。
'******************************************************************************************

Sub Del()
    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先选定" & ChannelShortName & "！</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel, ArticlePath, arrUser
    arrUser = ""
    sqlDel = "select A.ArticleID,A.Title,A.UpdateTime,A.CreateTime,A.Inputer,A.Status,A.Deleted,A.PaginationType,A.PresentExp,A.ReceiveUser,A.ClassID,C.ParentDir,C.ClassDir from PE_Article A left join PE_Class C on A.ClassID=C.ClassID"
    If InStr(ArticleID, ",") > 0 Then
        sqlDel = sqlDel & " where A.ArticleID in (" & ArticleID & ") order by A.ArticleID"
    Else
        sqlDel = sqlDel & " where A.ArticleID=" & ArticleID
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
            ErrMsg = ErrMsg & "<li>删除 <font color='red'>" & rsDel("Title") & "</font> 失败！原因：没有操作权限！</li>"
        Else
            If FoundInArr(arrUser, rsDel("Inputer"), ",") = True Then
                If arrUser = "" Then
                    arrUser = rsDel("Inputer")
                Else
                    arrUser = arrUser & "," & rsDel("Inputer")
                End If
            End If
            ArticlePath = HtmlDir & GetItemPath(StructureType, rsDel("ParentDir"), rsDel("ClassDir"), rsDel("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsDel("UpdateTime"), rsDel("ArticleID"))
            If fso.FileExists(Server.MapPath(ArticlePath & FileExt_Item)) Then
                DelSerialFiles Server.MapPath(ArticlePath & FileExt_Item)
            End If
            If rsDel("PaginationType") > 0 Then
                DelSerialFiles Server.MapPath(ArticlePath) & "_*" & FileExt_Item
            End If

            If rsDel("Status") = 3 Then
                Conn.Execute ("update PE_User set UserExp=UserExp-" & rsDel("PresentExp") & " where UserName='" & rsDel("Inputer") & "'")
            End If
            rsDel("Deleted") = True
            rsDel("CreateTime") = rsDel("UpdateTime")
            Call Del_User_UnsignedItems(rsDel("ArticleID"), rsDel("ReceiveUser"))
            rsDel.Update
        End If
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing

    Call UpdateChannelData(ChannelID)
    Call UpdateUserData(0, arrUser, 0, 0)

    Call ClearSiteCache(0)
    Call WriteSuccessMsg("操作成功！", "Admin_Article.asp?ChannelID=" & ChannelID)
    Call CreateAllJS
End Sub

Sub DelFile()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先选定" & ChannelShortName & "！</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel, ArticlePath
    sqlDel = "select A.ArticleID,A.UpdateTime,A.PaginationType,C.ParentDir,C.ClassDir from PE_Article A left join PE_Class C on A.ClassID=C.ClassID where A.ArticleID in (" & ArticleID & ") order by A.ArticleID"
    Set rsDel = Conn.Execute(sqlDel)
    Do While Not rsDel.EOF
        ArticlePath = HtmlDir & GetItemPath(StructureType, rsDel("ParentDir"), rsDel("ClassDir"), rsDel("UpdateTime")) & GetItemFileName(FileNameType, ChannelDir, rsDel("UpdateTime"), rsDel("ArticleID"))
        If fso.FileExists(Server.MapPath(ArticlePath & FileExt_Item)) Then
            DelSerialFiles Server.MapPath(ArticlePath & FileExt_Item)
        End If
        If rsDel("PaginationType") > 0 Then
            DelSerialFiles Server.MapPath(ArticlePath) & "_*" & FileExt_Item
        End If
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    Conn.Execute ("update PE_Article set CreateTime=UpdateTime where ArticleID in (" & ArticleID & ")")
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub ConfirmDel()
    If AdminPurview = 2 And AdminPurview_Channel > 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，你的权限不够！</li>"
        Exit Sub
    End If
    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先选定" & ChannelShortName & "！</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel
    sqlDel = "select UploadFiles,VoteID from PE_Article where ArticleID in (" & ArticleID & ")"
    Set rsDel = Conn.Execute(sqlDel)
    Do While Not rsDel.EOF
        Call DelUploadFiles(rsDel("UploadFiles"))
        Conn.Execute ("delete from PE_Vote where ID=" & rsDel("VoteID") & "")
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    Conn.Execute ("delete from PE_Article where ArticleID in (" & ArticleID & ")")
    Conn.Execute ("delete from PE_Comment where ModuleType=" & ModuleType & " and InfoID in (" & ArticleID & ")")
    Conn.Execute ("delete from PE_ConsumeLog where ModuleType=" & ModuleType & " and InfoID in (" & ArticleID & ")")
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
    ArticleID = ""
    sqlDel = "select ArticleID,UploadFiles,VoteID from PE_Article where Deleted=" & PE_True & " and ChannelID=" & ChannelID
    Set rsDel = Conn.Execute(sqlDel)
    Do While Not rsDel.EOF
        If ArticleID = "" Then
            ArticleID = rsDel(0)
        Else
            ArticleID = ArticleID & "," & rsDel(0)
        End If
        Call DelUploadFiles(rsDel("UploadFiles"))
        Conn.Execute ("delete from PE_Vote where ID=" & rsDel("VoteID") & "")
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    If ArticleID <> "" Then
        Conn.Execute ("delete from PE_Article where Deleted=" & PE_True & " and ChannelID=" & ChannelID & "")
        Conn.Execute ("delete from PE_Comment where ModuleType=" & ModuleType & " and InfoID in (" & ArticleID & ")")
        Conn.Execute ("delete from PE_ConsumeLog where ModuleType=" & ModuleType & " and InfoID in (" & ArticleID & ")")
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
    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先选定" & ChannelShortName & "！</li>"
        Exit Sub
    End If
    
    Dim sqlDel, rsDel, arrUser
    arrUser = ""
    If InStr(ArticleID, ",") > 0 Then
        sqlDel = "select * from PE_Article where ArticleID in (" & ArticleID & ")"
    Else
        sqlDel = "select * from PE_Article where ArticleID=" & ArticleID
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
    Call WriteSuccessMsg("操作成功！", "Admin_Article.asp?ChannelID=" & ChannelID)
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
    sqlDel = "select * from PE_Article where Deleted=" & PE_True & " and ChannelID=" & ChannelID
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
    Call WriteSuccessMsg("操作成功！", "Admin_Article.asp?ChannelID=" & ChannelID)
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

Sub Del_User_UnsignedItems(ByVal ArticleID, ByVal strUser)
    If IsNull(strUser) Or strUser = "" Then Exit Sub
    strUser = Replace(strUser, "|", ",")
    strUser = Replace(strUser, ",", "','")
    
    Dim arrID
    arrID = Split(ArticleID, ",")
    Dim sqlUser, rsUser, i, tmpUnsignedItems, tmpArticleID
    Set rsUser = Server.CreateObject("adodb.recordset")
    sqlUser = "select UserID,UserName,UnsignedItems from PE_User where UserName in ('" & strUser & "')"
    rsUser.Open sqlUser, Conn, 1, 3
    Do While Not rsUser.EOF
        For i = 0 To UBound(arrID)
            If FoundInArr(rsUser("UnsignedItems"), CStr(arrID(i)), ",") = True Then
                tmpUnsignedItems = "," & rsUser("UnsignedItems") & ","
                tmpArticleID = "," & arrID(i) & ","
                tmpUnsignedItems = Replace(tmpUnsignedItems, tmpArticleID, ",")
                If tmpUnsignedItems = "," Then
                    rsUser("UnsignedItems") = ""
                Else
                    rsUser("UnsignedItems") = Mid(tmpUnsignedItems, 2, Len(tmpUnsignedItems) - 2)
                End If
                rsUser.Update
            End If
        Next
        rsUser.MoveNext
    Loop
    rsUser.Close
    Set rsUser = Nothing
End Sub

Sub Add_User_UnsignedItems(ByVal ArticleID, ByVal ReceiveUser)
    If IsNull(ReceiveUser) Or ReceiveUser = "" Then Exit Sub
    ReceiveUser = Replace(ReceiveUser, "|", ",")
    ReceiveUser = Replace(ReceiveUser, ",", "','")
    
    Dim sqlUser, rsUser, i
    
    Set rsUser = Server.CreateObject("adodb.recordset")
    sqlUser = "select UserID,UserName,UnsignedItems from PE_User where UserName in ('" & ReceiveUser & "')"
    rsUser.Open sqlUser, Conn, 1, 3
    Do While Not rsUser.EOF
        If rsUser("UnsignedItems") = "" Or IsNull(rsUser("UnsignedItems")) Then
            rsUser("UnsignedItems") = ArticleID
        Else
            If FoundInArr(rsUser("UnsignedItems"), CStr(ArticleID), ",") = False Then
                rsUser("UnsignedItems") = rsUser("UnsignedItems") & "," & ArticleID
            End If
        End If
        rsUser.Update
        rsUser.MoveNext
    Loop
    rsUser.Close
    Set rsUser = Nothing
End Sub

Function UnsignedItemsState(ArticleID)
    Dim rsState, sqlState, strState, arrUser, i
    Dim NotReceiveUser
    sqlState = "select top 1 ReceiveUser,Received from PE_Article where ArticleID=" & ArticleID
    Set rsState = Conn.Execute(sqlState)
    If Not (rsState.BOF And rsState.EOF) Then
        If rsState("Received") = "" Then
            NotReceiveUser = rsState("ReceiveUser")
        Else
            NotReceiveUser = ""
            arrUser = Split(rsState("ReceiveUser"), ",")
            For i = 0 To UBound(arrUser)
                If FoundInArr(rsState("Received"), arrUser(i), "|") = False Then
                    If NotReceiveUser = "" Then
                        NotReceiveUser = arrUser(i)
                    Else
                        NotReceiveUser = NotReceiveUser & "," & arrUser(i)
                    End If
                End If
            Next
        End If
        If NotReceiveUser <> "" Then
            strState = strState & "<a href='' onclick='return false' title='"
            strState = strState & "要求签收用户：" & rsState("ReceiveUser") & vbCrLf
            strState = strState & "已经签收用户：" & rsState("Received") & vbCrLf
            strState = strState & "尚未签收用户：" & NotReceiveUser
            strState = strState & "'><font color=red>[签收中]</font></a>"
        Else
            strState = strState & "<a href='#' title='"
            strState = strState & "签收用户：" & rsState("ReceiveUser")
            strState = strState & "'><font color=green>[已签收]</font></a>"
        End If
    End If
    rsState.Close
    Set rsState = Nothing
    UnsignedItemsState = strState
End Function



'*****************************************
'函 数 名：getWordNumber()
'参    数：str 字符串
'返 回 值：文章字数
'作    用：计算文章的字数 可以计算纯中文，纯英文，中英混排，误差范围在20字以内
'作    者：壮志，刘永涛
'创建日期：2005-09-07
'*****************************************
Function getWordNumber(ByVal str)
    str = nohtml(PE_HtmlDecode(str))
    regEx.Pattern = "[a-z\-]+|\.+"
    str = regEx.Replace(str, "动")
    str = Replace(str, " ", "")
    getWordNumber = Len(str)
End Function


Sub ExportExcel()
    Dim strSql, SelectType, rsArticleOut, PayStatus, searchDate
    SelectType = Trim(Request("SelectType")) '检查输入的ID，和 日期
    strSql = "Select * From PE_Article Where Copymoney>0  And Status=3"
    PayStatus = Trim(Request("PayStatus"))
    If PayStatus = "False" Then
        PayStatus = PE_False
        searchDate = "UpDateTime"
    Else
        PayStatus = PE_True
        searchDate = "PayDate"
    End If
    Select Case SelectType
    Case "ID"
        Dim BeginID, EndID
        BeginID = PE_CLng(Trim(Request("BeginID")))
        EndID = PE_CLng(Trim(Request("EndID")))
        If BeginID <> 0 And EndID <> 0 Then
            strSql = strSql & "And Ispayed=" & PayStatus & " and (ArticleID Between " & BeginID & " and " & EndID & ")"
        End If
    Case "Date"
        Dim BeginDate, EndDate
        BeginDate = Trim(Request("BeginDate"))
        EndDate = Trim(Request("EndDate"))
        If BeginDate = "" Then
            BeginDate = "1900-1-1"
        Else
            BeginDate = ReplaceBadChar(BeginDate)
        End If
        If EndDate = "" Then
            EndDate = FormatDateTime(Date, 2)
        Else
            EndDate = ReplaceBadChar(EndDate)
        End If
        If IsDate(BeginDate) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入正确的起始日期！</li>"
        End If
        If IsDate(EndDate) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入正确的结束日期！</li>"
        End If
        If FoundErr = True Then
            Exit Sub
        End If
        If SystemDatabaseType = "SQL" Then
            strSql = strSql & "And Ispayed=" & PayStatus & " and UpdateTime Between '" & BeginDate & "' and '" & EndDate & "'"
        Else
            strSql = strSql & "And Ispayed=" & PayStatus & " and UpdateTime Between #" & BeginDate & "# and #" & EndDate & "#"
        End If
    Case Else
        If InStr(ArticleID, ",") >= 0 And ArticleID <> "" Then
            strSql = strSql & "And Ispayed=" & PayStatus & " And ArticleID in (" & ArticleID & ")"
        Else
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请先选定" & ChannelShortName & "！</li>"
            Exit Sub
        End If
    End Select
    Set rsArticleOut = Conn.Execute(strSql)
    If rsArticleOut.BOF And rsArticleOut.EOF Then
        Response.Write "<script language='javascript'>alert('没有查询到相关数据')</script>"
    Else
        Call outHead2
        Response.Write "<table border=""0"" cellspacing=""1"" style=""border-collapse: collapse;table-layout:fixed"" id=""AutoNumber1"" height=""32"">" & vbCrLf
        Response.Write "<tr>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>文章ID</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>文章标题</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>录入者</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>作者</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>稿费受益者</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>文章字数</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>稿费(单位：元)</b></span></td>" & vbCrLf
        Response.Write "<td align=""center""><span lang=""zh-cn""><b>已支付</b></span></td>" & vbCrLf
        If PayStatus = PE_True Then
            Response.Write "<td align=""center""><span lang=""zh-cn""><b>支付时间</b></span></td>" & vbCrLf
        Else
            Response.Write "<td align=""center""><span lang=""zh-cn""><b>录入时间</b></span></td>" & vbCrLf
        End If
        Response.Write "</tr>" & vbCrLf
        Do While Not rsArticleOut.EOF
            Response.Write "<tr>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("ArticleID") & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("Title") & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("Inputer") & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("Author") & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("Beneficiary") & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & getWordNumber(rsArticleOut("Content")) & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("Copymoney") & "</span></td>" & vbCrLf
            Response.Write "<td align=""center""><span lang=""zh-cn"">"
            If rsArticleOut("IsPayed") Then
                Response.Write "是"
            Else
                Response.Write "否"
            End If
            Response.Write "</span></td>" & vbCrLf
            If PayStatus = PE_True Then
                Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("payDate") & "</span></td>" & vbCrLf
            Else
                Response.Write "<td align=""center""><span lang=""zh-cn"">" & rsArticleOut("UpdateTime") & "</span></td>" & vbCrLf
            End If
            Response.Write "</tr>" & vbCrLf
            rsArticleOut.MoveNext
        Loop
        rsArticleOut.Close
        Set rsArticleOut = Nothing
        Response.Write "</table>" & vbCrLf
    End If
End Sub

Sub ConfirmPay()
   Dim strSql, rsArticle, i, arrArticleID
   strSql = "Update PE_Article Set IsPayed=" & PE_True & ",PayDate=" & PE_Now & " Where Copymoney>0 And Ispayed=" & PE_False & " And Status=3"
   If InStr(ArticleID, ",") >= 0 And ArticleID <> "" Then
      strSql = strSql & " And ArticleID in (" & ArticleID & ")"
   Else
      FoundErr = True
      ErrMsg = ErrMsg & "<li>请先选定" & ChannelShortName & "！</li>"
      Exit Sub
   End If
   Conn.Execute (strSql)
   Conn.Close
   Set Conn = Nothing
   Response.Redirect "Admin_Article.asp?ChannelID=" & ChannelID & "&ManageType=PayMoney&Status=9"
End Sub

Sub outHead2()
    Response.Write "<html><head>" & vbCrLf
    Response.ContentType = "application/vnd.ms-excel" & vbCrLf
    Response.AddHeader "Content-Disposition", "attachment"
    Response.Write "<meta http-equiv=""Content-Language"" content=""zh-cn"">" & vbCrLf
    Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
    Response.Write "<title>稿费列表</title>" & vbCrLf
    Response.Write "<body>" & vbCrLf
End Sub


%>
