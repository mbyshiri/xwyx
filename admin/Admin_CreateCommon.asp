<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Class.asp"-->
<!--#include file="../Include/PowerEasy.Special.asp"-->
<!--#include file="../Include/PowerEasy.Common.Front.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 0      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

'定义生成相关的变量
Dim CreateType, IsAutoCreate, tmpFileName, tmpPageTitle, tmpNavPath
Dim Pages
Dim TotalCreate, CurrentCreatePage, iCount, iTotalPage
Dim IsShowReturn
Dim strUrlParameter

Dim TimingCreate, CollectionCreateHTML, ArticleNum
Dim CreateItemType, CreateItemTopNewNum, CreateItemDate
Dim ChannelProperty, arrChannelProperty, TimingCreateNum
Dim CreateChannelItem, CreateNum, arrTimingCreate

If ChannelID = 0 Then
    'Response.Write "频道参数丢失！"
    'Response.End
End If
ClassID = Trim(Request("ClassID"))
SpecialID = Trim(Request("SpecialID"))

IsAutoCreate = False
CreateType = Trim(Request("CreateType"))

If CreateType = "" Then
    CreateType = 1
Else
    CreateType = PE_CLng(CreateType)
End If
CurrentCreatePage = PE_CLng(Trim(Request("CreatePage")))


Response.Write "<html><head><title>" & ChannelShortName & "生成</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf


Sub CreateIndex()
    If ChannelID <> PrevChannelID Then
        Call GetChannel(ChannelID)
        PrevChannelID = ChannelID
    End If
    If UseCreateHTML = 0 Then
        Response.Write "因为此频道设置了“不生成HTML”，所以不用生成首页。"
        Exit Sub
    End If
    
    Response.Write "<b>正在生成此频道的首页（" & HtmlDir & "/Index" & FileExt_Index & "）……"
    MaxPerPage = MaxPerPage_Index
    strPageTitle = ""
    SkinID = DefaultSkinID
    PageTitle = "首页"
    strFileName = ChannelUrl_ASPFile & "/Index.asp"
    strPageTitle = SiteTitle
    strNavPath = XmlText("BaseText", "Nav", "您现在的位置：") & "&nbsp;<a class='LinkPath' href='" & SiteUrl & "'>" & SiteName & "</a>"
    CurrentPage = 1
    
    If ShowNameOnPath = True And ChannelName <> "" Then
        strPageTitle = strPageTitle & " >> " & ChannelName & " >> " & PageTitle
        strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;<a class='LinkPath'   href='" & ChannelUrl & "/Index" & FileExt_Index & "'>" & ChannelName & "</a>&nbsp;" & strNavLink & "&nbsp;" & PageTitle
    Else
        strPageTitle = strPageTitle & " >> " & PageTitle
        strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;" & PageTitle
    End If

    Call PE_Content.GetHTML_Index
    Call WriteToFile(HtmlDir & "/Index" & FileExt_Index, strHtml)

    Response.Write "……………………生成首页成功！</b>" & vbCrLf
End Sub


Sub CreateClass()
    'On Error Resume Next
    If ChannelID <> PrevChannelID Then
        Call GetChannel(ChannelID)
        PrevChannelID = ChannelID
    End If
    If UseCreateHTML = 0 Or UseCreateHTML = 2 Then
        Response.Write "<b>因为此频道设置了“不生成HTML”或“栏目页不生成HTML”，所以不用生成栏目页。</b><br>"
        Exit Sub
    End If
    
    Dim rsCreate, sql
    Dim tmpDir, tmpTemplateID
    If IsAutoCreate = False Then
        Response.Write "<b>正在生成栏目列表页面……请稍候！<font color='red'>在此过程中请勿刷新此页面！！！</font></b><br>"
        Response.Flush
    End If
    sql = "select * from PE_Class where ClassType=1 and ClassPurview<2 and ChannelID=" & ChannelID
    Select Case CreateType
    Case 1 '选定的栏目
		If Action = "CreateOther" Then
			IsAutoCreate = True
            ClassID = PE_Clng(Trim(Request("ClassID")))
			Call GetClass()
			ClassID = ParentPath & "," & ClassID
		End If
        If IsValidID(ClassID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要生成的栏目ID</li>"
            Exit Sub
        End If
        If Left(ClassID, 1) = "," Then
            ClassID = Right(ClassID, Len(ClassID) - 1)
        End If
        If InStr(ClassID, ",") > 0 Then
            sql = sql & " and ClassID in (" & ClassID & ")"
        Else
            sql = sql & " and ClassID=" & ClassID & ""
        End If
    Case 2 '所有栏目
         
    Case Else
        Response.Write "参数错误！"
        Exit Sub
    End Select
    sql = sql & " order by RootID,OrderID"
    tmpTemplateID = 0
    strTemplate = GetTemplate(ChannelID, 2, 0)
    arrTemplate = Split(strTemplate, "{$$$}")
    If UBound(arrTemplate) < 1 Then
        Response.Write "您的栏目页模板有误，缺少小类模板！"
        Exit Sub
    End If
    
    Set rsCreate = Server.CreateObject("ADODB.Recordset")
    rsCreate.Open sql, Conn, 1, 1
    If rsCreate.Bof And rsCreate.EOF Then
        TotalCreate = 0
        rsCreate.Close
        Set rsCreate = Nothing
        Exit Sub
    Else
        TotalCreate = rsCreate.RecordCount
    End If
    
    Call MoveRecord(rsCreate)
    Call ShowTotalCreate("个栏目")
    
    Do While Not rsCreate.EOF
        PageTitle = ""
        FoundErr = False
        If rsCreate("TemplateID") <> tmpTemplateID Then
            strTemplate = GetTemplate(ChannelID, 2, rsCreate("TemplateID"))
            arrTemplate = Split(strTemplate, "{$$$}")
            If UBound(arrTemplate) < 1 Then
                Response.Write rsCreate("ClassName") & "使用的栏目页模板有误，缺少小类模板！"
                Exit Sub
            End If
            
            tmpTemplateID = rsCreate("TemplateID")
        End If
        If ChannelID <> PrevChannelID Then
            Call GetChannel(ChannelID)
            PrevChannelID = ChannelID
        End If
        CurrentPage = 1
        strPageTitle = tmpPageTitle
        strNavPath = tmpNavPath
        ClassID = rsCreate("ClassID")
        strFileName = ChannelUrl_ASPFile & "/ShowClass.asp?ClassID=" & ClassID
        Call GetClass
        tmpDir = HtmlDir & GetListPath(StructureType, ListFileType, ParentDir, ClassDir)

        If CreateMultiFolder(tmpDir) = False Then
            Response.Write "请检查服务器。系统不能创建生成文件所需要的文件夹。"
            Exit Sub
        End If
        tmpFileName = tmpDir & GetListFileName(ListFileType, ClassID, CurrentPage, CurrentPage) & FileExt_List
        
        Call PE_Content.GetHtml_Class
        Call WriteToFile(tmpFileName, strHtml)

        iCount = iCount + 1
        Response.Write "<li>成功生成第 <font color='red'><b>" & iCount & " </b></font>个栏目的列表：" & tmpFileName & "</li><br>" & vbCrLf
        Response.Flush

        ItemCount = PE_CLng(Conn.Execute("select Count(*) from PE_Article where ClassID=" & ClassID & "")(0))
        If Child > 0 And ItemCount > 0 Then
            ClassShowType = 2
            tmpFileName = tmpDir & GetList_1FileName(ListFileType, ClassID) & FileExt_List
            
            Call PE_Content.GetHtml_Class
            Call WriteToFile(tmpFileName, strHtml)

            Response.Write "&nbsp;&nbsp;&nbsp;成功生成第 <font color='red'><b>" & iCount & " </b></font>个栏目的第 <font color='blue'>" & CurrentPage & "</font> 页列表：" & tmpFileName & "<br>" & vbCrLf
            Response.Flush
        End If
        
        If UseCreateHTML = 1 And (IsAutoCreate = False Or (IsAutoCreate = True And AutoCreateType = 1)) Then
            If TotalPut Mod MaxPerPage = 0 Then
                Pages = TotalPut \ MaxPerPage
            Else
                Pages = TotalPut \ MaxPerPage + 1
            End If
            If Pages > 1 Then
                For CurrentPage = 2 To Pages
                    If ChannelID <> PrevChannelID Then
                        Call GetChannel(ChannelID)
                        PrevChannelID = ChannelID
                    End If
                    tmpFileName = tmpDir & GetListFileName(ListFileType, ClassID, CurrentPage, Pages) & FileExt_List
                    If IsAutoCreate = True And CurrentPage > UpdatePages Then
                        Call Update_ShowPage(tmpFileName, "UpdateClass")
                        'If CurrentPage = Pages Then Response.Write "&nbsp;&nbsp;&nbsp;成功更新第 <font color='red'><b>" & iCount & " </b></font>个栏目的第 <font color='blue'>" & UpdatePages + 1 & " 至 " & Pages & "</font> 页<br>" & vbCrLf
                    Else
                        Call PE_Content.GetHtml_Class
                        Call WriteToFile(tmpFileName, strHtml)
                        Response.Write "&nbsp;&nbsp;&nbsp;成功生成第 <font color='red'><b>" & iCount & " </b></font>个栏目的第 <font color='blue'>" & CurrentPage & "</font> 页列表：" & tmpFileName & "<br>" & vbCrLf
                        Response.Flush
                    End If
                Next
            End If
        End If
        ClassShowType = 1
        rsCreate.MoveNext
        If iCount Mod MaxPerPage_Create = 0 Then Exit Do
    Loop
    rsCreate.Close
    Set rsCreate = Nothing
End Sub

Sub CreateSpecial()
    'On Error Resume Next
    If ChannelID <> PrevChannelID Then
        Call GetChannel(ChannelID)
        PrevChannelID = ChannelID
    End If
    If UseCreateHTML = 0 Or UseCreateHTML = 2 Then
        Response.Write "<b>因为此频道设置了“不生成HTML”或“专题页不生成HTML”，所以不用生成专题页。</b><br>"
        Exit Sub
    End If

    tmpDir = HtmlDir & "/Special"
    If CreateMultiFolder(tmpDir) = False Then
        Response.Write "请检查服务器。系统不能创建生成文件所需要的文件夹。"
        Exit Sub
    End If
    
    If IsAutoCreate = False Then
        Response.Write "<b>正在生成专题列表页面……请稍候！<font color='red'>在此过程中请勿刷新此页面！！！</font></b><br>"
        Response.Flush
    End If
    Dim rsCreate, sql
    Dim tmpDir, tmpTemplateID
    PageTitle = ""
    sql = "select * from PE_Special where ChannelID=" & ChannelID
    Select Case CreateType
    Case 1 '选定的专题
		If Action = "CreateOther" Then
			IsAutoCreate = True
		End If
        If IsValidID(SpecialID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要生成的专题ID</li>"
            Exit Sub
        End If
        If InStr(SpecialID, ",") > 0 Then
            sql = sql & " and SpecialID in (" & SpecialID & ")"
        Else
            sql = sql & "and SpecialID=" & SpecialID
        End If
    Case 2 '所有专题
         
    Case Else
        Response.Write "参数错误！"
        Exit Sub
    End Select
    sql = sql & " order by OrderID"
    Set rsCreate = Server.CreateObject("ADODB.Recordset")
    rsCreate.Open sql, Conn, 1, 1
    If rsCreate.Bof And rsCreate.EOF Then
        TotalCreate = 0
        rsCreate.Close
        Set rsCreate = Nothing
        Exit Sub
    Else
        TotalCreate = rsCreate.RecordCount
    End If


    tmpTemplateID = 0
    strTemplate = GetTemplate(ChannelID, 4, 0)
    Call MoveRecord(rsCreate)
    Call ShowTotalCreate("个专题")

    Do While Not rsCreate.EOF
        If rsCreate("TemplateID") <> tmpTemplateID Then
            strTemplate = GetTemplate(ChannelID, 4, rsCreate("TemplateID"))
            tmpTemplateID = rsCreate("TemplateID")
        End If
        strPageTitle = tmpPageTitle
        strNavPath = tmpNavPath
        CurrentPage = 1
        SpecialID = rsCreate("SpecialID")
        If ChannelID <> PrevChannelID Then
            Call GetChannel(ChannelID)
            PrevChannelID = ChannelID
        End If
        strFileName = ChannelUrl_ASPFile & "/ShowSpecial.asp?ClassID=" & ClassID & "&SpecialID=" & SpecialID
        Call GetSpecial
        MaxPerPage = MaxPerPage_Special
        tmpDir = HtmlDir & "/Special/" & SpecialDir
        If Not fso.FolderExists(Server.MapPath(tmpDir)) Then
            fso.CreateFolder Server.MapPath(tmpDir)
        End If
        
        tmpFileName = tmpDir & "/Index" & FileExt_List
        strHtml = strTemplate
        Call PE_Content.GetHtml_Special
        Call WriteToFile(tmpFileName, strHtml)

        iCount = iCount + 1
        Response.Write "<li>成功生成第 <font color='red'><b>" & iCount & " </b></font>个专题的列表：" & tmpFileName & "</li><br>" & vbCrLf
        Response.Flush
        If UseCreateHTML = 1 And (IsAutoCreate = False Or (IsAutoCreate = True And AutoCreateType = 1)) Then
            If TotalPut Mod MaxPerPage = 0 Then
                Pages = TotalPut \ MaxPerPage
            Else
                Pages = TotalPut \ MaxPerPage + 1
            End If
            If Pages > 1 Then
                For CurrentPage = 2 To Pages

                    tmpFileName = tmpDir & "/List_" & Pages - CurrentPage + 1 & FileExt_List
                    If IsAutoCreate = True And CurrentPage > UpdatePages Then
                        Call Update_ShowPage(tmpFileName, "UpdateSpecial")
                        'If CurrentPage = Pages Then Response.Write "&nbsp;&nbsp;&nbsp;成功更新第 <font color='red'><b>" & iCount & " </b></font>个专题的第 <font color='blue'>" & UpdatePages + 1 & " 至 " & Pages & "</font> 页<br>" & vbCrLf
                    Else
                        strHtml = strTemplate
                        Call PE_Content.GetHtml_Special
                        Call WriteToFile(tmpFileName, strHtml)
                        Response.Write "&nbsp;&nbsp;&nbsp;成功生成第 <font color='red'><b>" & iCount & " </b></font>个专题的第 <font color='blue'>" & CurrentPage & "</font> 页列表：" & tmpFileName & "<br>" & vbCrLf
                        Response.Flush
                    End If
                Next
            End If
        End If
        rsCreate.MoveNext
        If iCount Mod MaxPerPage_Create = 0 Then Exit Do
    Loop
    rsCreate.Close
    Set rsCreate = Nothing
End Sub

Sub CreateSiteIndex()
    Response.Write "<br><iframe id='CreateSiteIndex' width='100%' height='30' frameborder='0' src='Admin_CreateSiteIndex.asp?ShowBack=No'></iframe>"
End Sub

Sub CreateSiteSpecial()
    If Trim(Request("SpecialID")) <> "" Then
        Response.Write "<br><iframe id='CreateSiteSpecial' width='100%' height='30' frameborder='0' src='Admin_CreateSiteSpecial.asp?SpecialID=" & Trim(Request("SpecialID")) & "&ShowBack=No&IsAutoCreate=true'></iframe>"
    End If
End Sub

Sub CreateAllJS()
    Response.Write "<br><iframe id='CreateJS' width='100%' height='100' frameborder='0' src='Admin_CreateJS.asp?ChannelID=" & ChannelID & "&ShowBack=No'></iframe>"
End Sub

Sub Update_ShowPage(FileName, iType)
    Dim hf, strUpdateHtml, strPath, strShowPage, strShowPage_en

    strUpdateHtml = ReadFileContent(FileName)

    Select Case iType
    Case "UpdateClass"
        strPath = ChannelUrl & GetListPath(StructureType, ListFileType, ParentDir, ClassDir)
        If ListFileType > 0 Then
            strShowPage = ShowPage_Html(strPath, ClassID, FileExt_List, "", TotalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName)
            strShowPage_en = ShowPage_en_Html(strPath, ClassID, FileExt_List, "", TotalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName)
        Else
            strShowPage = ShowPage_Html(strPath, 0, FileExt_List, "", TotalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName)
            strShowPage_en = ShowPage_en_Html(strPath, 0, FileExt_List, "", TotalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName)
        End If
    Case "UpdateSpecial"
        strPath = ChannelUrl & "/Special/" & SpecialDir
        strShowPage = ShowPage_Html(strPath, 0, FileExt_List, "", TotalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName)
        strShowPage_en = ShowPage_en_Html(strPath, 0, FileExt_List, "", TotalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName)
    Case "UpdateSiteSpecial"
        strPath = InstallDir & "/Special/" & SpecialDir
        strShowPage = ShowPage_Html(strPath, 0, FileExt_SiteSpecial, "", TotalPut, MaxPerPage, CurrentPage, True, True, "个内容")
        strShowPage_en = ShowPage_en_Html(strPath, 0, FileExt_SiteSpecial, "", TotalPut, MaxPerPage, CurrentPage, True, True, "个内容")
    End Select

    regEx.Pattern = "<!--\s分页开始\s-->([\s\S]*?)<!--\s分页结束\s-->"
    Set Matches = regEx.Execute(strUpdateHtml)
    If Matches.Count > 0 Then
        strUpdateHtml = regEx.Replace(strUpdateHtml, strShowPage)
    End If
    regEx.Pattern = "<!--\sShowPage\sBegin\s-->([\s\S]*?)<!--\sShowPage\sEnd\s-->"
    Set Matches = regEx.Execute(strUpdateHtml)
    If Matches.Count > 0 Then
        strUpdateHtml = regEx.Replace(strUpdateHtml, strShowPage_en)
    End If

    Call WriteToFile(FileName, strUpdateHtml)
End Sub

Sub ShowProcess()
	Dim iCreatePage
	If CreateType = 9 Or (CreateType = 8 And CreateItemType = 3)Then
		iCreatePage = CurrentCreatePage
	Else
		iCreatePage = CurrentCreatePage + 1
    End If
	strFileName = "Admin_Create" & ModuleName & ".asp?Action=" & Action & "&CreateType=" & CreateType & "&ChannelID=" & ChannelID & "&ClassID=" & Trim(Request("ClassID")) & "&SpecialID=" & Trim(Request("SpecialID")) & "&CreatePage=" & iCreatePage & strUrlParameter
    strFileName = Replace(strFileName, " ", "")
    If CurrentCreatePage < iTotalPage Then
        If SleepTime > 0 Then
            Response.Write "<p align='center'>" & SleepTime & "秒后将自动继续生成下一页！</p>" & vbCrLf
        End If
        Call Refresh(strFileName,SleepTime)
    Else
        Response.Write "<p align='center'>已经生成所有页面！</p>" & vbCrLf
        If Trim(Request("ShowBack")) <> "No" And CreateType <> 7 And CreateType <> 8 Then
            Response.Write "<p align='center'><a href='Admin_CreateHTML.asp?ChannelID=" & ChannelID & "'>【返回】</a></p>" & vbCrLf
        End If
        If IsShowReturn = True Then '兼容采集，生成文章后进行生成栏目和首页
            Call Refresh("Admin_CreateArticle.asp?Action=CreateOther&CreateType=1&ChannelID=" & ChannelID & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&CollectionCreateHTML=" & Trim(Request("CollectionCreateHTML")) & "&CreateNum=" & Trim(Request("CreateNum")) & "&ShowBack=No&ChannelProperty=" &  Trim(Request("ChannelProperty")) & "&TimingCreateNum=" & Trim(Request("TimingCreateNum")) & "&TimingCreate=" & Trim(Request("TimingCreate")),5)		
            'Response.Write "<meta http-equiv=""refresh"" content=5;url='Admin_CreateArticle.asp?Action=CreateOther&CreateType=1&ChannelID=" & ChannelID & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&CollectionCreateHTML=" & Trim(Request("CollectionCreateHTML")) & "&CreateNum=" & Trim(Request("CreateNum")) & "&ShowBack=No&ChannelProperty=" &  Trim(Request("ChannelProperty")) & "&TimingCreateNum=" & Trim(Request("TimingCreateNum")) & "&TimingCreate=" & Trim(Request("TimingCreate")) & "'>" & vbCrLf
        End If
    End If
End Sub

Sub MoveRecord(rsCreate)
	If (TotalCreate Mod MaxPerPage_Create) = 0 Then
        iTotalPage = TotalCreate \ MaxPerPage_Create
    Else
        iTotalPage = TotalCreate \ MaxPerPage_Create + 1
    End If
    If CurrentCreatePage < 1 Then
        CurrentCreatePage = 1
    End If
	If CreateType = 9 Or (CreateType = 8 And CreateItemType = 3)Then
		Exit Sub
	End If
    If (CurrentCreatePage - 1) * MaxPerPage_Create > TotalCreate Then
        CurrentCreatePage = iTotalPage
    End If
    If CurrentCreatePage > 1 Then
        If (CurrentCreatePage - 1) * MaxPerPage_Create < TotalCreate Then
			rsCreate.Move (CurrentCreatePage - 1) * MaxPerPage_Create
        Else
            CurrentCreatePage = 1
        End If
    End If
    iCount = (CurrentCreatePage - 1) * MaxPerPage_Create
End Sub

Sub ShowTotalCreate(ItemName)
    If IsAutoCreate = False Then
        Response.Write "总共需要生成 <font color='red'><b>" & TotalCreate & "</b></font> " & ItemName
        Response.Write "，每页生成 <font color='red'><b>" & MaxPerPage_Create & "</b></font> " & ItemName
        Response.Write "，共需要分 <font color='red'><b>" & iTotalPage & "</b></font> 页生成"
        Response.Write "，当前正在生成 <font color='red'><b>" & CurrentCreatePage & "</b></font> 页<br>" & vbCrLf
        Response.Flush
    End If
End Sub
%>
