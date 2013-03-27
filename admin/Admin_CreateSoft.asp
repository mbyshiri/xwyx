<!--#include file="Admin_CreateCommon.asp"-->
<!--#include file="../Include/PowerEasy.Soft.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim PE_Content
Set PE_Content = New Soft
PE_Content.Init
tmpPageTitle = strPageTitle    '保存页面标题到临时变量中，以做为栏目及内容页循环生成时初始值
tmpNavPath = strNavPath
SoftID = Trim(Request("SoftID"))
Select Case Action
Case "CreateSoft"
    Call CreateSoft
Case "CreateClass"
    Call CreateClass
Case "CreateSpecial"
    Call CreateSpecial
Case "CreateIndex"
    Call CreateIndex
Case "CreateSoft2"
    If AutoCreateType > 0 Then
        IsAutoCreate = True
        Call CreateSoft
        If ClassID > 0 Then
            ClassID = ParentPath & "," & ClassID
            Call CreateClass
        End If
        SpecialID = Trim(Request("SpecialID"))
        If SpecialID <> "" Then Call CreateSpecial
        '在生成首页前，要将栏目ID和专题ID置为0
        ClassID = 0
        arrChildID = 0
        SpecialID = 0
        Call CreateIndex
        Call CreateSiteIndex     '生成网站首页
        Call CreateSiteSpecial   '生成全站专题
    End If
Case "CreateOther" '定时生成创建除内容页以外的其他页
    TimingCreate = Trim(Request("TimingCreate"))
    TimingCreateNum = PE_CLng(Trim(Request("TimingCreateNum")))

    If Trim(Request("ChannelProperty")) <> "" Then
        CreateChannelItem = Split(Trim(Request("ChannelProperty")), ",")
        ChannelID = CreateChannelItem(0)
        CreateType = 2

        If CreateChannelItem(5) = "True" Then
            Call CreateClass
            Call CreateAllJS
            Response.Write "<b>生成所有JS文件成功！</b><br>"
        End If

        If CreateChannelItem(6) = "True" Then
            Call CreateSpecial
        End If

        If CreateChannelItem(7) = "True" Then
            Call CreateIndex
        End If

        If TimingCreateNum >= UBound(Split(TimingCreate, "$")) Then
            Call CreateSiteIndex     '生成网站首页
        End If

        TimingCreateNum = TimingCreateNum + 1
        strFileName = "Admin_Timing.asp?Action=DoTiming&TimingCreateNum=" & TimingCreateNum & "&TimingCreate=" & Trim(Request("TimingCreate"))
    End If
    
    If Trim(Request("TimingCreate")) <> "" Then
        Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
        Response.Write "function aaa(){window.location.href='" & strFileName & "';}" & vbCrLf
        Response.Write "setTimeout('aaa()',5000);" & vbCrLf
        Response.Write "</script>" & vbCrLf
    End If
Case Else
    FoundErr = True
    ErrMsg = ErrMsg & "<li>参数错误！</li>"
End Select

Call ShowProcess

Response.Write "</body></html>"
Set PE_Content = Nothing
Call CloseConn


Sub CreateSoft()
    'On Error Resume Next
    Dim sql, strFields, SoftPath
    Dim tmpSoft, tmpTemplateID
    
    tmpTemplateID = 0
    
    If IsAutoCreate = False Then
        Response.Write "<b>正在生成" & ChannelShortName & "页面……请稍候！<font color='red'>在此过程中请勿刷新此页面！！！</font></b><br>"
        Response.Flush
    End If
    
    sql = "select * from PE_Soft where Deleted=" & PE_False & " and Status=3 and ChannelID=" & ChannelID
    Select Case CreateType
    Case 1 '选定的下载
        If IsValidID(SoftID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请正确指定要生成的" & ChannelShortName & "ID</li>"
            Exit Sub
        End If
        If InStr(SoftID, ",") > 0 Then
            sql = sql & " and SoftID in (" & SoftID & ")"
        Else
            sql = sql & " and SoftID=" & SoftID
        End If
        strUrlParameter = "&SoftID=" & SoftID
    Case 2 '选定的栏目
        ClassID = PE_CLng(Trim(Request("ClassID")))
        If ClassID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要生成的栏目ID</li>"
            Exit Sub
        End If
        Call GetClass
        If ClassPurview > 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>此栏目不是开放栏目，所以此栏目下的下载不能生成HTML！"
        End If
        If FoundErr = True Then Exit Sub
        If InStr(arrChildID, ",") > 0 Then
            sql = sql & " and ClassID in (" & arrChildID & ")"
        Else
            sql = sql & " and ClassID=" & ClassID
        End If
    Case 3 '所有下载
        
    Case 4 '最新的下载
        Dim TopNew
        TopNew = PE_CLng(Trim(Request("TopNew")))
        If TopNew <= 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定有效的数目！"
            Exit Sub
        End If
        sql = "select top " & TopNew & " * from PE_Soft where Deleted=" & PE_False & " and Status=3 and ChannelID=" & ChannelID
        strUrlParameter = "&TopNew=" & TopNew
    Case 5 '指定更新时间
        Dim BeginDate, EndDate
        BeginDate = Trim(Request("BeginDate"))
        EndDate = Trim(Request("EndDate"))
        If Not (IsDate(BeginDate) And IsDate(EndDate)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入有效的日期！</li>"
            Exit Sub
        End If
        If SystemDatabaseType = "SQL" Then
            sql = sql & " and UpdateTime between '" & BeginDate & "' and '" & EndDate & "'"
        Else
            sql = sql & " and UpdateTime between #" & BeginDate & "# and #" & EndDate & "#"
        End If
        strUrlParameter = "&BeginDate=" & BeginDate & "&EndDate=" & EndDate
    Case 6 '指定ID范围
        Dim BeginID, EndID
        BeginID = Trim(Request("BeginID"))
        EndID = Trim(Request("EndID"))
        If Not (IsNumeric(BeginID) And IsNumeric(EndID)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入数字！</li>"
            Exit Sub
        End If
        sql = sql & " and SoftID between " & BeginID & " and " & EndID
        strUrlParameter = "&BeginID=" & BeginID & "&EndID=" & EndID
    Case 8 '定时间生成范围
        TimingCreate = Trim(Request("TimingCreate"))
        ChannelProperty = Trim(Request("ChannelProperty"))
        TimingCreateNum = PE_CLng(Trim(Request("TimingCreateNum")))
        IsShowReturn = True
        arrChannelProperty = Split(ChannelProperty, ",")
        ChannelID = arrChannelProperty(0)
        CreateItemType = arrChannelProperty(2)
        CreateItemTopNewNum = arrChannelProperty(3)
        CreateItemDate = arrChannelProperty(4)
        Select Case CreateItemType
        Case 1
            
        Case 2
            sql = sql & " and DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ")<" & CreateItemDate & ""
        Case 3
            sql = "select top " & MaxPerPage_Create & " * from PE_Soft where Deleted=" & PE_False & " and Status=3 and ChannelID=" & ChannelID
			sql = sql & " and (CreateTime is null or CreateTime<=UpdateTime)"
        Case 4
            
        End Select
        strUrlParameter = "&TimingCreate=" & TimingCreate & "&TimingCreateNum=" & TimingCreateNum & "&ChannelProperty=" & Trim(Request("ChannelProperty"))
    Case 9 '所有未生成的软件
        sql = "select top " & MaxPerPage_Create & " * from PE_Soft where Deleted=" & PE_False & " and Status=3 and ChannelID=" & ChannelID
		sql = sql & " and (CreateTime is null or CreateTime<=UpdateTime)"
    Case Else
        Response.Write "参数错误！"
        Exit Sub
    End Select
    Set rsSoft = Server.CreateObject("ADODB.Recordset")
    rsSoft.Open sql, Conn, 1, 1
    If rsSoft.Bof And rsSoft.EOF Then
        TotalCreate = 0
		iTotalPage = 0
        rsSoft.Close
        Set rsSoft = Nothing
        Exit Sub
    Else
        If CreateType = 9 Or (CreateType = 8 And CreateItemType = 3)Then
			TotalCreate = PE_Clng(Conn.Execute("select count(*) from PE_Soft where Deleted=" & PE_False & " and Status=3 and ChannelID=" & ChannelID & " and (CreateTime is null or CreateTime<=UpdateTime)")(0))
		Else
			TotalCreate = rsSoft.RecordCount
		End If
    End If
    
    PageTitle = ChannelShortName & "信息"
    strFileName = ChannelUrl_ASPFile & "/ShowSoft.asp"

    strTemplate = GetTemplate(ChannelID, 3, tmpTemplateID)

    Call MoveRecord(rsSoft)
    Call ShowTotalCreate(ChannelItemUnit & ChannelShortName)

    Do While Not rsSoft.EOF
        FoundErr = False
        strPageTitle = tmpPageTitle
        strNavPath = tmpNavPath
        ClassID = rsSoft("ClassID")
        If ChannelID <> PrevChannelID Then
            Call GetChannel(ChannelID)
            PrevChannelID = ChannelID
        End If
        Call GetClass

        SpecialID = 0
        SoftID = rsSoft("SoftID")
        CurrentPage = 1
        
        SoftPath = HtmlDir & GetItemPath(StructureType, ParentDir, ClassDir, rsSoft("UpdateTime"))
        If CreateMultiFolder(SoftPath) = False Then
            Response.Write "请检查服务器。系统不能创建生成文件所需要的文件夹，"
            Exit Sub
        End If
        SoftPath = SoftPath & GetItemFileName(FileNameType, ChannelDir, rsSoft("UpdateTime"), SoftID)
            
        tmpFileName = SoftPath & FileExt_Item

        SkinID = GetIDByDefault(rsSoft("SkinID"), DefaultItemSkin)
        TemplateID = GetIDByDefault(rsSoft("TemplateID"), DefaultItemTemplate)
            
        SoftName = Replace(Replace(Replace(Replace(rsSoft("SoftName") & "", "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")

        If TemplateID <> tmpTemplateID Then
            strTemplate = GetTemplate(ChannelID, 3, TemplateID)
            tmpTemplateID = TemplateID
        End If

        strHTML = strTemplate
        Call PE_Content.GetHtml_Soft
        'strHTML = Replace(strHTML, "{$DownloadUrl}", CreateDownloadUrl(rsSoft("DownloadUrl"),iShowModule))
        Call WriteToFile(tmpFileName, strHTML)

        iCount = iCount + 1
        Response.Write "<li>成功生成第 <font color='red'><b>" & iCount & " </b></font> " & ChannelItemUnit & ChannelShortName & " 生成的ID </b><FONT color='Red'>" & SoftID & "</FONT> 地址 " & tmpFileName & "</li><br>" & vbCrLf
        Response.Flush
        '生成内容结束，更新内容的生成时间
        Conn.Execute ("update PE_Soft set CreateTime=" & PE_Now & " where SoftID=" & SoftID)

        If Response.IsClientConnected = False Then Exit Do
        If iCount Mod MaxPerPage_Create = 0 Then Exit Do
        rsSoft.MoveNext
    Loop
    rsSoft.Close
    Set rsSoft = Nothing
End Sub

%>
