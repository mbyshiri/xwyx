<!--#include file="Admin_CreateCommon.asp"-->
<!--#include file="../Include/PowerEasy.Soft.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim PE_Content
Set PE_Content = New Soft
PE_Content.Init
tmpPageTitle = strPageTitle    '����ҳ����⵽��ʱ�����У�����Ϊ��Ŀ������ҳѭ������ʱ��ʼֵ
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
        '��������ҳǰ��Ҫ����ĿID��ר��ID��Ϊ0
        ClassID = 0
        arrChildID = 0
        SpecialID = 0
        Call CreateIndex
        Call CreateSiteIndex     '������վ��ҳ
        Call CreateSiteSpecial   '����ȫվר��
    End If
Case "CreateOther" '��ʱ���ɴ���������ҳ���������ҳ
    TimingCreate = Trim(Request("TimingCreate"))
    TimingCreateNum = PE_CLng(Trim(Request("TimingCreateNum")))

    If Trim(Request("ChannelProperty")) <> "" Then
        CreateChannelItem = Split(Trim(Request("ChannelProperty")), ",")
        ChannelID = CreateChannelItem(0)
        CreateType = 2

        If CreateChannelItem(5) = "True" Then
            Call CreateClass
            Call CreateAllJS
            Response.Write "<b>��������JS�ļ��ɹ���</b><br>"
        End If

        If CreateChannelItem(6) = "True" Then
            Call CreateSpecial
        End If

        If CreateChannelItem(7) = "True" Then
            Call CreateIndex
        End If

        If TimingCreateNum >= UBound(Split(TimingCreate, "$")) Then
            Call CreateSiteIndex     '������վ��ҳ
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
    ErrMsg = ErrMsg & "<li>��������</li>"
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
        Response.Write "<b>��������" & ChannelShortName & "ҳ�桭�����Ժ�<font color='red'>�ڴ˹���������ˢ�´�ҳ�棡����</font></b><br>"
        Response.Flush
    End If
    
    sql = "select * from PE_Soft where Deleted=" & PE_False & " and Status=3 and ChannelID=" & ChannelID
    Select Case CreateType
    Case 1 'ѡ��������
        If IsValidID(SoftID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����ȷָ��Ҫ���ɵ�" & ChannelShortName & "ID</li>"
            Exit Sub
        End If
        If InStr(SoftID, ",") > 0 Then
            sql = sql & " and SoftID in (" & SoftID & ")"
        Else
            sql = sql & " and SoftID=" & SoftID
        End If
        strUrlParameter = "&SoftID=" & SoftID
    Case 2 'ѡ������Ŀ
        ClassID = PE_CLng(Trim(Request("ClassID")))
        If ClassID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ���ɵ���ĿID</li>"
            Exit Sub
        End If
        Call GetClass
        If ClassPurview > 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����Ŀ���ǿ�����Ŀ�����Դ���Ŀ�µ����ز�������HTML��"
        End If
        If FoundErr = True Then Exit Sub
        If InStr(arrChildID, ",") > 0 Then
            sql = sql & " and ClassID in (" & arrChildID & ")"
        Else
            sql = sql & " and ClassID=" & ClassID
        End If
    Case 3 '��������
        
    Case 4 '���µ�����
        Dim TopNew
        TopNew = PE_CLng(Trim(Request("TopNew")))
        If TopNew <= 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ����Ч����Ŀ��"
            Exit Sub
        End If
        sql = "select top " & TopNew & " * from PE_Soft where Deleted=" & PE_False & " and Status=3 and ChannelID=" & ChannelID
        strUrlParameter = "&TopNew=" & TopNew
    Case 5 'ָ������ʱ��
        Dim BeginDate, EndDate
        BeginDate = Trim(Request("BeginDate"))
        EndDate = Trim(Request("EndDate"))
        If Not (IsDate(BeginDate) And IsDate(EndDate)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��������Ч�����ڣ�</li>"
            Exit Sub
        End If
        If SystemDatabaseType = "SQL" Then
            sql = sql & " and UpdateTime between '" & BeginDate & "' and '" & EndDate & "'"
        Else
            sql = sql & " and UpdateTime between #" & BeginDate & "# and #" & EndDate & "#"
        End If
        strUrlParameter = "&BeginDate=" & BeginDate & "&EndDate=" & EndDate
    Case 6 'ָ��ID��Χ
        Dim BeginID, EndID
        BeginID = Trim(Request("BeginID"))
        EndID = Trim(Request("EndID"))
        If Not (IsNumeric(BeginID) And IsNumeric(EndID)) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���������֣�</li>"
            Exit Sub
        End If
        sql = sql & " and SoftID between " & BeginID & " and " & EndID
        strUrlParameter = "&BeginID=" & BeginID & "&EndID=" & EndID
    Case 8 '��ʱ�����ɷ�Χ
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
    Case 9 '����δ���ɵ����
        sql = "select top " & MaxPerPage_Create & " * from PE_Soft where Deleted=" & PE_False & " and Status=3 and ChannelID=" & ChannelID
		sql = sql & " and (CreateTime is null or CreateTime<=UpdateTime)"
    Case Else
        Response.Write "��������"
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
    
    PageTitle = ChannelShortName & "��Ϣ"
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
            Response.Write "�����������ϵͳ���ܴ��������ļ�����Ҫ���ļ��У�"
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
        Response.Write "<li>�ɹ����ɵ� <font color='red'><b>" & iCount & " </b></font> " & ChannelItemUnit & ChannelShortName & " ���ɵ�ID </b><FONT color='Red'>" & SoftID & "</FONT> ��ַ " & tmpFileName & "</li><br>" & vbCrLf
        Response.Flush
        '�������ݽ������������ݵ�����ʱ��
        Conn.Execute ("update PE_Soft set CreateTime=" & PE_Now & " where SoftID=" & SoftID)

        If Response.IsClientConnected = False Then Exit Do
        If iCount Mod MaxPerPage_Create = 0 Then Exit Do
        rsSoft.MoveNext
    Loop
    rsSoft.Close
    Set rsSoft = Nothing
End Sub

%>
