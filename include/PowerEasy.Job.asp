<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Class Job

Private strListStr_Font

Public Sub Init()
    ChannelShortName = "ְλ"

    If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))

    strListStr_Font = XmlText_Class("JobList/UpdateTimeColor_New", "color=""red""")

    
    strNavPath = XmlText("BaseText", "Nav", "�����ڵ�λ�ã�") & "&nbsp;<a class='LinkPath' href='" & SiteUrl & "'>" & SiteName & "</a>"
    strPageTitle = SiteTitle
    
    Call GetChannel(ChannelID)
    HtmlDir = strInstallDir & ChannelDir
    If Trim(ChannelName) <> "" And ShowChannelName <> False Then
        If UseCreateHTML > 0 Then
            strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;<a class='LinkPath' href='" & ChannelUrl & "/Index" & FileExt_Index & "'>" & ChannelName & "</a>"
        Else
            strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;<a class='LinkPath' href='" & ChannelUrl & "/Index.asp'>" & ChannelName & "</a>"
        End If
        strPageTitle = strPageTitle & " >> " & ChannelName
    End If
    
End Sub


Private Sub ReleaseDate_OptionJS()
    Call Init
    Dim ReleaseDateJS, ReleaseDate_FileName
    Dim strReleaseDate
    strReleaseDate = "<select name='SearchDateNum'>"
    strReleaseDate = strReleaseDate & "<option value=''>-- ��ѡ�񷢲�����--</option>"
    strReleaseDate = strReleaseDate & "<option value='0'>һ����</option>"
    strReleaseDate = strReleaseDate & "<option value='1'>������</option>"
    strReleaseDate = strReleaseDate & "<option value='2'>������</option>"
    strReleaseDate = strReleaseDate & "<option value='6'>��һ��</option>"
    strReleaseDate = strReleaseDate & "<option value='13'>������</option>"
    strReleaseDate = strReleaseDate & "<option value='29'>��һ����</option>"
    strReleaseDate = strReleaseDate & "<option value='59'>��������</option>"
    strReleaseDate = strReleaseDate & "<option value='89'>��������</option>"
    strReleaseDate = strReleaseDate & "</select>"
    strReleaseDate = "document.write(""" & strReleaseDate & """);"
    ReleaseDate_FileName = strInstallDir & "Job/JS/ReleaseDate_Option.js"
    Call WriteToFile(ReleaseDate_FileName, strReleaseDate)
End Sub


'��������GetPositionList
'��  �ã���ʾְλ���Ƶ���Ϣ
'��  ����
'1        PositionNum ---ְλ����������0����ֻ��ѯǰ����ְλ
'2        IsUrgent ------------�Ƿ��ǽ�����Ƹ��TrueΪֻ��ʾ������Ƹְλ��FalseΪ��ʾ������Ƹְλ
'3        DateNum ----���ڷ�Χ���������0����ֻ��ʾ��������ڸ��µ�ְλ
'4        OrderType ----����ʽ��1----��ְλID����2----��ְλID����3----������ʱ�併��4----������ʱ������
'5        ShowType -----��ʾ��ʽ��1Ϊ������Ƹ��2Ϊ������Ƹ��3Ϊ��ҳ��ʾ��Ƹ��Ϣ�б�
'6        TitleLen  ----ְλ��������ַ�����һ������=����Ӣ���ַ�����Ϊ0������ʾ����ְλ��
'7        WorkPlaceNameLen-----�����ص�����ַ�����һ������=����Ӣ���ַ�����Ϊ0������ʾ����ְλ��
'8        SubCompanyNameLen----���˵�λ����ַ�����һ������=����Ӣ���ַ�����Ϊ0������ʾ����ְλ��
'9        PShowPoints-----ְλ������������ַ���ʱ�Ƿ���ʾʡ�Ժţ�True---��ʾ�� False---����ʾ
'10       WShowPoints-----�����ص�������������ַ���ʱ�Ƿ���ʾʡ�Ժţ�True---��ʾ�� False---����ʾ
'11       SShowPoints-----���˵�λ������������ַ���ʱ�Ƿ���ʾʡ�Ժţ�True---��ʾ�� False---����ʾ
'12       ShowDateType ------��ʾ�������ڵ���ʽ��0Ϊ����ʾ��1Ϊ��ʾ�����գ�2Ϊֻ��ʾ���գ�3Ϊ�ԡ���-�ա���ʽ��ʾ���ա�
'13       ShowPositionID -----------�Ƿ���ʾְλID��0Ϊ����ʾ�� 1Ϊ��ʾ
'14       ShowPositionName -----------�Ƿ���ʾְλ���ƣ� 0Ϊ����ʾ�� 1Ϊ��ʾ
'15       ShowWorkPlaceName -----------�Ƿ���ʾ�����ص㣬 0Ϊ����ʾ�� 1Ϊ��ʾ
'16       ShowSubCompanyName -----------�Ƿ���ʾ���˵�λ�� 0Ϊ����ʾ�� 1Ϊ��ʾ
'17       ShowPositionNum -----------�Ƿ���ʾ��Ƹ������ 0Ϊ����ʾ�� 1Ϊ��ʾ
'18       ShowPositionStatus -----------�Ƿ�ְλ״̬�� 0Ϊ����ʾ�� 1Ϊ��ʾ
'19       ShowValidDate -----------�Ƿ���ʾ��Ч�ڣ� 0Ϊ����ʾ�� 1Ϊ��ʾ
'20       ShowUrgentSign -----------�Ƿ���ʾ������Ƹ��־��TrueΪ��ʾ��FalseΪ����ʾ
'21       ShowNewSign -------�Ƿ���ʾ����Ƹ��־��TrueΪ��ʾ��FalseΪ����ʾ
'22       UsePage ----------�Ƿ��ҳ��ʾ��TrueΪ��ҳ��ʾ��FalseΪ����ҳ��ʾ
'23       OpenType -----����ְλ�򿪷�ʽ��0Ϊ��ԭ���ڴ򿪣�1Ϊ���´��ڴ�
'=================================================

Private Function GetPositionList(PositionNum, IsUrgent, DateNum, OrderType, ShowType, TitleLen, WorkPlaceNameLen, SubCompanyNameLen, PShowPoints, WShowPoints, SShowPoints, ShowDateType, ShowPositionID, ShowPositionName, ShowWorkPlaceName, ShowSubCompanyName, ShowPositionNum, ShowPositionStatus, ShowValidDate, ShowUrgentSign, ShowNewSign, UsePage, OpenType)

    Dim sqlPosition, rsPositionList, iCount, strPositionList, TitleStrstrLink, TitleStr, strLink
    Dim iTop, iElite, iCommon, iHot, iNew, iTitle1, iTitle2
    iCount = 0

    If TitleLen < 0 Or TitleLen > 200 Then
        TitleLen = 50
    End If
    
    If PositionNum > 0 Then
        sqlPosition = "select top " & PositionNum & " "
    Else
        sqlPosition = "select "
    End If
    sqlPosition = sqlPosition & "P.PositionID,P.PositionName,W.WorkPlaceName,P.ReleaseDate,P.PositionNum,P.ValidDate,P.PositionStatus,P.SubCompanyName from PE_Position P left join PE_WorkPlace W on P.WorkPlaceID=W.WorkPlaceID"
    
    If IsUrgent = True Then
        sqlPosition = sqlPosition & " where P.Urgent=0"
        If DateNum > 0 Then
            sqlPosition = sqlPosition & " and DateDiff(" & PE_DatePart_D & ",P.ReleaseDate," & PE_Now & ")<" & DateNum & " and P.PositionStatus=0"
        Else
            sqlPosition = sqlPosition & " and P.PositionStatus=0"
        End If
    Else
        If DateNum > 0 Then
            sqlPosition = sqlPosition & " where DateDiff(" & PE_DatePart_D & ",P.ReleaseDate," & PE_Now & ")<" & DateNum & " and P.PositionStatus=0"
        Else
            sqlPosition = sqlPosition & " where P.PositionStatus=0"
        End If
    End If
    sqlPosition = sqlPosition & " order by "
    Select Case PE_CLng(OrderType)
    Case 1
        sqlPosition = sqlPosition & "P.PositionID desc"
    Case 2
        sqlPosition = sqlPosition & "P.PositionID asc"
    Case 3
        sqlPosition = sqlPosition & "P.ReleaseDate desc,P.PositionID desc"
    Case 4
        sqlPosition = sqlPosition & "P.ReleaseDate asc,P.PositionID desc"
    Case Else
        sqlPosition = sqlPosition & "P.PositionID desc"
    End Select
    Set rsPositionList = Server.CreateObject("ADODB.Recordset")
    rsPositionList.Open sqlPosition, Conn, 1, 1
    If rsPositionList.BOF And rsPositionList.EOF Then
        If UsePage = True Then totalPut = 0
        If IsUrgent = True Then
            strPositionList = "<li>" & XmlText_Class("PositionList/t1", "û��") & XmlText_Class("PositionList/t1", "����") & "��Ƹ��Ϣ</li>"
        Else
            If DateNum > 0 Then
                strPositionList = "<li>" & XmlText_Class("PositionList/t1", "û��") & XmlText_Class("PositionList/t1", "���") & DateNum & "����Ƹ��Ϣ</li>"
            Else
                strPositionList = "<li>" & XmlText_Class("PositionList/t1", "û��") & XmlText_Class("PositionList/t1", "�κ�") & "��Ƹ��Ϣ</li>"
            End If
        End If
        rsPositionList.Close
        Set rsPositionList = Nothing
        GetPositionList = strPositionList
        Exit Function
    End If
    If UsePage = True Then
        totalPut = rsPositionList.RecordCount
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
                rsPositionList.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
    End If
    Dim iPositionIDLen, iPositionNameLen, iWorkPlaceNameLen, iSubCompanyNameLen, iPositionNumLen, iPositionStatusLen, iValidDateLen, iReleaseDateLen
    iPositionIDLen = R_XmlText_Class("ShowPosition/PositionIDLen", "60")
    iPositionNameLen = R_XmlText_Class("ShowPosition/PositionNameLen", "80")
    iWorkPlaceNameLen = R_XmlText_Class("ShowPosition/WorkPlaceNameLen", "80")
    iSubCompanyNameLen = R_XmlText_Class("ShowPosition/SubCompanyNameLen", "120")
    iPositionNumLen = R_XmlText_Class("ShowPosition/PositionNumLen", "60")
    iPositionStatusLen = R_XmlText_Class("ShowPosition/PositionStatusLen", "60")
    iValidDateLen = R_XmlText_Class("ShowPosition/ValidDateLen", "60")
    iReleaseDateLen = R_XmlText_Class("ShowPosition/ReleaseDateLen", "80")
    

    If ShowType = 3 Then
        strPositionList = "<table width='100%' cellpadding='0' cellspacing='0'>"
        strPositionList = strPositionList & "<tr'>"
        If ShowPositionID > 0 Then
            strPositionList = strPositionList & "<td width='" & iPositionIDLen & "' align='center'>" & R_XmlText_Class("ShowPosition/PositionID", "���") & "</td>"
        End If
        If ShowPositionName > 0 Then
            strPositionList = strPositionList & "<td width='" & iPositionNameLen & "' align='center'>" & R_XmlText_Class("ShowPosition/PositionName", "ְλ����") & "</td>"
        End If
        If ShowWorkPlaceName > 0 Then
            strPositionList = strPositionList & "<td width='" & iWorkPlaceNameLen & "' align='center'>" & R_XmlText_Class("ShowPosition/WorkPlaceName", "�����ص�") & "</td>"
        End If
        If ShowSubCompanyName > 0 Then
            strPositionList = strPositionList & "<td width='" & iSubCompanyNameLen & "' align='center'>" & R_XmlText_Class("ShowPosition/SubCompanyName", "���˵�λ") & "</td>"
        End If
        If ShowPositionNum > 0 Then
            strPositionList = strPositionList & "<td width='" & iPositionNumLen & "' align='center'>" & R_XmlText_Class("ShowPosition/PositionNum", "��Ƹ����") & "</td>"
        End If
        If ShowPositionStatus > 0 Then
            strPositionList = strPositionList & "<td width='" & iPositionStatusLen & "' align='center'>" & R_XmlText_Class("ShowPosition/PositionStatus", "ְλ״̬") & "</td>"
        End If
        If ShowValidDate > 0 Then
            strPositionList = strPositionList & "<td width='" & iValidDateLen & "' align='center'>" & R_XmlText_Class("ShowPosition/ValidDate", "��Ч��") & "</td>"
        End If
        If ShowDateType > 0 Then
            strPositionList = strPositionList & "<td width='" & iReleaseDateLen & "' align='center'>" & R_XmlText_Class("ShowPosition/ReleaseDate", "��������") & "</td>"
        End If
        strPositionList = strPositionList & "</tr>"
    Else
        strPositionList = ""
    End If
        
    Do While Not rsPositionList.EOF
        
        If TitleLen > 0 Then
            TitleStr = GetSubStr(rsPositionList("PositionName"), TitleLen, PShowPoints)
        Else
            TitleStr = rsPositionList("PositionName")
        End If
        
        strLink = "<a href='SupplyInfo.asp?PositionID=" & rsPositionList("PositionID") & "'"
        If OpenType = 0 Then
            strLink = strLink & " target='_self'>"
        Else
            strLink = strLink & " target='_blank'>"
        End If
        strLink = strLink & TitleStr & "</a>"

        If ShowType = 1 Then
            strPositionList = strPositionList & "&nbsp;" & strLink
            If ShowWorkPlaceName > 0 Then
                If WorkPlaceNameLen > 0 Then
                    strPositionList = strPositionList & "&nbsp;" & GetSubStr(rsPositionList("WorkPlaceName"), WorkPlaceNameLen, WShowPoints)
                Else
                    strPositionList = strPositionList & "&nbsp;" & rsPositionList("WorkPlaceName")
                End If
            End If
            If ShowSubCompanyName > 0 Then
                If SubCompanyNameLen > 0 Then
                    strPositionList = strPositionList & "&nbsp;" & GetSubStr(rsPositionList("SubCompanyName"), SubCompanyNameLen, SShowPoints)
                Else
                    strPositionList = strPositionList & "&nbsp;" & rsPositionList("SubCompanyName")
                End If
            End If
            If ShowPositionNum > 0 Then
                strPositionList = strPositionList & "&nbsp;" & rsPositionList("PositionNum")
            End If
            If ShowPositionStatus > 0 Then
                strPositionList = strPositionList & "&nbsp;" & GetPositionStatus(rsPositionList("PositionStatus"), rsPositionList("ReleaseDate"), rsPositionList("ValidDate"))
            End If
            If ShowValidDate > 0 Then
                strPositionList = strPositionList & "&nbsp;" & rsPositionList("ValidDate")
            End If
            If ShowUrgentSign = True Then
                strPositionList = strPositionList & "<img src='" & strInstallDir & "images/Urgent.gif' >"
            End If
            If ShowDateType > 0 Then
                strPositionList = strPositionList & "&nbsp;("
                strPositionList = strPositionList & GetUpdateTimeStr(rsPositionList("ReleaseDate"), ShowDateType)
                strPositionList = strPositionList & ")"
            End If
            strPositionList = strPositionList & "<br>"
        ElseIf ShowType = 2 Then
            strPositionList = strPositionList & "&nbsp;" & strLink
            If ShowWorkPlaceName > 0 Then
                If WorkPlaceNameLen > 0 Then
                    strPositionList = strPositionList & "&nbsp;" & GetSubStr(rsPositionList("WorkPlaceName"), WorkPlaceNameLen, WShowPoints)
                Else
                    strPositionList = strPositionList & "&nbsp;" & rsPositionList("WorkPlaceName")
                End If
            End If
            If ShowSubCompanyName > 0 Then
                If SubCompanyNameLen > 0 Then
                    strPositionList = strPositionList & "&nbsp;" & GetSubStr(rsPositionList("SubCompanyName"), SubCompanyNameLen, SShowPoints)
                Else
                    strPositionList = strPositionList & "&nbsp;" & rsPositionList("SubCompanyName")
                End If
            End If
            If ShowPositionNum > 0 Then
                strPositionList = strPositionList & "&nbsp;" & rsPositionList("PositionNum")
            End If
            If ShowPositionStatus > 0 Then
                strPositionList = strPositionList & "&nbsp;" & GetPositionStatus(rsPositionList("PositionStatus"), rsPositionList("ReleaseDate"), rsPositionList("ValidDate"))
            End If
            If ShowValidDate > 0 Then
                strPositionList = strPositionList & "&nbsp;" & rsPositionList("ValidDate")
            End If
            If ShowNewSign = True Then
                strPositionList = strPositionList & "<img src='" & strInstallDir & "images/j_New.gif' >"
            End If
            If ShowDateType > 0 Then
                strPositionList = strPositionList & "&nbsp;("
                strPositionList = strPositionList & GetUpdateTimeStr(rsPositionList("ReleaseDate"), ShowDateType)
                strPositionList = strPositionList & ")"
            End If
            strPositionList = strPositionList & "<br>"
         ElseIf ShowType = 3 Then
            strPositionList = strPositionList & "<tr class='listbg'>"
            If ShowPositionID > 0 Then
                strPositionList = strPositionList & "<td width='" & iPositionIDLen & "' align='center'>" & rsPositionList("PositionID") & "</td>"
            End If
            If ShowPositionName > 0 Then
                strPositionList = strPositionList & "<td width='" & iPositionNameLen & "' align='center'>" & strLink & "</td>"
            End If
            If ShowWorkPlaceName > 0 Then
                If WorkPlaceNameLen > 0 Then
                    strPositionList = strPositionList & "<td width='" & iWorkPlaceNameLen & "' align='center'>" & GetSubStr(rsPositionList("WorkPlaceName"), WorkPlaceNameLen, WShowPoints) & "</td>"
                Else
                    strPositionList = strPositionList & "<td width='" & iWorkPlaceNameLen & "' align='center'>" & rsPositionList("WorkPlaceName") & "</td>"
                End If
            End If
            If ShowSubCompanyName > 0 Then
                If SubCompanyNameLen > 0 Then
                    strPositionList = strPositionList & "<td width='" & iSubCompanyNameLen & "' align='center'>" & GetSubStr(rsPositionList("SubCompanyName"), SubCompanyNameLen, SShowPoints) & "</td>"
                Else
                    strPositionList = strPositionList & "<td width='" & iSubCompanyNameLen & "' align='center'>" & rsPositionList("SubCompanyName") & "</td>"
                End If
            End If
            If ShowPositionNum > 0 Then
                strPositionList = strPositionList & "<td width='" & iPositionNumLen & "' align='center'>" & rsPositionList("PositionNum") & "</td>"
            End If
            If ShowPositionStatus > 0 Then
                strPositionList = strPositionList & "<td width='" & iPositionStatusLen & "' align='center'>" & GetPositionStatus(rsPositionList("PositionStatus"), rsPositionList("ReleaseDate"), rsPositionList("ValidDate")) & "</td>"
            End If
            If ShowValidDate > 0 Then
                strPositionList = strPositionList & "<td width='" & iValidDateLen & "' align='center'>" & rsPositionList("ValidDate") & "</td>"
            End If
            If ShowDateType > 0 Then
                strPositionList = strPositionList & "<td width='" & iReleaseDateLen & "' align='center'>" & GetUpdateTimeStr(rsPositionList("ReleaseDate"), ShowDateType) & "</td>"
            End If
            strPositionList = strPositionList & "</tr>"
        End If
        rsPositionList.MoveNext
        iCount = iCount + 1
        If UsePage = True And iCount >= MaxPerPage Then Exit Do
    Loop
    If ShowType = 3 Then
        strPositionList = strPositionList & "</table>"
    End If
    rsPositionList.Close
    Set rsPositionList = Nothing
    GetPositionList = strPositionList
End Function


Private Function GetUpdateTimeStr(UpdateTime, ShowDateType)
    Dim strUpdateTime
    If Not IsDate(UpdateTime) Then
        GetUpdateTimeStr = ""
        Exit Function
    End If
    Select Case ShowDateType
    Case 1
        strUpdateTime = Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2)
    Case 2
        strUpdateTime = Month(UpdateTime) & strMonth & Day(UpdateTime) & strDay
    Case 3
        strUpdateTime = Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2)
    Case 4
        strUpdateTime = Year(UpdateTime) & strYear & Month(UpdateTime) & strMonth & Day(UpdateTime) & strDay
    Case 5
        strUpdateTime = UpdateTime
    Case 6
        strUpdateTime = UpdateTime
    End Select
    If DateDiff("D", UpdateTime, Now()) < DaysOfNew Then
        strUpdateTime = "<font " & strListStr_Font & ">" & strUpdateTime & "</font>"
    End If
    GetUpdateTimeStr = strUpdateTime
End Function




Private Function GetPositionStatus(PositionStatus, ReleaseDate, ValidDate)
    Dim MyPositionStatus, strPositionStatus
    Dim CurrentDate, MyReleaseDate, MyValidDate
    MyPositionStatus = PE_CLng(PositionStatus)
    MyReleaseDate = ReleaseDate
    MyValidDate = PE_CLng(ValidDate)
    If MyReleaseDate <> "" And IsDate(MyReleaseDate) = True Then
        MyReleaseDate = CDate(MyReleaseDate)
    Else
        MyReleaseDate = PE_Now
    End If
    CurrentDate = DateAdd("d", 0, Date)
    If DateDiff("d", MyReleaseDate, CurrentDate) <= ValidDate Then
        If MyPositionStatus = 0 Then
            strPositionStatus = "������Ƹ��"
        Else
            If MyPositionStatus = 1 Then
                strPositionStatus = "��ֹͣ��Ƹ"
            End If
        End If
    Else
        strPositionStatus = "�ѹ���Ч��"
    End If
    GetPositionStatus = strPositionStatus
End Function







Private Function GetPositionStatus_Search(PositionStatus, ReleaseDate, ValidDate)
    Dim MyPositionStatus, strPositionStatus
    Dim CurrentDate, MyReleaseDate, MyValidDate

    MyPositionStatus = PE_CLng(PositionStatus)
    MyReleaseDate = ReleaseDate
    MyValidDate = PE_CLng(ValidDate)
    If MyReleaseDate <> "" And IsDate(MyReleaseDate) = True Then
        MyReleaseDate = CDate(MyReleaseDate)
    Else
        MyReleaseDate = PE_Now
    End If
    CurrentDate = DateAdd("d", 0, Date)
    If DateDiff("d", MyReleaseDate, CurrentDate) <= ValidDate Then
        If MyPositionStatus = 0 Then
            strPositionStatus = "������Ƹ��"
        ElseIf MyPositionStatus = 1 Then
            strPositionStatus = "��ֹͣ��Ƹ"
        End If
    Else
        strPositionStatus = "�ѹ���Ч��"
    End If
    GetPositionStatus_Search = strPositionStatus
End Function




'=================================================
'��������GetCorrelativePosition
'��  �ã���ʾ�������ְλ
'��  ����
'0        PositionNum ----�����ʾ���ٸ����ְλ��Ϣ��0Ϊ���е����ְλ
'1        OrderType ----����ʽ��1----��ְλID����2----��ְλID����3----��������ʱ�併��4----������ʱ������
'2        TitleLen   ----ְλ��������ַ�����һ������=����Ӣ���ַ�����Ϊ0������ʾ����ְλ��
'3        ShowDateType ------��ʾ�������ڵ���ʽ��0Ϊ����ʾ��1Ϊ��ʾ�����գ�2Ϊֻ��ʾ���գ�3Ϊ�ԡ���-�ա���ʽ��ʾ���ա�

'4        Cols       ----ÿ�е������������������ͻ��С�
'5        OpenType -----����ְλ�򿪷�ʽ��0Ϊ��ԭ���ڴ򿪣�1Ϊ���´��ڴ�

'=================================================


Private Function GetCorrelativePosition(PositionNum, OrderType, TitleLen, ShowDateType, Cols, OpenType, PositionID, PositionKeyword)
    Dim rsCorrelative, sqlCorrelative
    Dim TitleStr, strLink, iTemp, iCols, strCorrelativePosition, strKey, arrKey, i
    iTemp = 1
    If PE_CLng(Cols) <> 0 Then
        iCols = PE_CLng(Cols)
    Else
        iCols = 1
    End If
    strCorrelativePosition = strCorrelativePosition & "  <p align='center'>"
    strKey = ReplaceBadChar(PositionKeyword)
    If InStr(strKey, "|") > 0 Then
        arrKey = Split(strKey, "|")
        strKey = "((PositionKeyword like '%" & arrKey(0) & "|%')"
        For i = 1 To UBound(arrKey)
            strKey = strKey & " or (PositionKeyword like '%|" & arrKey(i) & "|%')"
        Next
        strKey = strKey & ")"
    Else
        strKey = "(PositionKeyword like '%" & strKey & "%')"
    End If

    If TitleLen < 0 Or TitleLen > 200 Then
        TitleLen = 50
    End If

    If PE_CLng(PositionNum) > 0 Then
        sqlCorrelative = "select top " & PE_CLng(PositionNum)
    Else
        sqlCorrelative = "select "
    End If
    sqlCorrelative = sqlCorrelative & " PositionID,PositionName,ReleaseDate from PE_Position where "
    sqlCorrelative = sqlCorrelative & strKey & " and PositionID<>" & PE_CLng(PositionID)
    sqlCorrelative = sqlCorrelative & " and PositionStatus=0"
    sqlCorrelative = sqlCorrelative & " order by "
    Select Case PE_CLng(OrderType)
    Case 1
        sqlCorrelative = sqlCorrelative & "PositionID desc"
    Case 2
        sqlCorrelative = sqlCorrelative & "PositionID asc"
    Case 3
        sqlCorrelative = sqlCorrelative & "ReleaseDate desc,PositionID desc"
    Case 4
        sqlCorrelative = sqlCorrelative & "ReleaseDate asc,PositionID asc"
    Case Else
        sqlCorrelative = sqlCorrelative & "PositionID desc"
    End Select
    Set rsCorrelative = Server.CreateObject("ADODB.RecordSet")
    rsCorrelative.Open sqlCorrelative, Conn, 1, 3
    Do While Not rsCorrelative.EOF
        If TitleLen > 0 Then
            TitleStr = GetSubStr(rsCorrelative("PositionName"), TitleLen, ShowSuspensionPoints)
        Else
            TitleStr = rsCorrelative("PositionName")
        End If
        
        strLink = "<a href='SupplyInfo.asp?PositionID=" & rsCorrelative("PositionID") & "'"
        If OpenType = 0 Then
            strLink = strLink & " target='_self'>"
        Else
            strLink = strLink & " target='_blank'>"
        End If
        strLink = strLink & TitleStr & "</a>"
        strCorrelativePosition = strCorrelativePosition & strLink
        If (iTemp Mod iCols) = 0 Then
            If ShowDateType > 0 Then
                strCorrelativePosition = strCorrelativePosition & "&nbsp;&nbsp;"
                strCorrelativePosition = strCorrelativePosition & GetUpdateTimeStr(rsCorrelative("ReleaseDate"), ShowDateType)
            End If
            strCorrelativePosition = strCorrelativePosition & "<br>"
        Else
            If ShowDateType > 0 Then
                strCorrelativePosition = strCorrelativePosition & "&nbsp;&nbsp;"
                strCorrelativePosition = strCorrelativePosition & GetUpdateTimeStr(rsCorrelative("ReleaseDate"), ShowDateType)
            End If
            strCorrelativePosition = strCorrelativePosition & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        End If
        rsCorrelative.MoveNext
        iTemp = iTemp + 1
    Loop
    rsCorrelative.Close
    Set rsCorrelative = Nothing
    strCorrelativePosition = strCorrelativePosition & "  </p>"
    GetCorrelativePosition = strCorrelativePosition
End Function



Private Sub SaveSupplyInfo()
    Dim PositionSupplyInfoID, PositionID, SubCompanyID, WorkPlaceID, mrs, MaxPositionSupplyInfoID
    Dim PositionName, SubCompanyName, WorkPlaceName, PositionNum, ValidDate, ReleaseDate, CurrentDate
    Dim rsSupply, sqlSupply
    Dim rsResume, sqlResume
    Dim rsPosition, sqlPosition

    PositionID = Trim(Request("PositionID"))
    SubCompanyID = Trim(Request("SubCompanyID"))
    WorkPlaceID = Trim(Request("WorkPlaceID"))
    PositionName = Request("PositionName")
    SubCompanyName = ReplaceBadChar(Trim(Request("SubCompanyName")))
    WorkPlaceName = ReplaceBadChar(Trim(Request("WorkPlaceName")))
    PositionNum = PE_CLng(Trim(Request("PositionNum")))
    ValidDate = PE_CLng(Trim(Request("ValidDate")))
    ReleaseDate = Trim(Request("ReleaseDate"))

    '���ж��Ƿ��Ѿ���¼
    If CheckUserLogined() = False Then
        Response.Redirect "" & strInstallDir & "User/User_Login.asp"
        Exit Sub
    End If
    '�жϸ�ְλ�Ƿ��ѹ���Ч��
    If ReleaseDate <> "" And IsDate(ReleaseDate) = True Then
        ReleaseDate = CDate(ReleaseDate)
    Else
        ReleaseDate = PE_Now
    End If
    CurrentDate = DateAdd("d", 0, Date)
    If DateDiff("d", ReleaseDate, CurrentDate) > ValidDate Then
        Response.Write "<html>"
        Response.Write "<head>"
        Response.Write "<title>ְλ����</title>"
        Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
        Response.Write "<link href='../Admin/Admin_Style.css' rel='stylesheet' type='text/css'>"
        Response.Write "</head>"
        Response.Write "<body>"
        Response.Write "<br><br>"
        Response.Write "<table class='border' align=center width='400' border='0' cellpadding='0' cellspacing='0' bordercolor='#999999'>"
        Response.Write "  <tr align=center> "
        Response.Write "    <td  height='22' align='center' class='title'> "
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr>"
        Response.Write "    <td><table width='100%' border='0' cellpadding='2' cellspacing='1'>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='400' height='88' align='center'><font color=red>�Բ������������ְλ�ѹ���Ч�ڣ����Բ��������ְλ��</font></td>"
        Response.Write "        </tr>"
        Response.Write "      </table></td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td height='30' align='center'>"
        Response.Write "��<a href='javascript:window.close();'>�رմ���</a>��"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "</table>" & vbCrLf
        Response.Write "</body>"
        Response.Write "</html>"
        Exit Sub
    End If


    '�жϵ�¼�û��Ƿ��Ѿ���д����
    Set rsResume = Server.CreateObject("Adodb.RecordSet")
    sqlResume = "select ResumeID from PE_Resume where UserName='" & UserName & "'"
    rsResume.Open sqlResume, Conn, 1, 3
    If rsResume.BOF And rsResume.EOF Then
        Response.Redirect "" & strInstallDir & "User/User_Job.asp?Action=Resume"
    End If
    rsResume.Close
    Set rsResume = Nothing


    If PositionID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ְλ����Ϊ�գ�</li>"
    Else
        PositionID = PE_CLng(PositionID)
    End If
    If SubCompanyID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���˵�λ����Ϊ�գ�</li>"
    Else
        SubCompanyID = PE_CLng(SubCompanyID)
    End If
    If WorkPlaceID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����ص㲻��Ϊ�գ�</li>"
    Else
        WorkPlaceID = PE_CLng(WorkPlaceID)
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    '�жϵ�¼�û��Ƿ��Ѿ��������ְλ
    Set rsPosition = Server.CreateObject("Adodb.RecordSet")
    sqlPosition = "select PositionSupplyInfoID from PE_PositionSupplyInfo where PositionID=" & PositionID & "And UserName='" & UserName & "'"
    rsPosition.Open sqlPosition, Conn, 1, 3
    If rsPosition.BOF And rsPosition.EOF Then
        Set rsSupply = Server.CreateObject("Adodb.RecordSet")
        Set mrs = Conn.Execute("select max(PositionSupplyInfoID) from PE_PositionSupplyInfo")
        If IsNull(mrs(0)) Then
            MaxPositionSupplyInfoID = 0
        Else
            MaxPositionSupplyInfoID = mrs(0)
        End If
        Set mrs = Nothing
        sqlSupply = "select Top 1 * from PE_PositionSupplyInfo"
        rsSupply.Open sqlSupply, Conn, 1, 3
        rsSupply.addnew
        rsSupply("PositionSupplyInfoID") = MaxPositionSupplyInfoID + 1
        rsSupply("UserName") = UserName
        rsSupply("PositionID") = PositionID
        rsSupply("SubCompanyID") = SubCompanyID
        rsSupply("WorkPlaceID") = WorkPlaceID
        rsSupply("CheckStatus") = 0
        rsSupply("SupplyDate") = DateAdd("d", 0, Date)
        rsSupply.Update
        rsSupply.Close
        Set rsSupply = Nothing

        Response.Write "<html>"
        Response.Write "<head>"
        Response.Write "<title>ְλ����</title>"
        Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
        Response.Write "<link href='../Admin/Admin_Style.css' rel='stylesheet' type='text/css'>"
        Response.Write "</head>"
        Response.Write "<body>"
        Response.Write "<br><br>"
        Response.Write "<table class='border' align=center width='400' border='0' cellpadding='0' cellspacing='0' bordercolor='#999999'>"
        Response.Write "  <tr align=center> "
        Response.Write "    <td  height='22' align='center' class='title'> "
        Response.Write "<b>" & UserName & "--���Ѿ��ɹ������ְλ��</b>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr>"
        Response.Write "    <td><table width='100%' border='0' cellpadding='2' cellspacing='1'>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='100' align='right'><strong>ְλ���ƣ�</strong></td>"
        Response.Write "          <td>" & PositionName & "</td>"
        Response.Write "        </tr>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='100' align='right'><strong>������λ��</strong></td>"
        Response.Write "          <td>" & SubCompanyName & "</td>"
        Response.Write "        </tr>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='100' align='right'><strong>�����ص㣺</strong></td>"
        Response.Write "          <td>" & WorkPlaceName & "</td>"
        Response.Write "        </tr>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='100' align='right'><strong>��Ƹ������</strong></td>"
        Response.Write "          <td>" & PositionNum & "</td>"
        Response.Write "        </tr>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='100' align='right'><strong>��Ч�ڣ�</strong></td>"
        Response.Write "          <td>" & ValidDate & "</td>"
        Response.Write "        </tr>"
        Response.Write "      </table></td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td height='30' align='center'>"
        Response.Write "��<a href='javascript:window.close();'>�رմ���</a>��"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "</table>" & vbCrLf
        Response.Write "</body>"
        Response.Write "</html>"
    Else
        Response.Write "<html>"
        Response.Write "<head>"
        Response.Write "<title>ְλ����</title>"
        Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
        Response.Write "<link href='../Admin/Admin_Style.css' rel='stylesheet' type='text/css'>"
        Response.Write "</head>"
        Response.Write "<body>"
        Response.Write "<br><br>"
        Response.Write "<table class='border' align=center width='400' border='0' cellpadding='0' cellspacing='0' bordercolor='#999999'>"
        Response.Write "  <tr align=center> "
        Response.Write "    <td  height='22' align='center' class='title'> "
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr>"
        Response.Write "    <td><table width='100%' border='0' cellpadding='2' cellspacing='1'>"
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='400' height='88' align='center'><font color=red>���Ѿ������˸�ְλ���벻Ҫ�ظ�����ͬһְλ��</font></td>"
        Response.Write "        </tr>"
        Response.Write "      </table></td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td height='30' align='center'>"
        Response.Write "��<a href='javascript:window.close();'>�رմ���</a>��"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "</table>" & vbCrLf
        Response.Write "</body>"
        Response.Write "</html>"
    End If
    rsPosition.Close
    Set rsPosition = Nothing
End Sub







Public Function GetListFromTemplate(ByVal strValue)
    Dim strList
    strList = strValue
    regEx.Pattern = "\{\$GetPositionList\((.*?)\)\}"
    Set Matches = regEx.Execute(strList)
    For Each Match In Matches
        strList = PE_Replace(strList, Match.value, GetListFromLabel(Match.SubMatches(0)))
    Next
    GetListFromTemplate = strList
End Function



Private Function GetListFromLabel(ByVal str1)
    Dim strTemp, arrTemp
    Dim tPositionNum, tDateNum, tOrderType, tShowType, tTitleLen, tShowDateType
    If str1 = "" Then
        GetListFromLabel = ""
        Exit Function
    End If
    
    strTemp = Replace(str1, Chr(34), "")
    arrTemp = Split(strTemp, ",")
    If UBound(arrTemp) <> 22 Then
        GetListFromLabel = "����ʽ��ǩ��{$GetPositionList(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
        Exit Function
    End If
    Select Case Trim(arrTemp(0))
    Case "PositionNum"
        tPositionNum = 8
    Case Else
        tPositionNum = PE_CLng(arrTemp(0))
    End Select

    GetListFromLabel = GetPositionList(PE_CLng(arrTemp(0)), PE_CBool(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)), PE_CLng(arrTemp(5)), PE_CLng(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CBool(arrTemp(8)), PE_CBool(arrTemp(9)), PE_CBool(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), PE_CLng(arrTemp(13)), PE_CLng(arrTemp(14)), PE_CLng(arrTemp(15)), PE_CLng(arrTemp(16)), PE_CLng(arrTemp(17)), PE_CLng(arrTemp(18)), PE_CBool(arrTemp(19)), PE_CBool(arrTemp(20)), PE_CBool(arrTemp(21)), PE_CLng(arrTemp(22)))
End Function



Private Sub ReplaceCommon()
    Call ReplaceCommonLabel
    
    strHtml = PE_Replace(strHtml, "{$MenuJS}", GetMenuJS(ChannelDir, ShowClassTreeGuide))
    strHtml = PE_Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(SkinID))
    
End Sub

Public Sub GetHtml_index()
    Dim GetPositionName, GetWorkPlaceName, GetPositionNum, GetReleaseDate, GetValidDate, GetSubCompanyName, GetContacter, GetTelephone, GetAddress, GetE_mail, GetPositionDescription, GetDutyRequest, GetStatus, GetSaveSupply
    Dim PositionList_Content, PositionList_Content2, iPositionID, iMaxPerPageNum, iMaxPerPage
    Dim rsPosition, sqlPosition, strTemp
    Dim strPositionKeyword, iCount
    Dim strCorrelativePosition, arrTemp
    Dim PositionListShowPage, iPerPageNum

    iCount = 0
    strPageTitle = ""
    PageTitle = "��ҳ"

    strFileName = ChannelUrl & "/Index.asp"
    strPageTitle = SiteTitle
    strPageTitle = strPageTitle & " >> " & ChannelName & " >> " & PageTitle
    strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;<a href='" & strInstallDir & "Job/Index.asp'>" & ChannelName & "</a>&nbsp;" & strNavLink & "&nbsp;" & PageTitle

    Call ReplaceCommonLabel
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)

    '�õ�ְλ��Ϣ�б�İ�����Ƶ�HTML����
    regEx.Pattern = "��PositionList_Content��([\s\S]*?)��\/PositionList_Content��"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        PositionList_Content = Match.value
    Next
    strHtml = regEx.Replace(strHtml, "{$PositionList_Content}")
    
    PositionList_Content = Replace(PositionList_Content, "��PositionList_Content��", "")
    PositionList_Content = Replace(PositionList_Content, "��/PositionList_Content��", "")
    PositionList_Content2 = ""
    
    '�õ�ÿ����ʾ������
    regEx.Pattern = "��PerPageNum=[1-9]��"
    Set Matches = regEx.Execute(PositionList_Content)
    PositionList_Content = regEx.Replace(PositionList_Content, "")
    For Each Match In Matches
        iPerPageNum = Match.value
    Next
    If iPerPageNum = "" Then
        iPerPageNum = 1
    Else
        iPerPageNum = Replace(Replace(iPerPageNum, "��PerPageNum=", ""), "��", "")

        If iPerPageNum = "" Then
            iPerPageNum = 1
        Else
            iPerPageNum = PE_CLng(iPerPageNum)
        End If
        If iPerPageNum = 0 Then iPerPageNum = 1
    End If
    MaxPerPage = iPerPageNum 'ÿҳ��ʾ�ļ�¼��

    '��ʼѭ�����õ�ְλ��Ϣ��HTML����
    sqlPosition = "select P.PositionID,P.PositionName,P.PositionKeyword,W.WorkPlaceID,W.WorkPlaceName,P.PositionNum,P.ReleaseDate,P.PositionStatus,P.ValidDate,S.SubCompanyID,S.SubCompanyName,S.Contacter,S.Telephone,S.Address,S.E_mail,P.PositionDescription,P.DutyRequest from (PE_Position P left join PE_WorkPlace W on P.WorkPlaceID=W.WorkPlaceID) left join PE_SubCompany S on P.SubCompanyID=S.SubCompanyID order by P.PositionID desc"
    Set rsPosition = Server.CreateObject("ADODB.Recordset")
    rsPosition.Open sqlPosition, Conn, 1, 1
    totalPut = rsPosition.RecordCount
    iPositionID = totalPut
    If CurrentPage < 1 Then
        CurrentPage = 1
    End If
    If (CurrentPage - 1) * MaxPerPage > totalPut Then
        If (totalPut Mod iMaxPerPage) = 0 Then
            CurrentPage = totalPut \ MaxPerPage
        Else
            CurrentPage = totalPut \ MaxPerPage + 1
        End If
    End If
    If CurrentPage > 1 Then
        If (CurrentPage - 1) * MaxPerPage < totalPut Then
            rsPosition.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If

    Do While Not rsPosition.EOF
        strPositionKeyword = rsPosition("PositionKeyword")
        strTemp = PositionList_Content
        GetPositionName = "<a href='SupplyInfo.asp?PositionID=" & rsPosition("PositionID") & "'target='_blank'>"
        GetPositionName = GetPositionName & rsPosition("PositionName")
        GetPositionName = GetPositionName & "</a>"
        GetWorkPlaceName = rsPosition("WorkPlaceName")
        GetPositionNum = rsPosition("PositionNum")
        GetReleaseDate = rsPosition("ReleaseDate")
        GetValidDate = rsPosition("ValidDate")
        GetSubCompanyName = rsPosition("SubCompanyName")
        GetContacter = rsPosition("Contacter")
        GetTelephone = rsPosition("Telephone")
        GetAddress = rsPosition("Address")
        GetE_mail = rsPosition("E_mail")
        GetPositionDescription = rsPosition("PositionDescription")
        GetDutyRequest = rsPosition("DutyRequest")
        GetStatus = GetPositionStatus(rsPosition("PositionStatus"), rsPosition("ReleaseDate"), rsPosition("ValidDate"))
        GetSaveSupply = GetSaveSupply & "  <p align='center'>"
        GetSaveSupply = GetSaveSupply & "   <input name='Supply' type='button'  id='Supply' value=' �����ְλ ' onClick=""window.location.href='SupplyInfo.asp?Action=SaveSupplyInfo&PositionID=" & rsPosition("PositionID") & "&SubCompanyID=" & rsPosition("SubCompanyID") & "&WorkPlaceID=" & rsPosition("WorkPlaceID") & "&PositionName=" & rsPosition("PositionName") & "&SubCompanyName=" & rsPosition("SubCompanyName") & "&WorkPlaceName=" & rsPosition("WorkPlaceName") & "&PositionNum=" & rsPosition("PositionNum") & "&ReleaseDate=" & rsPosition("ReleaseDate") & "&ValidDate=" & rsPosition("ValidDate") & "';"" style='cursor:hand;'>&nbsp;&nbsp"
        GetSaveSupply = GetSaveSupply & " </p>"

        strTemp = PE_Replace(strTemp, "{$PositionName}", GetPositionName)
        strTemp = PE_Replace(strTemp, "{$WorkPlaceName}", GetWorkPlaceName)
        strTemp = PE_Replace(strTemp, "{$PositionNum}", GetPositionNum)
        strTemp = PE_Replace(strTemp, "{$ReleaseDate}", GetReleaseDate)
        strTemp = PE_Replace(strTemp, "{$ValidDate}", GetValidDate)
        strTemp = PE_Replace(strTemp, "{$SubCompanyName}", GetSubCompanyName)
        strTemp = PE_Replace(strTemp, "{$Contacter}", GetContacter)
        strTemp = PE_Replace(strTemp, "{$Telephone}", GetTelephone)
        strTemp = PE_Replace(strTemp, "{$Address}", GetAddress)
        strTemp = PE_Replace(strTemp, "{$E_mail}", GetE_mail)
        strTemp = PE_Replace(strTemp, "{$PositionDescription}", GetPositionDescription)
        strTemp = PE_Replace(strTemp, "{$DutyRequest}", GetDutyRequest)
        strTemp = PE_Replace(strTemp, "{$PositionStatus}", GetStatus)
        strTemp = PE_Replace(strTemp, "{$SaveSupply}", GetSaveSupply)


        regEx.Pattern = "\{\$CorrelativePosition\((.*?)\)\}"
        Set Matches = regEx.Execute(strTemp)
        For Each Match In Matches
            arrTemp = Split(Match.SubMatches(0), ",")
            strCorrelativePosition = GetCorrelativePosition(arrTemp(0), arrTemp(1), arrTemp(2), arrTemp(3), arrTemp(4), arrTemp(5), rsPosition("PositionID"), strPositionKeyword)
            strTemp = Replace(strTemp, Match.value, strCorrelativePosition)
        Next

        PositionListShowPage = "<tr><td>"
        If totalPut < MaxPerPage Then
            If iPositionID = 1 Then
                PositionListShowPage = PositionListShowPage & ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False)
            End If
        Else
            If (iCount + 1) Mod MaxPerPage = 0 Then
                PositionListShowPage = PositionListShowPage & ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False)
            Else
                If CurrentPage * MaxPerPage >= totalPut And ((MaxPerPage - 1) - (CurrentPage * MaxPerPage - totalPut)) = iCount Then
                    PositionListShowPage = PositionListShowPage & ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False)
                End If
            End If
        End If
        PositionListShowPage = PositionListShowPage & "</td></tr>"
        strTemp = Replace(strTemp, "{$ShowPage}", PositionListShowPage)
        rsPosition.MoveNext
        iPositionID = iPositionID - 1
        iCount = iCount + 1
        If CurrentPage * MaxPerPage < totalPut And iCount > MaxPerPage Then Exit Do
        PositionList_Content2 = PositionList_Content2 & strTemp
        PositionList_Content2 = PositionList_Content2
    Loop
    rsPosition.Close
    Set rsPosition = Nothing
    strHtml = Replace(strHtml, "{$PositionList_Content}", PositionList_Content2)
End Sub


Public Sub GetHtml_Job()
    strPageTitle = ""
    PageTitle = "��ҳ"
    strFileName = "Index.asp"
    strPageTitle = SiteTitle & " >> " & ChannelName & " >> " & PageTitle
    strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;<a href='" & strInstallDir & "Job/Index.asp'>" & ChannelName & "</a>&nbsp;" & strNavLink & "&nbsp;" & PageTitle

    Call ReplaceCommonLabel
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)
    strHtml = GetListFromTemplate(strHtml)
    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit, False))
End Sub

Public Sub GetHtml_List()
    Call ReplaceCommonLabel
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)
    strHtml = GetSearchResultFromTemplate(strHtml)


    If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit, False))
End Sub

Public Sub SupplyInfo()
    Dim GetPositionName, GetWorkPlaceName, GetPositionNum, GetReleaseDate, GetValidDate, GetSubCompanyName, GetContacter, GetTelephone, GetAddress, GetE_mail, GetPositionDescription, GetDutyRequest, GetStatus, GetSaveSupply, GetWinColse
    Dim rs, sql, strPositionSupplyInfo
    Dim iPositionID, strPositionKeyword
    Dim rsCorrelative, sqlCorrelative
    Dim strKey, arrKey, i, arrTemp
    Dim PositionID, WorkPlaceID, SubCompanyID
    If Action = "SaveSupplyInfo" Then
        Call SaveSupplyInfo
        Exit Sub
    End If
    PositionID = Request("PositionID")
    If PositionID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����ְλID!</li>"
    Else
        PositionID = PE_CLng(PositionID)
    End If
    If FoundErr = True Then
        Response.Write ErrMsg
        Exit Sub
    End If

    sql = "select P.PositionID,P.PositionName,P.PositionKeyword,W.WorkPlaceID,W.WorkPlaceName,P.PositionNum,P.ReleaseDate,P.PositionStatus,P.ValidDate,S.SubCompanyID,S.SubCompanyName,S.Contacter,S.Telephone,S.Address,S.E_mail,P.PositionDescription,P.DutyRequest from (PE_Position P left join PE_WorkPlace W on P.WorkPlaceID=W.WorkPlaceID) left join PE_SubCompany S on P.SubCompanyID=S.SubCompanyID where P.PositionID=" & PositionID & ""
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ְλ���Ѿ���ɾ����</li>"
    End If
    If FoundErr = True Then
        rs.Close
        Set rs = Nothing
        Response.Write ErrMsg
        Exit Sub
    End If
    iPositionID = rs("PositionID")
    strPositionKeyword = rs("PositionKeyword")
    GetPositionName = rs("PositionName")
    GetWorkPlaceName = rs("WorkPlaceName")
    GetPositionNum = rs("PositionNum")
    GetReleaseDate = rs("ReleaseDate")
    GetValidDate = rs("ValidDate")
    GetSubCompanyName = rs("SubCompanyName")
    GetContacter = rs("Contacter")
    GetTelephone = rs("Telephone")
    GetAddress = rs("Address")
    GetE_mail = rs("E_mail")
    GetPositionDescription = rs("PositionDescription")
    GetDutyRequest = rs("DutyRequest")
    GetStatus = GetPositionStatus(rs("PositionStatus"), rs("ReleaseDate"), rs("ValidDate"))
    GetSaveSupply = GetSaveSupply & "   <input name='Supply' type='button'  id='Supply' value=' �����ְλ ' onClick=""window.location.href='SupplyInfo.asp?Action=SaveSupplyInfo&PositionID=" & rs("PositionID") & "&SubCompanyID=" & rs("SubCompanyID") & "&WorkPlaceID=" & rs("WorkPlaceID") & "&PositionName=" & rs("PositionName") & "&SubCompanyName=" & rs("SubCompanyName") & "&WorkPlaceName=" & rs("WorkPlaceName") & "&PositionNum=" & rs("PositionNum") & "&ReleaseDate=" & rs("ReleaseDate") & "&ValidDate=" & rs("ValidDate") & "';"" style='cursor:hand;color:#000000;'>&nbsp;&nbsp"

    strPageTitle = ""
    PageTitle = "ְλ��Ϣ"
    strFileName = ChannelUrl & "/Index.asp"
    strPageTitle = SiteTitle & " >> " & ChannelName & " >> " & PageTitle
    strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;<a href='" & strInstallDir & "Job/Index.asp'>" & ChannelName & "</a>&nbsp;" & strNavLink & "&nbsp;" & PageTitle

    strHtml = Replace(strHtml, "{$PositionID}", PositionID)
    Call ReplaceCommonLabel
    strHtml = Replace(strHtml, "{$PageTitle}", strPageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)

    strHtml = PE_Replace(strHtml, "{$PositionName}", GetPositionName)
    strHtml = PE_Replace(strHtml, "{$WorkPlaceName}", GetWorkPlaceName)
    strHtml = PE_Replace(strHtml, "{$PositionNum}", GetPositionNum)
    strHtml = PE_Replace(strHtml, "{$ReleaseDate}", GetReleaseDate)
    strHtml = PE_Replace(strHtml, "{$ValidDate}", GetValidDate)
    strHtml = PE_Replace(strHtml, "{$SubCompanyName}", GetSubCompanyName)
    strHtml = PE_Replace(strHtml, "{$Contacter}", GetContacter)
    strHtml = PE_Replace(strHtml, "{$Telephone}", GetTelephone)
    strHtml = PE_Replace(strHtml, "{$Address}", GetAddress)
    strHtml = PE_Replace(strHtml, "{$E_mail}", GetE_mail)
    strHtml = PE_Replace(strHtml, "{$PositionDescription}", GetPositionDescription)
    strHtml = PE_Replace(strHtml, "{$DutyRequest}", GetDutyRequest)
    strHtml = PE_Replace(strHtml, "{$PositionStatus}", GetStatus)
    strHtml = PE_Replace(strHtml, "{$SaveSupply}", GetSaveSupply)

    Dim strCorrelativePosition
    regEx.Pattern = "\{\$CorrelativePosition\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        strCorrelativePosition = GetCorrelativePosition(arrTemp(0), arrTemp(1), arrTemp(2), arrTemp(3), arrTemp(4), arrTemp(5), iPositionID, strPositionKeyword)
        strHtml = Replace(strHtml, Match.value, strCorrelativePosition)
    Next
End Sub

Private Function GetSearchResultFromTemplate(ByVal strValue)
    Dim strSearchResult
    strSearchResult = strValue
    regEx.Pattern = "\{\$GetSearchResult\((.*?)\)\}"
    Set Matches = regEx.Execute(strSearchResult)
    For Each Match In Matches
        strSearchResult = PE_Replace(strSearchResult, Match.value, GetSearchResultFromLabel(Match.SubMatches(0)))
    Next
    GetSearchResultFromTemplate = strSearchResult
End Function

Private Function GetSearchResultFromLabel(ByVal str1)
    Dim strTemp, arrTemp
    Dim tPositionNum, tDateNum, tOrderType, tShowType, tTitleLen, tShowDateType
    If str1 = "" Then
        GetSearchResultFromLabel = ""
        Exit Function
    End If
    
    strTemp = Replace(str1, Chr(34), "")
    arrTemp = Split(strTemp, ",")
    If UBound(arrTemp) <> 17 Then
        GetSearchResultFromLabel = "����ʽ��ǩ��{$GetSearchResult(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
        Exit Function
    End If
    GetSearchResultFromLabel = GetSearchResult(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)), PE_CBool(arrTemp(5)), PE_CBool(arrTemp(6)), PE_CBool(arrTemp(7)), PE_CLng(arrTemp(8)), PE_CLng(arrTemp(9)), PE_CLng(arrTemp(10)), PE_CLng(arrTemp(11)), PE_CLng(arrTemp(12)), PE_CLng(arrTemp(13)), PE_CLng(arrTemp(14)), PE_CLng(arrTemp(15)), PE_CBool(arrTemp(16)), PE_CLng(arrTemp(17)))
End Function

'=================================================
'��������GetSearchResult
'��  �ã���ҳ��ʾ�������
'��  ����
'1        ShowNum ----������ʾ��¼����0Ϊ��ʾ���з��������ļ�¼��������0��ʾ���õļ�¼��
'2        OrderType ----����ʽ��1----��ְλID����2----��ְλID����3----������ʱ�併��4----������ʱ������
'3        TitleLen  ----ְλ��������ַ�����һ������=����Ӣ���ַ�����Ϊ0������ʾ����ְλ��
'4        WorkPlaceNameLen----�����ص���������ַ�����һ������=����Ӣ���ַ�����Ϊ0������ʾ��������
'5        SubCompanyNameLen---���˵�λ��������ַ�����һ������=����Ӣ���ַ�����Ϊ0������ʾ��������
'6        PShowPoints-----ְλ������������ַ���ʱ�Ƿ���ʾʡ�Ժţ�True---��ʾ�� False---����ʾ
'7        WShowPoints-----�����ص�������������ַ���ʱ�Ƿ���ʾʡ�Ժţ�True---��ʾ�� False---����ʾ
'8        SShowPoints-----���˵�λ������������ַ���ʱ�Ƿ���ʾʡ�Ժţ�True---��ʾ�� False---����ʾ
'9        ShowDateType ------��ʾ�������ڵ���ʽ��0Ϊ����ʾ��1Ϊ��ʾ�����գ�2Ϊֻ��ʾ���գ�3Ϊ�ԡ���-�ա���ʽ��ʾ���ա�
'10       ShowPositionID -----------�Ƿ���ʾְλID��0Ϊ����ʾ�� 1Ϊ��ʾ
'11       ShowPositionName -----------�Ƿ���ʾְλ���ƣ� 0Ϊ����ʾ�� 1Ϊ��ʾ
'12       ShowWorkPlaceName -----------�Ƿ���ʾ�����ص㣬 0Ϊ����ʾ�� 1Ϊ��ʾ
'13       ShowSubCompanyName -----------�Ƿ���ʾ���˵�λ�� 0Ϊ����ʾ�� 1Ϊ��ʾ
'14       ShowPositionNum -----------�Ƿ���ʾ��Ƹ������ 0Ϊ����ʾ�� 1Ϊ��ʾ
'15       ShowPositionStatus -----------�Ƿ���ʾְλ״̬�� 0Ϊ����ʾ�� 1Ϊ��ʾ
'16       ShowValidDate -----------�Ƿ���ʾ��Ч�ڣ� 0Ϊ����ʾ�� 1Ϊ��ʾ
'17       UsePage -----------�Ƿ��ҳ��ʾ��TrueΪ��ҳ��ʾ��FalseΪ����ҳ��ʾ��ÿҳ��ʾ�����������MaxPerPageָ��
'18       OpenType -----����ְλ�򿪷�ʽ��0Ϊ��ԭ���ڴ򿪣�1Ϊ���´��ڴ�
'=================================================

Private Function GetSearchResult(ShowNum, OrderType, TitleLen, WorkPlaceNameLen, SubCompanyNameLen, PShowPoints, WShowPoints, SShowPoints, ShowDateType, ShowPositionID, ShowPositionName, ShowWorkPlaceName, ShowSubCompanyName, ShowPositionNum, ShowPositionStatus, ShowValidDate, UsePage, OpenType)
    Dim sqlSearch, rsSearch, iCount, arrPositionID, strSearchResult, Content
    Dim SearchJobCategoryID, SearchSubCompanyID, SearchWorkPlaceID, SearchDateNum
    Dim TitleStr, strLink

    SearchJobCategoryID = PE_CLng(Request("SearchJobCategoryID"))
    SearchSubCompanyID = PE_CLng(Request("SearchSubCompanyID"))
    SearchWorkPlaceID = PE_CLng(Request("SearchWorkPlaceID"))
    SearchDateNum = PE_CLng(Request("SearchDateNum"))

    strSearchResult = ""
    If PE_CLng(ShowNum) > 0 Then
        sqlSearch = "select top " & PE_CLng(ShowNum)
    Else
        sqlSearch = "select "
    End If
    sqlSearch = sqlSearch & " P.PositionID,P.ReleaseDate,P.PositionName,P.PositionNum,P.ValidDate,P.PositionStatus,P.SubCompanyName,W.WorkPlaceName from PE_Position P left join PE_WorkPlace W on P.WorkPlaceID=W.WorkPlaceID"
    If Keyword <> "" Then
        sqlSearch = sqlSearch & " where P.PositionName like '%" & Keyword & "%' "
        If SearchJobCategoryID > 0 Then
            sqlSearch = sqlSearch & " and P.JobCategoryID=" & SearchJobCategoryID
        End If
        If SearchSubCompanyID > 0 Then
            sqlSearch = sqlSearch & " and P.SubCompanyID=" & SearchSubCompanyID
        End If
        If SearchWorkPlaceID > 0 Then
            sqlSearch = sqlSearch & " and P.WorkPlaceID=" & SearchWorkPlaceID
        End If
        If SearchDateNum > 0 Then
            sqlSearch = sqlSearch & " and DateDiff(" & PE_DatePart_D & ",P.ReleaseDate," & PE_Now & ")<=" & SearchDateNum
        End If
        sqlSearch = sqlSearch & " and P.PositionStatus=0"
    Else
        If SearchJobCategoryID > 0 Then
            sqlSearch = sqlSearch & " where P.JobCategoryID=" & SearchJobCategoryID
            If SearchSubCompanyID > 0 Then
                sqlSearch = sqlSearch & " and P.SubCompanyID=" & SearchSubCompanyID
            End If
            If SearchWorkPlaceID > 0 Then
                sqlSearch = sqlSearch & " and P.WorkPlaceID=" & SearchWorkPlaceID
            End If
            If SearchDateNum > 0 Then
                sqlSearch = sqlSearch & " and DateDiff(" & PE_DatePart_D & ",P.ReleaseDate," & PE_Now & ")<=" & SearchDateNum
            End If
            sqlSearch = sqlSearch & " and P.PositionStatus=0"
        Else
            If SearchSubCompanyID > 0 Then
                sqlSearch = sqlSearch & " where P.SubCompanyID=" & SearchSubCompanyID
                If SearchWorkPlaceID > 0 Then
                    sqlSearch = sqlSearch & " and P.WorkPlaceID=" & SearchWorkPlaceID
                End If
                If SearchDateNum > 0 Then
                    sqlSearch = sqlSearch & " and DateDiff(" & PE_DatePart_D & ",P.ReleaseDate," & PE_Now & ")<=" & SearchDateNum
                End If
                sqlSearch = sqlSearch & " and P.PositionStatus=0"
            Else
                If SearchWorkPlaceID > 0 Then
                    sqlSearch = sqlSearch & " where P.WorkPlaceID=" & SearchWorkPlaceID
                    If SearchDateNum > 0 Then
                        sqlSearch = sqlSearch & " and DateDiff(" & PE_DatePart_D & ",P.ReleaseDate," & PE_Now & ")<=" & SearchDateNum
                    End If
                    sqlSearch = sqlSearch & " and P.PositionStatus=0"
                Else
                    If SearchDateNum > 0 Then
                        sqlSearch = sqlSearch & " where DateDiff(" & PE_DatePart_D & ",P.ReleaseDate," & PE_Now & ")<=" & SearchDateNum
                        sqlSearch = sqlSearch & " and P.PositionStatus=0"
                    Else
                        sqlSearch = sqlSearch & " where P.PositionStatus=0"
                    End If
                End If
            End If
        End If
    End If
    sqlSearch = sqlSearch & " order by "
    Select Case OrderType
    Case 1
        sqlSearch = sqlSearch & "P.PositionID desc"
    Case 2
        sqlSearch = sqlSearch & "P.PositionID asc"
    Case 3
        sqlSearch = sqlSearch & "P.ReleaseDate desc,P.PositionID desc"
    Case 4
        sqlSearch = sqlSearch & "P.ReleaseDate asc,P.PositionID desc"
    Case Else
        sqlSearch = sqlSearch & "P.PositionID desc"
    End Select

    Set rsSearch = Server.CreateObject("ADODB.Recordset")
    rsSearch.Open sqlSearch, Conn, 1, 1
    If rsSearch.BOF And rsSearch.EOF Then
        If UsePage = True Then totalPut = 0
        strSearchResult = "<p align='center'><br><br>" & R_XmlText_Class("ShowSearch/NoFound", "û�л�û���ҵ��κ�ְλ��Ϣ") & "<br><br></p>"
        rsSearch.Close
        Set rsSearch = Nothing
        GetSearchResult = strSearchResult
        Exit Function
    Else
        If UsePage = True Then
            totalPut = rsSearch.RecordCount
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
                    rsSearch.Move (CurrentPage - 1) * MaxPerPage
                Else
                    CurrentPage = 1
                End If
            End If
        End If
    End If
    
    Dim iPositionIDLen, iPositionNameLen, iWorkPlaceNameLen, iSubCompanyNameLen, iPositionNumLen, iPositionStatusLen, iValidDateLen, iReleaseDateLen
    iPositionIDLen = R_XmlText_Class("ShowSearch/PositionIDLen", "60")
    iPositionNameLen = R_XmlText_Class("ShowSearch/PositionNameLen", "80")
    iWorkPlaceNameLen = R_XmlText_Class("ShowSearch/WorkPlaceNameLen", "80")
    iSubCompanyNameLen = R_XmlText_Class("ShowSearch/SubCompanyNameLen", "120")
    iPositionNumLen = R_XmlText_Class("ShowSearch/PositionNumLen", "60")
    iPositionStatusLen = R_XmlText_Class("ShowSearch/PositionStatusLen", "60")
    iValidDateLen = R_XmlText_Class("ShowSearch/ValidDateLen", "60")
    iReleaseDateLen = R_XmlText_Class("ShowSearch/ReleaseDateLen", "80")
    
    strSearchResult = strSearchResult & "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    strSearchResult = strSearchResult & "         <tr height='22'> "
    If ShowPositionID > 0 Then
        strSearchResult = strSearchResult & "            <td width='" & iPositionIDLen & "' align='center'>" & R_XmlText_Class("ShowSearch/PositionID", "���") & "</td>"
    End If
    If ShowPositionName > 0 Then
        strSearchResult = strSearchResult & "            <td width='" & iPositionNameLen & "' align='center'>" & R_XmlText_Class("ShowSearch/PositionName", "ְλ����") & "</td>"
    End If
    If ShowWorkPlaceName > 0 Then
        strSearchResult = strSearchResult & "            <td width='" & iWorkPlaceNameLen & "' align='center'>" & R_XmlText_Class("ShowSearch/WorkPlaceName", "�����ص�") & "</td>"
    End If
    If ShowSubCompanyName > 0 Then
        strSearchResult = strSearchResult & "            <td width='" & iSubCompanyNameLen & "' align='center'>" & R_XmlText_Class("ShowSearch/SubCompanyName", "���˵�λ") & "</td>"
    End If
    If ShowPositionStatus > 0 Then
        strSearchResult = strSearchResult & "            <td width='" & iPositionStatusLen & "' align='center'>" & R_XmlText_Class("ShowSearch/PositionStatus", "ְλ״̬") & "</td>"
    End If
    If ShowPositionNum > 0 Then
        strSearchResult = strSearchResult & "            <td width='" & iPositionNumLen & "' align='center'>" & R_XmlText_Class("ShowSearch/PositionNum", "��Ƹ����") & "</td>"
    End If
    If ShowDateType > 0 Then
        strSearchResult = strSearchResult & "            <td width='" & iReleaseDateLen & "' align='center'>" & R_XmlText_Class("ShowSearch/ReleaseDate", "��������") & "</td>"
    End If
    If ShowValidDate > 0 Then
        strSearchResult = strSearchResult & "            <td width='" & iValidDateLen & "' align='center'>" & R_XmlText_Class("ShowSearch/ValidDate", "��Ч��") & "</td>"
    End If
    strSearchResult = strSearchResult & "          </tr>"
    iCount = 0


    Do While Not rsSearch.EOF
        If TitleLen > 0 Then
            TitleStr = GetSubStr(rsSearch("PositionName"), TitleLen, PShowPoints)
        Else
            TitleStr = rsSearch("PositionName")
        End If
        
        strLink = "<a href='SupplyInfo.asp?PositionID=" & rsSearch("PositionID") & "'"
        If OpenType = 0 Then
            strLink = strLink & " target='_self'>"
        Else
            strLink = strLink & " target='_blank'>"
        End If
        strLink = strLink & TitleStr & "</a>"

        strSearchResult = strSearchResult & "      <tr>"
        If ShowPositionID > 0 Then
            strSearchResult = strSearchResult & "        <td width='" & iPositionIDLen & "' align='center'>" & rsSearch("PositionID") & "</td>"
        End If
        If ShowPositionName > 0 Then
            strSearchResult = strSearchResult & "        <td width='" & iPositionNameLen & "' align='center'>" & strLink & " </td>"
        End If
        If ShowWorkPlaceName > 0 Then
            If WorkPlaceNameLen > 0 Then
                strSearchResult = strSearchResult & "      <td width='" & iWorkPlaceNameLen & "' align='center'>" & GetSubStr(rsSearch("WorkPlaceName"), WorkPlaceNameLen, WShowPoints) & "</td>"
                strSearchResult = strSearchResult & "      <td width='" & iWorkPlaceNameLen & "' align='center'>" & GetSubStr(rsSearch("WorkPlaceName"), WorkPlaceNameLen, WShowPoints) & "</td>"
            Else
                strSearchResult = strSearchResult & "      <td width='" & iWorkPlaceNameLen & "' align='center'>" & rsSearch("WorkPlaceName") & "</td>"
            End If
        End If
        If ShowSubCompanyName > 0 Then
            If SubCompanyNameLen > 0 Then
                strSearchResult = strSearchResult & "      <td width='" & iSubCompanyNameLen & "' align='center'>" & GetSubStr(rsSearch("SubCompanyName"), SubCompanyNameLen, SShowPoints) & "</td>"
            Else
                strSearchResult = strSearchResult & "      <td width='" & iSubCompanyNameLen & "' align='center'>" & rsSearch("SubCompanyName") & "</td>"
            End If
        End If
        If ShowPositionStatus > 0 Then
            strSearchResult = strSearchResult & "      <td width='" & iPositionStatusLen & "' align='center'>" & GetPositionStatus_Search(rsSearch("PositionStatus"), rsSearch("ReleaseDate"), rsSearch("ValidDate")) & "</td>"
        End If
        If ShowPositionNum > 0 Then
            strSearchResult = strSearchResult & "      <td width='" & iPositionNumLen & "' align='center'>" & rsSearch("PositionNum") & "</td>"
        End If
        If ShowDateType > 0 Then
            strSearchResult = strSearchResult & "      <td width='" & iReleaseDateLen & "' align='center'>" & GetUpdateTimeStr(rsSearch("ReleaseDate"), ShowDateType) & "</td>"
        End If
        If ShowValidDate > 0 Then
            strSearchResult = strSearchResult & "      <td width='" & iValidDateLen & "' align='center'>" & rsSearch("ValidDate") & "</td>"
        End If
        strSearchResult = strSearchResult & "     </tr>"
        rsSearch.MoveNext
        iCount = iCount + 1
        If UsePage = True And iCount >= MaxPerPage Then Exit Do
    Loop
    rsSearch.Close
    Set rsSearch = Nothing
    strSearchResult = strSearchResult & "     </table>"
    GetSearchResult = strSearchResult
End Function

Function XmlText_Class(ByVal iSmallNode, ByVal DefChar)
    XmlText_Class = XmlText("Job", iSmallNode, DefChar)
End Function

Function R_XmlText_Class(ByVal iSmallNode, ByVal DefChar)
    R_XmlText_Class = Replace(XmlText("Job", iSmallNode, DefChar), "{$ChannelShortName}", ChannelShortName)
End Function

End Class
%>
