<!--#include file="Start.asp"-->
<!--#include file="Include/PowerEasy.Cache.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Response.ContentType = "text/xml; charset=gb2312"
    
Dim strtmp, SubNode, DynaDom, DynaNode

Set XMLDOM = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
strtmp = "<?xml version=""1.0"" encoding=""gb2312""?>"

XMLDOM.appendChild (XMLDOM.createProcessingInstruction("xml", "version=""1.0"" encoding=""gb2312"""))
XMLDOM.appendChild (XMLDOM.createElement("root"))
XMLDOM.documentElement.Attributes.setNamedItem(XMLDOM.createNode(2, "version", "")).Text = "PowerEasy Cms 2006"


'��������
Set DynaDom = CreateObject("Microsoft.XMLDOM")
DynaDom.async = False
DynaDom.Load Request
Set DynaNode = DynaDom.getElementsByTagName("root")
If DynaNode.length < 1 Then
    Set Node = XMLDOM.createNode(1, "serverbackinfo", "")
    XMLDOM.documentElement.appendChild (Node)
    Set SubNode = Node.appendChild(XMLDOM.createElement("stat"))
    SubNode.Text = "err"
    Set SubNode = Node.appendChild(XMLDOM.createElement("infomation"))
    SubNode.Text = "�������ݴ���!"
Else
    Dim id, page, tempvaluearr
    id = PE_CLng(DynaNode(0).selectSingleNode("id").Text)
    If id > 0 Then
        If PE_CLng(DynaNode(0).selectSingleNode("page").Text) > 0 Then
            page = PE_CLng(DynaNode(0).selectSingleNode("page").Text)
        Else
            page = 1
        End If
        If DynaNode(0).selectSingleNode("value").Text <> "" Then
            tempvaluearr = Split(DynaNode(0).selectSingleNode("value").Text, "|")
        End If

        '��ʼ�����̬��ǩ����
        Dim rsLabel
        Set rsLabel = Conn.Execute("select LabelID,LabelName,LabelType,PageNum,LabelIntro,LabelContent from PE_Label where LabelID=" & id)
        If rsLabel.BOF And rsLabel.EOF Then
            Set Node = XMLDOM.createNode(1, "serverbackinfo", "")
            XMLDOM.documentElement.appendChild (Node)
            Set SubNode = Node.appendChild(XMLDOM.createElement("stat"))
            SubNode.Text = "err"
            Set SubNode = Node.appendChild(XMLDOM.createElement("infomation"))
            SubNode.Text = "��ǩ������!"
        Else
            '�ҵ���ǩ���д���
            Dim rsLabelRe, PageNum, TempSql, LoopTemp, loopTempMatch, InfoID, tempvalue
            Dim DyTemp, j, InfoTemp, InfoTempMatch, MatchesInfo, FieldTemp, FieldArry, FieldTempText
            PageNum = rsLabel("PageNum")

            LoopTemp = rsLabel("LabelContent")
            LoopTemp = Replace(Replace(Replace(Replace(LoopTemp, "{$Now}", Now()), "{$NowDay}", Day(Now())), "{$NowMonth}", Month(Now())), "{$NowYear}", Year(Now()))
            LoopTemp = Replace(Replace(Replace(Replace(LoopTemp, "{$PE_True}", PE_True), "{$PE_False}", PE_False), "{$PE_Now}", PE_Now), "{$PE_OrderType}", PE_OrderType)
            If rsLabel("LabelType") = 3 Then '�����Ͷ�̬��ǩ�Ĵ������
                 For j = 0 To UBound(tempvaluearr)
                     LoopTemp = Replace(LoopTemp, "{input(" & j & ")}", tempvaluearr(j))
                 Next
            End If

            regEx.Pattern = "\{Loop\}([\s\S]*?)\{\/Loop\}"
            Set Matches = regEx.Execute(LoopTemp)
            For Each Match In Matches
                loopTempMatch = Match.Value
            Next
            LoopTemp = regEx.Replace(LoopTemp, "{$SqlReplaceText}")
            loopTempMatch = Replace(Replace(loopTempMatch, "{loop}", ""), "{/loop}", "")

            TempSql = Replace(Replace(Replace(Replace(rsLabel("LabelIntro"), "{$Now}", Now()), "{$NowDay}", Day(Now())), "{$NowMonth}", Month(Now())), "{$NowYear}", Year(Now()))
            TempSql = Replace(Replace(Replace(Replace(TempSql, "{$PE_True}", PE_True), "{$PE_False}", PE_False), "{$PE_Now}", PE_Now), "{$PE_OrderType}", PE_OrderType)
            If rsLabel("LabelType") = 3 Then '�����Ͷ�̬��ǩ�Ĵ������
                For j = 0 To UBound(tempvaluearr) - 1
                    TempSql = Replace(TempSql, "{input(" & j & ")}", ReplaceLabelBadChar(tempvaluearr(j)))
                Next
            End If
                '��ʼѭ����������
                Dim totalpage, iMod
                InfoID = 0
                On Error Resume Next
                Set rsLabelRe = Server.CreateObject("adodb.recordset")
                rsLabelRe.Open TempSql, Conn, 1, 1
                If Err Then
                    Err.Clear
                    DyTemp = "SQL��ѯ��"
                Else
                    totalPut = rsLabelRe.RecordCount
                    If (totalPut Mod PageNum) = 0 Then
                        totalpage = totalPut \ PageNum
                    Else
                        totalpage = totalPut \ PageNum + 1
                    End If
                    If page < 1 Then
                        page = 1
                    End If
                    If (page - 1) * PageNum > totalPut Then
                        If (totalPut Mod PageNum) = 0 Then
                            page = totalPut \ PageNum
                        Else
                            page = totalPut \ PageNum + 1
                        End If
                    End If
                    If page > 1 Then
                        If (page - 1) * PageNum < totalPut Then
                            iMod = 0
                            If page > PageNum Then
                                iMod = totalPut Mod PageNum
                                If iMod <> 0 Then iMod = PageNum - iMod
                            End If
                            rsLabelRe.Move (page - 1) * PageNum - iMod
                        Else
                            page = 1
                        End If
                    End If

                    If rsLabelRe.BOF And rsLabelRe.EOF Then
                        DyTemp = "������"
                    Else
                        Do While Not rsLabelRe.EOF
                        regEx.Pattern = "\{Infobegin\}([\s\S]*?)\{Infoend\}"
                        Set Matches = regEx.Execute(loopTempMatch)
                        If Matches.Count = 0 Then
                            rsLabelRe.MoveNext
                        Else
                            For Each Match In Matches
                                If Not rsLabelRe.EOF Then
                                    InfoTemp = Match.Value
                                    InfoTempMatch = Replace(Replace(InfoTemp, "{Infobegin}", ""), "{Infoend}", "") '�õ����յĵ�һ�ֶ�����
                                    regEx.Pattern = "\{\$Field\((.*?)\)\}"
                                    Set MatchesInfo = regEx.Execute(InfoTempMatch)
                                    For Each Match2 In MatchesInfo
                                        FieldTemp = Match2.Value
                                        FieldArry = Split(Match2.SubMatches(0), ",")
                                        If UBound(FieldArry) > 1 Then '������ȷ,���д���
                                            Select Case FieldArry(1)
                                            Case "Text" '���ı���ʽ�������
                                                If rsLabelRe(PE_CLng(FieldArry(0))) = "" Or IsNull(rsLabelRe(PE_CLng(FieldArry(0)))) Then
                                                    FieldTempText = ""
                                                Else
                                                    If FieldArry(2) = 0 Then
                                                        Select Case FieldArry(3)
                                                        Case 1
                                                            FieldTempText = Replace(rsLabelRe(PE_CLng(FieldArry(0))), "<", "&lt;")
                                                        Case 2
                                                            FieldTempText = nohtml(rsLabelRe(PE_CLng(FieldArry(0))))
                                                        Case Else
                                                            FieldTempText = rsLabelRe(PE_CLng(FieldArry(0)))
                                                        End Select
                                                    Else
                                                        Select Case FieldArry(3)
                                                        Case 1
                                                            If FieldArry(4) = 0 Then
                                                                FieldTempText = GetSubStr(Replace(rsLabelRe(PE_CLng(FieldArry(0))), "<", "&lt;"), PE_CLng(FieldArry(2)), True)
                                                            Else
                                                                FieldTempText = GetSubStr(Replace(rsLabelRe(PE_CLng(FieldArry(0))), "<", "&lt;"), PE_CLng(FieldArry(2)), False)
                                                            End If
                                                        Case 2
                                                            If FieldArry(4) = 0 Then
                                                                FieldTempText = GetSubStr(nohtml(rsLabelRe(PE_CLng(FieldArry(0)))), PE_CLng(FieldArry(2)), True)
                                                            Else
                                                                FieldTempText = GetSubStr(nohtml(rsLabelRe(PE_CLng(FieldArry(0)))), PE_CLng(FieldArry(2)), False)
                                                            End If
                                                        Case Else
                                                            If FieldArry(4) = 0 Then
                                                                FieldTempText = GetSubStr(rsLabelRe(PE_CLng(FieldArry(0))), PE_CLng(FieldArry(2)), True)
                                                            Else
                                                                FieldTempText = GetSubStr(rsLabelRe(PE_CLng(FieldArry(0))), PE_CLng(FieldArry(2)), False)
                                                            End If
                                                        End Select
                                                    End If
                                                 End If
                                            Case "Num" '�����ַ�ʽ�������
                                                If rsLabelRe(PE_CLng(FieldArry(0))) = "" Or IsNull(rsLabelRe(PE_CLng(FieldArry(0)))) Then
                                                    FieldTempText = "0"
                                                Else
                                                    Select Case FieldArry(2)
                                                    Case 0
                                                        If FieldArry(3) = "0" Then
                                                            FieldTempText = Int(rsLabelRe(PE_CLng(FieldArry(0))))
                                                        Else
                                                            FieldTempText = String(Int(rsLabelRe(PE_CLng(FieldArry(0)))), FieldArry(3))
                                                        End If
                                                    Case 1
                                                        FieldTempText = FormatNumber(rsLabelRe(PE_CLng(FieldArry(0))), FieldArry(3))
                                                    Case 2
                                                        FieldTempText = FormatPercent(rsLabelRe(PE_CLng(FieldArry(0))))
                                                    End Select
                                               End If
                                            Case "Time" '��ʱ�䷽ʽ�������
                                                Dim temptime, temptimetext
                                                If rsLabelRe(PE_CLng(FieldArry(0))) = "" Or IsNull(rsLabelRe(PE_CLng(FieldArry(0)))) Then
                                                    FieldTempText = ""
                                                Else
                                                    If IsDate(rsLabelRe(PE_CLng(FieldArry(0)))) Then '�ж��ֶ������Ƿ���ȷ
                                                        temptime = rsLabelRe(PE_CLng(FieldArry(0)))
                                                        Select Case FieldArry(2)
                                                        Case 0
                                                            FieldTempText = Replace(Replace(Replace(Replace(Replace(Replace(FieldArry(3), "{year}", Year(temptime)), "{month}", Month(temptime)), "{day}", Day(temptime)), "{Hour}", Hour(temptime)), "{Minute}", Minute(temptime)), "{Second}", Second(temptime))
                                                        Case 1, 2
                                                            If FieldArry(2) = 1 Then
                                                                temptimetext = Replace(FieldArry(3), "{year}", Year(temptime))
                                                            Else
                                                                temptimetext = Replace(FieldArry(3), "{year}", Right(Year(temptime), 2))
                                                            End If
                                                            If Len(Month(temptime)) = 1 Then
                                                                temptimetext = Replace(temptimetext, "{month}", "0" & Month(temptime))
                                                            Else
                                                                temptimetext = Replace(temptimetext, "{month}", Month(temptime))
                                                            End If
                                                            If Len(Day(temptime)) = 1 Then
                                                                temptimetext = Replace(temptimetext, "{day}", "0" & Day(temptime))
                                                            Else
                                                                temptimetext = Replace(temptimetext, "{day}", Day(temptime))
                                                            End If
                                                            FieldTempText = temptimetext
                                                        Case 3
                                                            FieldTempText = FormatDateTime(temptime, PE_CLng(FieldArry(3)))
                                                        End Select
                                                    Else
                                                        FieldTempText = "���ֶη�ʱ����"
                                                    End If
                                                End If
                                            Case "yn" '���Ƿ�ʽ�������
                                                If rsLabelRe(PE_CLng(FieldArry(0))) = "" Or IsNull(rsLabelRe(PE_CLng(FieldArry(0)))) Then
                                                    FieldTempText = ""
                                                Else
                                                    If rsLabelRe(PE_CLng(FieldArry(0))) = True Then
                                                        FieldTempText = FieldArry(2)
                                                    Else
                                                        FieldTempText = FieldArry(3)
                                                    End If
                                                End If
                                            Case "GetUrl"
                                                FieldTempText = GetInfoUrl(rsLabelRe(PE_CLng(FieldArry(0))), FieldArry(2), FieldArry(3))
                                            Case "GetClass"
                                                FieldTempText = GetInfoClass(rsLabelRe(PE_CLng(FieldArry(0))), FieldArry(2))
                                            Case "GetSpecil"
                                                FieldTempText = GetInfoSpecil(rsLabelRe(PE_CLng(FieldArry(0))), FieldArry(2))
                                            Case "GetChannel"
                                                FieldTempText = GetInfoChannel(rsLabelRe(PE_CLng(FieldArry(0))), FieldArry(2))
                                            Case Else
                                                FieldTempText = "��ǩ��������"
                                            End Select
                                        Else
                                            FieldTempText = "��ǩ��������"
                                        End If
                                        If Trim(FieldTempText & "") = "" Then
                                            InfoTempMatch = Replace(InfoTempMatch, FieldTemp, "")
                                        Else
                                            InfoTempMatch = Replace(InfoTempMatch, FieldTemp, FieldTempText)
                                        End If
                                    Next
                                    Dim tempid
                                    tempid = 1 + (PageNum * (page - 1))
                                    DyTemp = DyTemp & Replace(InfoTempMatch, "{$AutoID}", InfoID + tempid)
                                    rsLabelRe.MoveNext
                                    InfoID = InfoID + 1
                                    If InfoID >= PageNum Then Exit Do
                                End If
                                Next
                            End If
                        Loop
                    End If
                End If
                rsLabelRe.Close
                LoopTemp = Replace(LoopTemp, "{$SqlReplaceText}", DyTemp)
                LoopTemp = Replace(LoopTemp, "{$totalPut}", totalPut)


                regEx.Pattern = "\{\$InstallDir\}(?!\{\$ChannelDir\})"
                LoopTemp = regEx.Replace(LoopTemp, InstallDir)
                LoopTemp = PE_Replace(LoopTemp, "{$ADDir}", ADDir)
                LoopTemp = PE_Replace(LoopTemp, "{$SiteUrl}", SiteUrl)
                LoopTemp = PE_Replace(LoopTemp, "{$SiteName}", SiteName)
                LoopTemp = PE_Replace(LoopTemp, "{$WebmasterEmail}", WebmasterEmail)
                LoopTemp = PE_Replace(LoopTemp, "{$WebmasterName}", WebmasterName)
                LoopTemp = PE_Replace(LoopTemp, "{$Copyright}", Copyright)
                LoopTemp = PE_Replace(LoopTemp, "{$Meta_Keywords}", Meta_Keywords)
                LoopTemp = PE_Replace(LoopTemp, "{$Meta_Description}", Meta_Description)
				
	  			
                '�滻{$YN(Condition,Fir,Sec)}��ǩ
                Dim strYN,arrYnTemp
                regEx.Pattern = "\{\$YN\((.*?)\)\}"
                Set Matches = regEx.Execute(LoopTemp)
                For Each Match In Matches
                    arrYnTemp = Split(Match.SubMatches(0), ",")
                    If UBound(arrYnTemp) <> 2 Then
                        strYN = "����ʽ��ǩ��{$YN(�����б�)}�Ĳ����������ԡ�����ģ���еĴ˱�ǩ��"
                    Else
                        strYN = YN(arrYnTemp(0), arrYnTemp(1), arrYnTemp(2))
                    End If
                    LoopTemp = Replace(LoopTemp, Match.Value, strYN)
                Next
	
	
	
            Set rsLabelRe = Nothing
            '�����XML
            Set Node = XMLDOM.createNode(1, "serverbackinfo", "")
            XMLDOM.documentElement.appendChild (Node)
            Set SubNode = Node.appendChild(XMLDOM.createElement("stat"))
            SubNode.Text = "ok"
            Set SubNode = Node.appendChild(XMLDOM.createElement("id"))
            SubNode.Text = id
            Set SubNode = Node.appendChild(XMLDOM.createElement("content"))
            SubNode.Text = LoopTemp
            Set SubNode = Node.appendChild(XMLDOM.createElement("rootdir"))
            SubNode.Text = InstallDir
            Set SubNode = Node.appendChild(XMLDOM.createElement("totalpage"))
            SubNode.Text = totalpage
            Set SubNode = Node.appendChild(XMLDOM.createElement("currentpage"))
            SubNode.Text = page
            Set SubNode = Node.appendChild(XMLDOM.createElement("totalitem"))
            SubNode.Text = totalPut
            Set SubNode = Node.appendChild(XMLDOM.createElement("value"))
            SubNode.Text = DynaNode(0).selectSingleNode("value").Text
        End If
        Set rsLabel = Nothing
    Else
        Set Node = XMLDOM.createNode(1, "serverbackinfo", "")
        XMLDOM.documentElement.appendChild (Node)
        Set SubNode = Node.appendChild(XMLDOM.createElement("stat"))
        SubNode.Text = "err"
        Set SubNode = Node.appendChild(XMLDOM.createElement("infomation"))
        SubNode.Text = "�������ݴ���!"
    End If
End If
strtmp = strtmp & XMLDOM.documentElement.xml
Response.Write strtmp
Set XMLDOM = Nothing
Set DynaDom = Nothing
Call CloseConn

'==================================================
'��������GetInfoChannel
'��  �ã���ȡ�����Ƶ������
'��  ����InfoID ------����ID
'      ��OutType -----�����ʽ
'==================================================
Function GetInfoChannel(InfoID, OutType)
    If IsNull(InfoID) = True Or IsNull(OutType) = True Then
        GetInfoChannel = ""
        Exit Function
    End If
    Dim sqlInfo, rsInfo, rsChannel2, strTemp
    sqlInfo = "select top 1 ChannelID,ChannelName,LinkUrl,ChannelDir,Disabled,UploadDir from PE_Channel Where ChannelID=" & InfoID
    Set rsInfo = Conn.Execute(sqlInfo)
    If Not (rsInfo.BOF And rsInfo.EOF) Then
        If rsInfo("Disabled") = True Then
                strTemp = ""
        Else
            Select Case OutType
            Case 1
                If IsNull(rsInfo("ChannelDir")) Then
                    strTemp = rsInfo("LinkUrl")
                Else
                    strTemp = rsInfo("ChannelDir")
                End If
            Case 2
                strTemp = rsInfo("ChannelName")
            Case 3
                strTemp = rsInfo("UploadDir")
            Case Else
                strTemp = "��ǩ������"
            End Select
        End If
    End If
    rsInfo.Close
    Set rsInfo = Nothing
    GetInfoChannel = strTemp
End Function

'==================================================
'��������GetInfoUrl
'��  �ã���ȡ�����·��
'��  ����InfoID ------����ID
'      ��DataType ------���ݿ�����
'==================================================
Function GetInfoUrl(InfoID, DataType, OutType)
    If IsNull(InfoID) = True Or IsNull(DataType) = True Or IsNull(OutType) = True Then
        GetInfoUrl = ""
        Exit Function
    End If
    Dim sqlInfo, rsInfo, rsChannel2, strTemp
    Dim ChannelDir, StructureType, FileNameType, FileExtType, iUseCreateHTML, CacheTemp, ChannelTemp
    Select Case DataType
    Case "Article"
        sqlInfo = "select top 1 A.ArticleID,A.ChannelID,A.ClassID,A.Title,A.UpdateTime,A.InfoPoint,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Article A inner join PE_Class C on A.ClassID=C.ClassID Where A.ArticleID=" & InfoID
    Case "Soft"
        sqlInfo = "select top 1 A.SoftID,A.ChannelID,A.ClassID,A.SoftName,A.UpdateTime,A.InfoPoint,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Soft A inner join PE_Class C on A.ClassID=C.ClassID Where A.SoftID=" & InfoID
    Case "Photo"
        sqlInfo = "select top 1 A.PhotoID,A.ChannelID,A.ClassID,A.PhotoName,A.UpdateTime,A.InfoPoint,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Photo A inner join PE_Class C on A.ClassID=C.ClassID Where A.PhotoID=" & InfoID
    Case "Product"
        sqlInfo = "select top 1 A.ProductID,A.ChannelID,A.ClassID,A.ProductName,A.UpdateTime,A.Stocks,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Product A inner join PE_Class C on A.ClassID=C.ClassID Where A.ProductID=" & InfoID
    Case Else
        GetInfoUrl = InfoID
        Exit Function
    End Select
    Set rsInfo = Conn.Execute(sqlInfo)
    If Not (rsInfo.BOF And rsInfo.EOF) Then
        If PE_Cache.CacheIsEmpty("InfoUrl_" & DataType) Then
            Set rsChannel2 = Conn.Execute("select ChannelID,ChannelDir,StructureType,FileNameType,FileExt_Item,UseCreateHTML from PE_Channel Where ChannelID=" & rsInfo(1) & " and Disabled=" & PE_False)
            If Not (rsChannel2.BOF And rsChannel2.EOF) Then
                ChannelDir = rsChannel2("ChannelDir")
                StructureType = rsChannel2("StructureType")
                FileNameType = rsChannel2("FileNameType")
                FileExtType = rsChannel2("FileExt_Item")
                iUseCreateHTML = rsChannel2("UseCreateHTML")
                CacheTemp = rsChannel2("ChannelID") & "|||" & rsChannel2("ChannelDir") & "|||" & rsChannel2("StructureType") & "|||" & rsChannel2("FileNameType") & "|||" & rsChannel2("FileExt_Item") & "|||" & rsChannel2("UseCreateHTML")
                PE_Cache.SetValue "InfoUrl_" & DataType, CacheTemp
            Else
                strTemp = InfoID
            End If
            rsChannel2.Close
            Set rsChannel2 = Nothing
        Else
            ChannelTemp = Split(PE_Cache.GetValue("InfoUrl_" & DataType), "|||")
            If rsInfo(1) = ChannelTemp(0) Then
                ChannelDir = ChannelTemp(1)
                StructureType = ChannelTemp(2)
                FileNameType = ChannelTemp(3)
                FileExtType = ChannelTemp(4)
                iUseCreateHTML = ChannelTemp(5)
            Else
                Set rsChannel2 = Conn.Execute("select ChannelID,ChannelDir,StructureType,FileNameType,FileExt_Item,UseCreateHTML from PE_Channel Where ChannelID=" & rsInfo(1) & " and Disabled=" & PE_False)
                If Not (rsChannel2.BOF And rsChannel2.EOF) Then
                    ChannelDir = rsChannel2("ChannelDir")
                    StructureType = rsChannel2("StructureType")
                    FileNameType = rsChannel2("FileNameType")
                    FileExtType = rsChannel2("FileExt_Item")
                    iUseCreateHTML = rsChannel2("UseCreateHTML")
                    CacheTemp = rsChannel2("ChannelID") & "|||" & rsChannel2("ChannelDir") & "|||" & rsChannel2("StructureType") & "|||" & rsChannel2("FileNameType") & "|||" & rsChannel2("FileExt_Item") & "|||" & rsChannel2("UseCreateHTML")
                    PE_Cache.SetValue "InfoUrl_" & DataType, CacheTemp
                Else
                    strTemp = InfoID
                End If
                rsChannel2.Close
                Set rsChannel2 = Nothing
            End If
        End If
        If strTemp <> InfoID Then
            Select Case OutType
            Case 1
                If iUseCreateHTML > 0 Then
                    If DataType = "Product" Then
                        strTemp = ChannelDir & GetItemPath(StructureType, rsInfo(7), rsInfo(6), rsInfo(4)) & GetItemFileName(FileNameType, ChannelDir, rsInfo(4), InfoID) & arrFileExt(FileExtType)
                    Else
                        If rsInfo(8) = 0 And rsInfo(5) = 0 Then
                            strTemp = ChannelDir & GetItemPath(StructureType, rsInfo(7), rsInfo(6), rsInfo(4)) & GetItemFileName(FileNameType, ChannelDir, rsInfo(4), InfoID) & arrFileExt(FileExtType)
                        Else
                            strTemp = ChannelDir & "/Show" & DataType & ".asp?" & DataType & "ID=" & rsInfo(0)
                        End If
                    End If
                Else
                    strTemp = ChannelDir & "/Show" & DataType & ".asp?" & DataType & "ID=" & rsInfo(0)
                End If
            Case 2
                strTemp = rsInfo(3)
            Case 3
                If iUseCreateHTML > 0 Then
                    If DataType = "Product" Then
                        strTemp = "<a href='" & InstallDir & ChannelDir & GetItemPath(StructureType, rsInfo(7), rsInfo(6), rsInfo(4)) & GetItemFileName(FileNameType, ChannelDir, rsInfo(4), InfoID) & arrFileExt(FileExtType) & "'>" & rsInfo(3) & "</a>"
                    Else
                        If rsInfo(8) = 0 And rsInfo(5) = 0 Then
                            strTemp = "<a href='" & InstallDir & ChannelDir & GetItemPath(StructureType, rsInfo(7), rsInfo(6), rsInfo(4)) & GetItemFileName(FileNameType, ChannelDir, rsInfo(4), InfoID) & arrFileExt(FileExtType) & "'>" & rsInfo(3) & "</a>"
                        Else
                            strTemp = "<a href='" & InstallDir & ChannelDir & "/Show" & DataType & ".asp?" & DataType & "ID=" & rsInfo(0) & "'>" & rsInfo(3) & "</a>"
                        End If
                    End If
                Else
                    strTemp = "<a href='" & InstallDir & ChannelDir & "/Show" & DataType & ".asp?" & DataType & "ID=" & rsInfo(0) & "'>" & rsInfo(3) & "</a>"
                End If
            Case Else
                strTemp = "��ǩ������"
            End Select
        End If
    End If
    rsInfo.Close
    Set rsInfo = Nothing
    GetInfoUrl = strTemp
End Function

'==================================================
'��������GetInfoClass
'��  �ã���ȡ����ķ���
'��  ����InfoID ------����ID
'      ��DataType ------���ݿ�����
'==================================================
Function GetInfoClass(InfoID, OutType)
    If IsNull(InfoID) = True Or IsNull(OutType) = True Then
        GetInfoClass = ""
        Exit Function
    End If
    Dim sqlInfo, rsInfo, rsChannel2, strTemp, PriChannelID
    Dim ChannelDir, ModuleType, StructureType, ListFileType, FileExtList, iUseCreateHTML
    sqlInfo = "select top 1 ClassID,ChannelID,ClassName,ClassDir,ParentDir,ClassPurview from PE_Class Where ClassID=" & InfoID
    Set rsInfo = Conn.Execute(sqlInfo)
    If Not (rsInfo.BOF And rsInfo.EOF) Then
        If rsInfo("ChannelID") <> PriChannelID Then
            Set rsChannel2 = Conn.Execute("select ChannelID,ChannelDir,ModuleType,StructureType,ListFileType,FileExt_List,UseCreateHTML from PE_Channel Where ChannelID=" & rsInfo("ChannelID") & " and Disabled=" & PE_False)
            If Not (rsChannel2.BOF And rsChannel2.EOF) Then
                ChannelDir = rsChannel2("ChannelDir")
                ModuleType = rsChannel2("ModuleType")
                StructureType = rsChannel2("StructureType")
                ListFileType = rsChannel2("ListFileType")
                FileExtList = rsChannel2("FileExt_List")
                iUseCreateHTML = rsChannel2("UseCreateHTML")
                PriChannelID = rsInfo("ChannelID")
            Else
                strTemp = "��Ŀ������"
            End If
            rsChannel2.Close
            Set rsChannel2 = Nothing
        End If

        If strTemp <> "��Ŀ������" Then
            Select Case OutType
            Case 1
                If iUseCreateHTML = 1 Or iUseCreateHTML = 3 Then
                    If ModuleType = 5 Then
                        strTemp = ChannelDir & GetListPath(StructureType, ListFileType, rsInfo("ParentDir"), rsInfo("ClassDir")) & GetListFileName(ListFileType, rsInfo("ClassID"), 1, 1) & arrFileExt(FileExtList)
                    Else
                        If rsInfo("ClassPurview") < 2 Then
                            strTemp = ChannelDir & GetListPath(StructureType, ListFileType, rsInfo("ParentDir"), rsInfo("ClassDir")) & GetListFileName(ListFileType, rsInfo("ClassID"), 1, 1) & arrFileExt(FileExtList)
                        Else
                            strTemp = ChannelDir & "/ShowClass.asp?ClassID=" & rsInfo("ClassID")
                        End If
                    End If
                Else
                    strTemp = ChannelDir & "/ShowClass.asp?ClassID=" & rsInfo("ClassID")
                End If
            Case 2
                strTemp = rsInfo("ClassName")
            Case 3
                If iUseCreateHTML = 1 Or iUseCreateHTML = 3 Then
                    If ModuleType = 5 Then
                        strTemp = "<a href='" & InstallDir & ChannelDir & GetListPath(StructureType, ListFileType, rsInfo("ParentDir"), rsInfo("ClassDir")) & GetListFileName(ListFileType, rsInfo("ClassID"), 1, 1) & arrFileExt(FileExtList) & "'>" & rsInfo("ClassName") & "</a>"
                    Else
                        If rsInfo("ClassPurview") < 2 Then
                            strTemp = "<a href='" & InstallDir & ChannelDir & GetListPath(StructureType, ListFileType, rsInfo("ParentDir"), rsInfo("ClassDir")) & GetListFileName(ListFileType, rsInfo("ClassID"), 1, 1) & arrFileExt(FileExtList) & "'>" & rsInfo("ClassName") & "</a>"
                        Else
                            strTemp = "<a href='" & InstallDir & ChannelDir & "/ShowClass.asp?ClassID=" & rsInfo("ClassID") & "'>" & rsInfo("ClassName") & "</a>"
                        End If
                    End If
                Else
                    strTemp = "<a href='" & InstallDir & ChannelDir & "/ShowClass.asp?ClassID=" & rsInfo("ClassID") & "'>" & rsInfo("ClassName") & "</a>"
                End If
            Case Else
                strTemp = "��ǩ������"
            End Select
            GetInfoClass = strTemp
        Else
            GetInfoClass = ""
        End If
    End If
    rsInfo.Close
    Set rsInfo = Nothing
End Function

'==================================================
'��������GetInfoSpecil
'��  �ã���ȡ�����ר��
'��  ����InfoID ------����ID
'      ��DataType ------���ݿ�����
'==================================================
Function GetInfoSpecil(InfoID, OutType)
    If IsNull(InfoID) = True Or IsNull(OutType) = True Then
        GetInfoSpecil = ""
        Exit Function
    End If
    Dim sqlInfo, rsInfo, rsChannel2, strTemp, PriChannelID
    Dim ChannelDir, iUseCreateHTML
    sqlInfo = "select top 1 A.ChannelID,I.SpecialID,SP.SpecialName,SP.SpecialDir from PE_Article A right join (PE_InfoS I left join PE_Special SP on I.SpecialID=SP.SpecialID) on A.ArticleID=I.ItemID Where A.ArticleID=" & InfoID
    Set rsInfo = Conn.Execute(sqlInfo)
    If Not (rsInfo.BOF And rsInfo.EOF) Then
        If rsInfo(0) <> PriChannelID Then
            Set rsChannel2 = Conn.Execute("select ChannelID,ChannelDir,UseCreateHTML from PE_Channel Where ChannelID=" & rsInfo(0) & " and Disabled=" & PE_False)
            If Not (rsChannel2.BOF And rsChannel2.EOF) Then
                ChannelDir = rsChannel2("ChannelDir")
                iUseCreateHTML = rsChannel2("UseCreateHTML")
                PriChannelID = rsInfo(0)
            Else
                strTemp = "ר�ⲻ����"
            End If
            rsChannel2.Close
            Set rsChannel2 = Nothing
        End If

        If strTemp <> "ר�ⲻ����" Then
            Select Case OutType
            Case 1
                If iUseCreateHTML = 1 Or iUseCreateHTML = 3 Then
                    strTemp = ChannelDir & "/" & rsInfo(3) & "Index.html"
                Else
                    strTemp = ChannelDir & "/ShowSpecial.asp?SpecialID=" & rsInfo(1)
                End If
            Case 2
                strTemp = rsInfo(2)
            Case 3
                If iUseCreateHTML = 1 Or iUseCreateHTML = 3 Then
                    strTemp = "<a href='" & InstallDir & ChannelDir & "/" & rsInfo(3) & "Index.html" & "'>" & rsInfo(2) & "</a>"
                Else
                    strTemp = "<a href='" & InstallDir & ChannelDir & "/ShowSpecial.asp?SpecialID=" & rsInfo(1) & "'>" & rsInfo(2) & "</a>"
                End If
            Case Else
                strTemp = "��ǩ������"
            End Select
            GetInfoSpecil = strTemp
        Else
            GetInfoSpecil = ""
        End If
    End If
    rsInfo.Close
    Set rsInfo = Nothing
End Function

'**************************************************
'��������GetItemPath
'��  �ã������Ŀ·��
'��  ����iStructureType ---- Ŀ¼�ṹ��ʽ
'        sParentDir ---- ����ĿĿ¼
'        sClassDir ---- ��ǰ��ĿĿ¼
'        UpdateTime ---- ��ĿĿ¼
'����ֵ�������Ŀ·��
'**************************************************
Public Function GetItemPath(iStructureType, sParentDir, sClassDir, UpdateTime)
    Select Case iStructureType
    Case 0      'Ƶ��/����/С��/�·�/�ļ�����Ŀ�ּ����ٰ��·ݱ��棩
        GetItemPath = sParentDir & sClassDir & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 1      'Ƶ��/����/С��/����/�ļ�����Ŀ�ּ����ٰ����ڷ֣�ÿ��һ��Ŀ¼��
        GetItemPath = sParentDir & sClassDir & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 2      'Ƶ��/����/С��/�ļ�����Ŀ�ּ������ٰ��·ݣ�
        GetItemPath = sParentDir & sClassDir & "/"
    Case 3      'Ƶ��/��Ŀ/�·�/�ļ�����Ŀƽ�����ٰ��·ݱ��棩
        GetItemPath = "/" & sClassDir & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 4      'Ƶ��/��Ŀ/����/�ļ�����Ŀƽ�����ٰ����ڷ֣�ÿ��һ��Ŀ¼��
        GetItemPath = "/" & sClassDir & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 5      'Ƶ��/��Ŀ/�ļ�����Ŀƽ�������ٰ��·ݣ�
        GetItemPath = "/" & sClassDir & "/"
    Case 6      'Ƶ��/�ļ���ֱ�ӷ���Ƶ��Ŀ¼�У�
        GetItemPath = "/"
    Case 7      'Ƶ��/HTML/�ļ���ֱ�ӷ���ָ���ġ�HTML���ļ����У�
        GetItemPath = "/HTML/"
    Case 8      'Ƶ��/���/�ļ���ֱ�Ӱ���ݱ��棬ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & "/"
    Case 9      'Ƶ��/�·�/�ļ���ֱ�Ӱ��·ݱ��棬ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 10     'Ƶ��/����/�ļ���ֱ�Ӱ����ڱ��棬ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 11     'Ƶ��/���/�·�/�ļ����Ȱ���ݣ��ٰ��·ݱ��棬ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 12     'Ƶ��/���/����/�ļ����Ȱ���ݣ��ٰ����ڷ֣�ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 13     'Ƶ��/�·�/����/�ļ����Ȱ��·ݣ��ٰ����ڷ֣�ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 14     'Ƶ��/���/�·�/����/�ļ����Ȱ���ݣ��ٰ����ڷ֣�ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    End Select
    GetItemPath = Replace(GetItemPath, "//", "/")
End Function

'**************************************************
'��������GetItemFileName
'��  �ã������Ŀ����
'��  ����iFileNameType ---- �ļ���������
'        sChannelDir ---- ��ǰƵ��Ŀ¼
'        UpdateTime ---- ����ʱ��
'        ItemID ---- ����ID��ArticleID/SoftID/PhotoID)
'����ֵ�������Ŀ����
'**************************************************
Public Function GetItemFileName(iFileNameType, sChannelDir, UpdateTime, ItemID)
    Select Case iFileNameType
    Case 0
        GetItemFileName = ItemID
    Case 1
        GetItemFileName = Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & Right("0" & Day(UpdateTime), 2) & Right("0" & Hour(UpdateTime), 2) & Right("0" & Minute(UpdateTime), 2) & Right("0" & Second(UpdateTime), 2)
    Case 2
        GetItemFileName = sChannelDir & "_" & ItemID
    Case 3
        GetItemFileName = sChannelDir & "_" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & Right("0" & Day(UpdateTime), 2) & Right("0" & Hour(UpdateTime), 2) & Right("0" & Minute(UpdateTime), 2) & Right("0" & Second(UpdateTime), 2)
    Case 4
        GetItemFileName = Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & Right("0" & Day(UpdateTime), 2) & Right("0" & Hour(UpdateTime), 2) & Right("0" & Minute(UpdateTime), 2) & Right("0" & Second(UpdateTime), 2) & "_" & ItemID
    Case 5
        GetItemFileName = sChannelDir & "_" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & Right("0" & Day(UpdateTime), 2) & Right("0" & Hour(UpdateTime), 2) & Right("0" & Minute(UpdateTime), 2) & Right("0" & Second(UpdateTime), 2) & "_" & ItemID
    End Select
End Function
'==================================================
'��������YN

'���ܣ�     �����жϺ���,���Ը���������������������������Ӧ�Ľ��
'condition: �����������,�������н��,������������Fir,�������Sec
'Fir:       ����������ʱ�����Fir������
'Sec :      ������������ʱ�����Sec������
'==================================================
Function YN(Condition, Fir, Sec)
    If Condition = "" Or IsNull(Condition) Then '�����жϲ���Ϊ��,�򷵻�Sec������
        YN = Sec
	Elseif LCase(Condition)="true" Then
	    YN=Fir 
	Elseif LCase(Condition)="false" Then
	    YN=Sec
    Else
        regEx.Pattern = "^[0-9\<\>\=\%\+\-\*\/\""]+$"    'ƥ��ֻ�����ֻ��������
        Dim Temp, result
        Temp = regEx.Test(Condition)  '�ж��Ƿ�ֻ�����ֺ������
        If Temp = True Then           '���ֻ�����ֺ������
		    Condition = Replace(Condition,"%"," mod ")
            result = Eval(Condition)  'ִ����������
            If (result) Then
                YN = Fir           '������Ϊ��,��������1
            Else
                YN = Sec             '������Ϊ��,��������2
            End If
        ElseIf InStr(Condition, "=") Then   '�ַ�����������ж�

            Dim Tempequal
            Tempequal = Split(Condition, "=")
            If Tempequal(0) = Tempequal(1) Then
                YN = Fir
            Else
                YN = Sec
            End If
        ElseIf InStr(Condition, "<>") Then   '�ַ������������ж�
            Dim Tempuneuqal
            Tempuneuqal = Split(Condition, "<>")
            If Tempuneuqal(0) <> Tempuneuqal(1) Then
                YN = Fir
            Else
                YN = Sec
            End If

        Else                            '������������óɷǷ�����
            YN = "�������Ͳ���ȷ"
        End If
    End If
End Function
%>
