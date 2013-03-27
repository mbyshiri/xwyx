<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

'**************************************************
'��������GetNumber_Option
'��  �ã���ʾ���������˵�
'��  ����MinNum ---- ��ʼ��
'        MaxNum ---- �����
'        CurrentNum ----selected Ĭ����
'����ֵ�������˵�����
'**************************************************
Public Function GetNumber_Option(MinNum, MaxNum, CurrentNum)
    Dim strNumber, i
    For i = MinNum To MaxNum
        If i = CurrentNum Then
            strNumber = strNumber & "<option value='" & i & "' selected>&nbsp;&nbsp;" & i & "&nbsp;&nbsp;</option>"
        Else
            strNumber = strNumber & "<option value='" & i & "'>&nbsp;&nbsp;" & i & "&nbsp;&nbsp;</option>"
        End If
    Next
    GetNumber_Option = strNumber
End Function

'**************************************************
'��������IsStyleDisplay
'��  �ã��Ƿ���ʾ��
'��  ����Compare1-----�Ƚ�ֵ1
'��  ����Compare2-----�Ƚ�ֵ2
'����ֵ���滻���ַ���
'**************************************************
Public Function IsStyleDisplay(ByVal Compare1, ByVal Compare2)
    If Compare1 = Compare2 Then
        IsStyleDisplay = " style='display:'"
    Else
        IsStyleDisplay = " style='display:none'"
    End If
End Function
'**************************************************
'��������RadioValue
'��  �ã���ʾ��ѡ����߶�ѡ���ֵ���ж��Ƿ�ѡ��
'��  ����compvalue ---- ѡ���Ŀǰʵ��ֵ
'        showvalue ---- ѡ�����ʾֵ
'**************************************************
Public Function RadioValue(compvalue, showvalue)
    If compvalue = showvalue Then
        RadioValue = "value='" & showvalue & "' checked"
    Else
        RadioValue = "value='" & showvalue & "'"
    End If
End Function

'**************************************************
'��������OptionValue
'��  �ã���ʾ�����б��ֵ���ж��Ƿ�ѡ��
'��  ����compvalue ---- ѡ���Ŀǰʵ��ֵ
'        showvalue ---- ѡ�����ʾֵ
'**************************************************
Public Function OptionValue(compvalue, showvalue)
    If compvalue = showvalue Then
        OptionValue = "value='" & showvalue & "' selected"
    Else
        OptionValue = "value='" & showvalue & "'"
    End If
End Function

'**************************************************
'��������GetPayOnlineProviderName
'��  �ã���ʾ����֧��������
'��  ����PayOnlineProviderID ---- ϵͳ����֧����ID
'����ֵ������֧��������
'**************************************************
Public Function GetPayOnlineProviderName(PayOnlineProviderID)
    Select Case PayOnlineProviderID
    Case 1, 10
        GetPayOnlineProviderName = "��������"
    Case 2
        GetPayOnlineProviderName = "�й�����֧����"
    Case 3
        GetPayOnlineProviderName = "�Ϻ���Ѹ"
    Case 4
        GetPayOnlineProviderName = "�㶫����"
    Case 5
        GetPayOnlineProviderName = "����֧��"
    Case 6
        GetPayOnlineProviderName = "�׸�ͨ"
    Case 7
        GetPayOnlineProviderName = "��������"
    Case 8, 12
        GetPayOnlineProviderName = "֧����"
    Case 9
        GetPayOnlineProviderName = "��Ǯ"
    Case 11
        GetPayOnlineProviderName = "��Ǯ������"
    Case 13
        GetPayOnlineProviderName = "�Ƹ�ͨ"
    End Select
End Function

'**************************************************
'��������GetArrItem
'��  �ã��õ�������ĳ��Ԫ�ص�ֵ
'��  ����arrTemp ---- Ҫȡ������
'        ItemIndex ---- �ڼ�λ��
'����ֵ������λ����Ԫ��
'**************************************************
Public Function GetArrItem(ByVal arrTemp, ByVal ItemIndex)
    If Not IsArray(arrTemp) Then
        GetArrItem = ""
        Exit Function
    End If
    ItemIndex = PE_CLng(ItemIndex)
    If ItemIndex < 0 Or ItemIndex > UBound(arrTemp) Then
        GetArrItem = ""
        Exit Function
    End If
    Dim strTemp
    strTemp = arrTemp(ItemIndex)
    If InStr(strTemp, "|") > 0 Then
        GetArrItem = Left(strTemp, InStr(strTemp, "|") - 1)
    Else
        GetArrItem = strTemp
    End If
End Function

'**************************************************
'��������Array2Option
'��  �ã��������������б���Ŀ
'��  ����arrTemp ---- ����
'        ItemIndex ---- ������Ĭ�ϵ�����
'����ֵ�������˵�
'**************************************************
Public Function Array2Option(ByVal arrTemp, ByVal ID)
    Dim strOption, i, arrValue
    strOption = "<option value='-1'> </option>"
    ID = PE_CLng(ID)
    For i = 0 To UBound(arrTemp)
        arrValue = Split(arrTemp(i), "|")
        If CLng(arrValue(1)) = 1 Then
            If ID > -1 Then
                If i = ID Then
                    strOption = strOption & "<option value='" & i & "' selected>" & arrValue(0) & "</option>"
                Else
                    strOption = strOption & "<option value='" & i & "'>" & arrValue(0) & "</option>"
                End If
            Else
                If CLng(arrValue(2)) = 1 Then
                    strOption = strOption & "<option value='" & i & "' selected>" & arrValue(0) & "</option>"
                Else
                    strOption = strOption & "<option value='" & i & "'>" & arrValue(0) & "</option>"
                End If
            End If
        End If
    Next
    Array2Option = strOption
End Function

'**************************************************
'��������GetArrFromDictionary
'��  �ã����ֵ��������ֵ
'��  ����strTableName ---- ������
'        strFieldName ---- ��������
'����ֵ����ѯ����ֵ
'**************************************************
Public Function GetArrFromDictionary(strTableName, strFieldName)
    Dim rsDictionary, strTemp
    Set rsDictionary = Conn.Execute("select FieldValue from PE_Dictionary where TableName='" & strTableName & "' and FieldName='" & strFieldName & "'")
    If rsDictionary.BOF And rsDictionary.EOF Then
        strTemp = ""
    Else
        strTemp = rsDictionary(0) & ""
    End If
    Set rsDictionary = Nothing
    GetArrFromDictionary = Split(strTemp, "$$$")
End Function

'**************************************************
'��������PopCalendarInit
'��  �ã���������js
'**************************************************
Public Sub PopCalendarInit()
    Response.Write "<script language='JavaScript' src='PopCalendar.js'></script>" & vbCrLf
    Response.Write "<script language='JavaScript'>" & vbCrLf
    Response.Write "    PopCalendar = getCalendarInstance()" & vbCrLf
    Response.Write "    PopCalendar.startAt = 0 // 0 - sunday ; 1 - monday" & vbCrLf
    Response.Write "    PopCalendar.showWeekNumber = 0 // 0 - don't show; 1 - show" & vbCrLf
    Response.Write "    PopCalendar.showTime = 0 // 0 - don't show; 1 - show" & vbCrLf
    Response.Write "    PopCalendar.showToday = 0 // 0 - don't show; 1 - show" & vbCrLf
    Response.Write "    PopCalendar.showWeekend = 1 // 0 - don't show; 1 - show" & vbCrLf
    Response.Write "    PopCalendar.showHolidays = 1 // 0 - don't show; 1 - show" & vbCrLf
    Response.Write "    PopCalendar.showSpecialDay = 1 // 0 - don't show, 1 - show" & vbCrLf
    Response.Write "    PopCalendar.selectWeekend = 0 // 0 - don't Select; 1 - Select" & vbCrLf
    Response.Write "    PopCalendar.selectHoliday = 0 // 0 - don't Select; 1 - Select" & vbCrLf
    Response.Write "    PopCalendar.addCarnival = 0 // 0 - don't Add; 1- Add to Holiday" & vbCrLf
    Response.Write "    PopCalendar.addGoodFriday = 0 // 0 - don't Add; 1- Add to Holiday" & vbCrLf
    Response.Write "    PopCalendar.language = 0 // 0 - Chinese; 1 - English" & vbCrLf
    Response.Write "    PopCalendar.defaultFormat = 'yyyy-mm-dd' //Default Format dd-mm-yyyy" & vbCrLf
    Response.Write "    PopCalendar.fixedX = -1 // x position (-1 if to appear below control)" & vbCrLf
    Response.Write "    PopCalendar.fixedY = -1 // y position (-1 if to appear below control)" & vbCrLf
    Response.Write "    PopCalendar.fade = .5 // 0 - don't fade; .1 to 1 - fade (Only IE) " & vbCrLf
    Response.Write "    PopCalendar.shadow = 1 // 0  - don't shadow, 1 - shadow" & vbCrLf
    Response.Write "    PopCalendar.move = 1 // 0  - don't move, 1 - move (Only IE)" & vbCrLf
    Response.Write "    PopCalendar.saveMovePos = 1  // 0  - don't save, 1 - save" & vbCrLf
    Response.Write "    PopCalendar.centuryLimit = 40 // 1940 - 2039" & vbCrLf
    Response.Write "    PopCalendar.initCalendar()" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub



'**************************************************
'��������ShowJS_Main
'��  �ã�ҳ�����js(����ڲȫѡ,ɾ����ʾ)
'��  ����ItemName ---- ��Ŀ����
'����ֵ��javascript ��֤
'**************************************************
Public Sub ShowJS_Main(ItemName)
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
    Response.Write "     if(confirm('ȷ��Ҫɾ��ѡ�е�" & ItemName & "��'))" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "     else" & vbCrLf
    Response.Write "         return false;" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
End Sub

'**************************************************
'��������ShowJS_Manage
'��  �ã�ͨ��Ƶ������js��֤(����ڲȫѡ,ɾ����ʾ,�ƶ�)
'��  ����ItemName ---- ��Ŀ����
'����ֵ��javascript ��֤
'**************************************************
Public Sub ShowJS_Manage(ItemName)
    Dim strJS
    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "function CheckItem(CB){" & vbCrLf
    Response.Write "  var tagname=(arguments.length>1)?arguments[1]:'TR';" & vbCrLf
    Response.Write "  if(document.myform.chkAll.checked){" & vbCrLf
    Response.Write "    document.myform.chkAll.checked = document.myform.chkAll.checked&0;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (CB.checked){hL(CB,tagname)};else{dL(CB,tagname)};" & vbCrLf
    Response.Write "  var TB=TO=0;" & vbCrLf
    Response.Write "  for (var i=0;i<myform.elements.length;i++) {" & vbCrLf
    Response.Write "    var e=myform.elements[i];" & vbCrLf
    Response.Write "    if ((e.name != 'chkAll') && (e.type=='checkbox')) {" & vbCrLf
    Response.Write "      TB++;" & vbCrLf
    Response.Write "      if (e.checked) TO++;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  myform.chkAll.checked=(TO==TB)?true:false;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckAll(form){" & vbCrLf
    Response.Write "  var tagname=(arguments.length>1)?arguments[1]:'TR';" & vbCrLf
    Response.Write "  for (var i=0;i<form.elements.length;i++){" & vbCrLf
    Response.Write "    var e = form.elements[i];" & vbCrLf
    Response.Write "    if (e.name != 'chkAll' && e.disabled == false && e.type == 'checkbox') {" & vbCrLf
    Response.Write "      e.checked = form.chkAll.checked;" & vbCrLf
    Response.Write "      if (e.checked){hL(e,tagname)};else{dL(e,tagname)};" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function hL(E,tagname){" & vbCrLf
    Response.Write "  while (E.tagName!=tagname) {E=E.parentElement;}" & vbCrLf
    Response.Write "  E.className='tdbg2';" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function dL(E,tagname){" & vbCrLf
    Response.Write "  while (E.tagName!=tagname) {E=E.parentElement;}" & vbCrLf
    Response.Write "  E.className='tdbg';" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function ConfirmDel(){" & vbCrLf
    Response.Write " if(document.myform.Action.value=='Del'){" & vbCrLf
    Response.Write "     if(confirm('ȷ��Ҫɾ��ѡ�е�" & ItemName & "�𣿱���������ѡ�е�" & ItemName & "�Ƶ�����վ�С���Ҫʱ���ɴӻ���վ�лָ���'))" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "     else" & vbCrLf
    Response.Write "         return false;" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write " else if(document.myform.Action.value=='ConfirmDel'){" & vbCrLf
    Response.Write "     if(confirm('ȷ��Ҫ����ɾ��ѡ�е�" & ItemName & "�𣿳���ɾ���󽫲��ָܻ���'))" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "     else" & vbCrLf
    Response.Write "         return false;" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write " else if(document.myform.Action.value=='ClearRecyclebin'){" & vbCrLf
    Response.Write "     if(confirm('ȷ��Ҫ��ջ���վ��һ����ս����ָܻ���'))" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "     else" & vbCrLf
    Response.Write "         return false;" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write " else if(document.myform.Action.value=='DelFromSpecial'){" & vbCrLf
    Response.Write "     if(confirm('ȷ��Ҫ��ѡ�е�" & ItemName & "��������ר����ɾ���𣿲����ɹ���" & ItemName & "���������κ�ר�⡣'))" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "     else" & vbCrLf
    Response.Write "         return false;" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
End Sub

Sub ShowContentManagePath(RootName)
    Response.Write "�����ڵ�λ�ã�&nbsp;" & ChannelName & "����&nbsp;&gt;&gt;&nbsp;<a href='" & FileName & "'>" & RootName & "</a>&nbsp;&gt;&gt;&nbsp;"
    If ClassID > 0 Then
        If ParentID > 0 Then
            Dim sqlPath, rsPath
            sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
            Set rsPath = Conn.Execute(sqlPath)
            Do While Not rsPath.EOF
                Response.Write "<a href='" & FileName & "&ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
                rsPath.MoveNext
            Loop
            rsPath.Close
            Set rsPath = Nothing
        End If
        Response.Write "<a href='" & FileName & "&ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
    End If
    If ManageType = "My" Then
        Response.Write "<font color=red>" & AdminName & "</font> ��ӵ�" & ChannelShortName & ""
    Else
        If Keyword = "" Then
            Select Case Status
            Case -2
                Response.Write "�˸�"
            Case -1
                Response.Write "�ݸ�"
            Case 0
                Response.Write "����˵�" & ChannelShortName & "��"
            Case 1
                Response.Write "����˵�" & ChannelShortName & "��"
            Case Else
                Response.Write "����" & ChannelShortName & "��"
            End Select
        Else
            Select Case strField
            Case "ID"
                Response.Write "ID����" & Keyword & "</font> "
            Case "Title"
                Response.Write "�����к��� <font color=red>" & Keyword & "</font> "
            Case "Content"
                Response.Write "�����к��� <font color=red>" & Keyword & "</font> "
            Case "Author"
                Response.Write "���������к��� <font color=red>" & Keyword & "</font> "
            Case "Inputer"
                Response.Write "<font color=red>" & Keyword & "</font> ���"
            Case "Editor"
                Response.Write "<font color=red>" & Keyword & "</font> ���"
            Case "Keyword"
                Response.Write "�ؼ���Ϊ <font color=red>" & Keyword & "</font> "
            Case "UpdateTime"
                Response.Write "����ʱ��Ϊ <font color=red>" & Keyword & "</font> "
            Case "SoftName", "PhotoName"
                Response.Write "�����к��� <font color=red>" & Keyword & "</font> "
            Case "SoftIntro", "PhotoIntro"
                Response.Write "�����к��� <font color=red>" & Keyword & "</font> "
            Case Else
                Response.Write "�����к��� <font color=red>" & Keyword & "</font> "
            End Select
            Select Case Status
            Case -2
                Response.Write "���˸�"
            Case -1
                Response.Write "�Ĳݸ�"
            Case 0
                Response.Write "����δ��˵�" & ChannelShortName & "��"
            Case 1
                Response.Write "��������˵�" & ChannelShortName & "��"
            Case Else
                Response.Write "��" & ChannelShortName & "��"
            End Select
        End If
    End If
End Sub

'**************************************************
'��������GetRootClass
'��  �ã���ʾ��Ŀ������
'��  ����ChannelID ---- Ƶ��ID
'        RootID ---- ����ĿID
'        FileName ---- ��Ŀ�ļ���
'����ֵ����Ŀ������
'**************************************************
Public Function GetRootClass()
    Dim sqlRoot, rsRoot, strRoot
    sqlRoot = "select ClassID,ClassName,RootID,Child from PE_Class where ChannelID=" & ChannelID & " and ParentID=0 and ClassType=1 order by RootID"
    Set rsRoot = Conn.Execute(sqlRoot)
    If rsRoot.BOF And rsRoot.EOF Then
        strRoot = "��û���κ���Ŀ�������������Ŀ��"
    Else
        strRoot = "|&nbsp;"
        Do While Not rsRoot.EOF
            If rsRoot(2) = RootID Then
                strRoot = strRoot & "<a href='" & FileName & "&ClassID=" & rsRoot(0) & "'><font color=red>" & rsRoot(1) & "</font></a> | "
            Else
                strRoot = strRoot & "<a href='" & FileName & "&ClassID=" & rsRoot(0) & "'>" & rsRoot(1) & "</a> | "
            End If
            rsRoot.MoveNext
        Loop
    End If
    rsRoot.Close
    Set rsRoot = Nothing
    GetRootClass = strRoot
End Function

'**************************************************
'��������GetChild_Root
'��  �ã���ʾ��Ŀ����Ŀ������
'��  ����ChannelID ---- Ƶ��ID
'        RootID ---- ����ĿID
'        ClassID ---- ��ĿID
'        ParentPath ---- ��·��
'        Depth ---- ��Ŀ���
'        FileName ---- ��Ŀ�ļ���
'����ֵ������Ŀ������
'**************************************************
Public Function GetChild_Root()
    Dim sqlChild, rsChild, arrParentPath, isCurrent, strChild, i
    If RootID <= 0 Then
        GetChild_Root = ""
        Exit Function
    End If
    sqlChild = "select ClassID,ClassName,Child from PE_Class where ChannelID=" & ChannelID & " and Depth=1 and RootID=" & RootID & " order by OrderID"
    Set rsChild = Conn.Execute(sqlChild)
    If Not (rsChild.BOF And rsChild.EOF) Then
        i = 1
        arrParentPath = Split(ParentPath, ",")
        strChild = "<tr class='tdbg'><td>"
        Do While Not rsChild.EOF
            If Depth <= 1 Then
                If rsChild(0) = ClassID Then
                    isCurrent = True
                Else
                    isCurrent = False
                End If
            Else
                If PE_CLng(arrParentPath(2)) = rsChild(0) Then
                    isCurrent = True
                Else
                    isCurrent = False
                End If
            End If
            If isCurrent = True Then
                strChild = strChild & "&nbsp;&nbsp;<a href='" & FileName & "&ClassID=" & rsChild(0) & "'><font color='red'>" & rsChild(1) & "</font></a>"
            Else
                strChild = strChild & "&nbsp;&nbsp;<a href='" & FileName & "&ClassID=" & rsChild(0) & "'>" & rsChild(1) & "</a>"
            End If
            If rsChild(2) > 0 Then
                strChild = strChild & "(" & rsChild(2) & ")"
            End If
            If i Mod 8 = 0 Then
                strChild = strChild & "<br>"
            Else
                strChild = strChild & "&nbsp;&nbsp;"
            End If
            rsChild.MoveNext
            i = i + 1
        Loop
        strChild = strChild & "</td></tr>"
    End If
    rsChild.Close
    Set rsChild = Nothing
    GetChild_Root = strChild
End Function

Function GetSpecial_Option(SpecialID)
    Dim sqlSpecial, rsSpecial, strOption, strOptionTemp
    sqlSpecial = "select ChannelID,SpecialID,SpecialName,OrderID from PE_Special where ChannelID=0 or ChannelID=" & ChannelID & "   order by ChannelID,OrderID"
    Set rsSpecial = Conn.Execute(sqlSpecial)
    Do While Not rsSpecial.EOF
        If rsSpecial("ChannelID") > 0 Then
            strOptionTemp = rsSpecial("SpecialName") & "����Ƶ����"
        Else
            strOptionTemp = rsSpecial("SpecialName") & "��ȫվ��"
        End If
        If FoundInArr(SpecialID, rsSpecial("SpecialID"), ",") = True Then
            strOption = strOption & "<option value='" & rsSpecial("SpecialID") & "' selected>" & strOptionTemp & "</option>"
        Else
            strOption = strOption & "<option value='" & rsSpecial("SpecialID") & "'>" & strOptionTemp & "</option>"
        End If
        rsSpecial.MoveNext
    Loop
    rsSpecial.Close
    Set rsSpecial = Nothing
    GetSpecial_Option = strOption
End Function

'**************************************************
'��������GetStars
'��  �ã���ʾ�ȼ�������
'��  ����Stars ---- ��Ŀ����
'����ֵ�������˵�����
'**************************************************
Public Function GetStars(Stars)
    Dim strTemp
    strTemp = strTemp & "<option value='5'"
    If Stars = 5 Then strTemp = strTemp & " selected"
    strTemp = strTemp & ">������</option>"
    strTemp = strTemp & "<option value='4'"
    If Stars = 4 Then strTemp = strTemp & " selected"
    strTemp = strTemp & ">�����</option>"
    strTemp = strTemp & "<option value='3'"
    If Stars = 3 Then strTemp = strTemp & " selected"
    strTemp = strTemp & ">����</option>"
    strTemp = strTemp & "<option value='2'"
    If Stars = 2 Then strTemp = strTemp & " selected"
    strTemp = strTemp & ">���</option>"
    strTemp = strTemp & "<option value='1'"
    If Stars = 1 Then strTemp = strTemp & " selected"
    strTemp = strTemp & ">��</option>"
    strTemp = strTemp & "<option value='0'"
    If Stars = 0 Then strTemp = strTemp & " selected"
    strTemp = strTemp & ">��</option>"
    GetStars = strTemp
End Function

'**************************************************
'��������GetAuthorList
'��  �ã���ʾ����
'��  ����ChannelID ---- Ƶ��ID
'        UserName ---- �û�����
'����ֵ����δ֪����������������Ա�������ࡿ
'**************************************************
Public Function GetAuthorList(FilePrefix, ChannelID, UserName)
    Dim Author, strAuthorList

    Author = Trim(Session("Author"))
    Dim strDefAuthor, strUnKnowAuthor
    strDefAuthor = XmlText("BaseText", "DefAuthor", "����")
    strUnKnowAuthor = XmlText("BaseText", "UnKnowAuthor", "δ֪")

    strAuthorList = "<font color='blue'><="
    strAuthorList = strAuthorList & "��<font color='green' onclick=""document.myform.Author.value='" & strDefAuthor & "'"" style=""cursor:hand;"">" & strDefAuthor & "</font>��"
    strAuthorList = strAuthorList & "��<font color='green' onclick=""document.myform.Author.value='" & strUnKnowAuthor & "'"" style=""cursor:hand;"">" & strUnKnowAuthor & "</font>��"
    strAuthorList = strAuthorList & "��<font color='green' onclick=""document.myform.Author.value='" & UserName & "'"" style=""cursor:hand;"">" & UserName & "</font>��"
    If Author <> "" And Author <> strDefAuthor And Author <> strUnKnowAuthor And Author <> UserName Then
        strAuthorList = strAuthorList & "��<font color='green' onclick=""document.myform.Author.value='" & FilterJS(Replace(Author, "'", "")) & "'"" style=""cursor:hand;"">" & FilterJS(Author) & "</font>��"
    End If
    strAuthorList = strAuthorList & "��<font color='green' onclick=""window.open('" & FilePrefix & "_SourceList.asp?TypeSelect=AuthorList&ChannelID=" & ChannelID & "', 'AuthorList', 'width=600,height=450,resizable=0,scrollbars=yes');"" style=""cursor:hand;"">����</font>��"
    strAuthorList = strAuthorList & "</font>"
    GetAuthorList = strAuthorList
End Function

'**************************************************
'��������GetCopyFromList
'��  �ã���ʾ��Դ
'��  ����FilePrefix ----������� Admin,User
'        ChannelID ---- Ƶ��ID
'����ֵ��<=����վԭ���������ࡿ
'**************************************************
Public Function GetCopyFromList(FilePrefix, ChannelID)
    Dim CopyFrom, strCopyFromList
    CopyFrom = Trim(Session("CopyFrom"))
    Dim strDefCopyFrom
    strDefCopyFrom = XmlText("BaseText", "DefCopyFrom", "��վԭ��")

    strCopyFromList = "<font color='blue'><="
    strCopyFromList = strCopyFromList & "��<font color='green' onclick=""document.myform.CopyFrom.value='" & strDefCopyFrom & "'"" style=""cursor:hand;"">" & strDefCopyFrom & "</font>��"
    If CopyFrom <> "" And CopyFrom <> strDefCopyFrom Then
        strCopyFromList = strCopyFromList & "��<font color='green' onclick=""document.myform.CopyFrom.value='" & FilterJS(Replace(CopyFrom, "'", "")) & "'"" style=""cursor:hand;"">" & FilterJS(CopyFrom) & "</font>��"
    End If
    strCopyFromList = strCopyFromList & "��<font color='green' onclick=""window.open('" & FilePrefix & "_SourceList.asp?TypeSelect=CopyFromList&ChannelID=" & ChannelID & "', 'CopyFromList', 'width=600,height=450,resizable=0,scrollbars=yes');"" style=""cursor:hand;"">����</font>��"
    strCopyFromList = strCopyFromList & "</font>"
    GetCopyFromList = strCopyFromList
End Function

'**************************************************
'��������GetKeywordList
'��  �ã���ʾ�ؼ���
'��  ����FilePrefix ----������� Admin,User
'        ChannelID ---- Ƶ��ID
'����ֵ����ʾƵ����ǰ4���ؼ��� +�����ࡿ
'**************************************************
Public Function GetKeywordList(FilePrefix, ChannelID)
    Dim sqlGetKey, rsGetKey, strKeywordList
    strKeywordList = "<font color='blue'><="
    sqlGetKey = "select top 4 * from PE_NewKeys where ChannelID=" & ChannelID & " or ChannelID=0 order by LastUseTime Desc"
    Set rsGetKey = Conn.Execute(sqlGetKey)
    If rsGetKey.BOF And rsGetKey.EOF Then
        strKeywordList = strKeywordList & "��<font color='green'>�޳��ùؼ���</font>��"
    Else
        Do While Not rsGetKey.EOF
            strKeywordList = strKeywordList & "��<font color='green' onclick=""document.myform.Keyword.value+=(document.myform.Keyword.value==''?'':'|')+'" & FilterJS(Replace(rsGetKey("KeyText"), "'", "")) & "'"" style=""cursor:hand;"">" & FilterJS(rsGetKey("KeyText")) & "</font>��"
            rsGetKey.MoveNext
        Loop
    End If
    rsGetKey.Close
    Set rsGetKey = Nothing
    strKeywordList = strKeywordList & "��<font color='green' onclick=""window.open('" & FilePrefix & "_SourceList.asp?ChannelID=" & ChannelID & "&TypeSelect=KeyList', 'KeyList', 'width=600,height=450,resizable=0,scrollbars=yes');"" style=""cursor:hand;"">����</font>��"
    strKeywordList = strKeywordList & "</font>"
    GetKeywordList = strKeywordList
End Function

'**************************************************
'��������SaveKeyword
'��  �ã�����ؼ���
'**************************************************
Public Sub SaveKeyword(strKeyword)
    Dim rsKeyword, sqlKeyword, arrKeyword, i
    strKeyword = ReplaceBadChar(strKeyword) 
    If strKeyword = "" Then Exit Sub
    arrKeyword = Split(strKeyword, "|")
    Set rsKeyword = Server.CreateObject("adodb.recordset")
    For i = 0 To UBound(arrKeyword)
        sqlKeyword = "Select ChannelID,KeyText,Hits,LastUseTime from PE_NewKeys Where ChannelID=" & ChannelID & " and KeyText='" & arrKeyword(i) & "'"
        rsKeyword.Open sqlKeyword, Conn, 1, 3
        If rsKeyword.BOF And rsKeyword.EOF Then
            If 	arrKeyword(i)<>"" then	
                rsKeyword.addnew
                rsKeyword("ChannelID") = ChannelID
                rsKeyword("KeyText") = arrKeyword(i)
                rsKeyword("Hits") = 0
                rsKeyword("LastUseTime") = Now()
                rsKeyword.Update
            End If				
        Else
            Do While Not rsKeyword.EOF
                If arrKeyword(i)<>"" then				
                    rsKeyword("Hits") = rsKeyword("Hits") + 1
                    rsKeyword("LastUseTime") = Now()
                    rsKeyword.Update
                End If 
                rsKeyword.MoveNext				
            Loop
        End If
        rsKeyword.Close
    Next
    Set rsKeyword = Nothing
End Sub

'**************************************************
'��������ReplaceText
'��  �ã����˷Ƿ��ַ���
'��  ����iText-----�����ַ���
'����ֵ���滻���ַ���
'**************************************************
Function ReplaceText(iText, iType)
    Dim rText, rsKey, sqlKey, i, Keyrow, Keycol
    If PE_Cache.GetValue("Site_ReplaceText") = "" Then
        Set rsKey = Server.CreateObject("Adodb.RecordSet")
        sqlKey = "Select Source,ReplaceText,OpenType,ReplaceType,Priority from PE_KeyLink where isUse=1 and LinkType=1 order by Priority"
        rsKey.Open sqlKey, Conn, 1, 1
        If Not (rsKey.BOF And rsKey.EOF) Then
            PE_Cache.SetValue "Site_ReplaceText", rsKey.GetString(, , "|||", "@@@", "")
            rsKey.Close
            Set rsKey = Nothing
        Else
            rsKey.Close
            Set rsKey = Nothing
            ReplaceText = iText
            Exit Function
        End If
    End If
    rText = iText
    Keyrow = Split(PE_Cache.GetValue("Site_ReplaceText"), "@@@")
    For i = 0 To UBound(Keyrow) - 1
        Keycol = Split(Keyrow(i), "|||")
        If Int(Keycol(3)) = 0 Or Int(Keycol(3)) = iType Then rText = PE_Replace(rText, Keycol(0), Keycol(1))
    Next
    ReplaceText = rText
End Function

'**************************************************
'��������ShowClassPath
'��  �ã���ʾ��Ŀ·��
'��  ������
'����ֵ����ʾ��Ŀ
'**************************************************
Public Function ShowClassPath()
    If ParentPath = "" Or IsNull(ParentPath) Then
        ShowClassPath = "�������κ���Ŀ"
        Exit Function
    End If
    Dim strPath
    If Depth > 0 Then
        Dim rsPath
        Set rsPath = Conn.Execute("select * from PE_Class where ClassID in (" & ParentPath & ") order by Depth")
        Do While Not rsPath.EOF
            strPath = strPath & rsPath("ClassName") & " >> "
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    strPath = strPath & ClassName
    ShowClassPath = strPath
End Function



'**************************************************
'��������GetUserGroup
'��  �ã���ʾ�û��鵼��
'��  ����arrGroupID ---- ָ��Ĭ���û���
'����ֵ���û����񵼺�
'**************************************************
Function GetUserGroup(arrGroupID, strDisabled)
    If IsNull(arrGroupID) Then Exit Function
    Dim rsGroup, strGroup, i
    strGroup = "<table width='95%' align='right'><tr>"
    Set rsGroup = Conn.Execute("select GroupID,GroupName from PE_UserGroup where GroupID<>-1 order by GroupType asc,GroupID asc")
    Do While Not rsGroup.EOF
        strGroup = strGroup & "<td><input type='checkbox' name='GroupID' value='" & rsGroup(0) & "'" & strDisabled
        If FoundInArr(arrGroupID, rsGroup(0), ",") = True Then
            strGroup = strGroup & " checked"
        End If
        strGroup = strGroup & ">" & rsGroup(1) & "</td>"
        i = i + 1
        rsGroup.MoveNext
        If i Mod 5 = 0 And Not rsGroup.EOF Then
            strGroup = strGroup & "</tr><tr>"
        End If
    Loop
    rsGroup.Close
    Set rsGroup = Nothing
    strGroup = strGroup & "</table>"
    GetUserGroup = strGroup
End Function

Sub UpdateChannelData(ByVal iChannelID)
    Dim rsChannel, sqlChannel, trs, ModuleName
    Dim ItemCount, ItemChecked, CommentCount, SpecialCount
    sqlChannel = "select ChannelID,ModuleType,ItemCount,ItemChecked,CommentCount,SpecialCount from PE_Channel"
    Set rsChannel = Server.CreateObject("Adodb.RecordSet")
    If IsValidID(iChannelID) = False Then
        iChannelID = ""
    End If
    
    If InStr(iChannelID, ",") > 0 Then
        sqlChannel = sqlChannel & " where ChannelID in (" & iChannelID & ")"
    ElseIf PE_CLng(iChannelID) > 0 Then
        sqlChannel = sqlChannel & " where ChannelID=" & iChannelID & ""
    Else
        sqlChannel = sqlChannel & " where ChannelType<=1 order by ChannelID"
    End If
    rsChannel.Open sqlChannel, Conn, 1, 3
    Do While Not rsChannel.EOF
        Select Case rsChannel("ModuleType")
        Case 7
            Dim i, HouseTableName
            For i = 1 To 5
                Select Case i
                Case 1
                    HouseTableName = "PE_HouseCS"
                Case 2
                    HouseTableName = "PE_HouseCZ"
                Case 3
                    HouseTableName = "PE_HouseQG"
                Case 4
                    HouseTableName = "PE_HouseQZ"
                Case 5
                    HouseTableName = "PE_HouseHZ"
                End Select
                Set trs = Conn.Execute("select Count(HouseID) from " & HouseTableName & " where Deleted=" & PE_False & "")
                ItemCount = ItemCount + trs(0)
                Set trs = Nothing
                Set trs = Conn.Execute("select Count(HouseID) from " & HouseTableName & " where Passed=" & PE_True & " and Deleted=" & PE_False & "")
                ItemChecked = ItemChecked + trs(0)
                Set trs = Nothing
            Next
            rsChannel("ItemCount") = ItemCount
            rsChannel("ItemChecked") = ItemChecked
        Case 8
            Set trs = Conn.Execute("select Count(PositionID) from PE_Position ")
            ItemCount = ItemCount + trs(0)
            Set trs = Nothing
            rsChannel("ItemCount") = ItemCount
        Case Else
            Select Case rsChannel("ModuleType")
            Case 1
                ModuleName = "Article"
            Case 2
                ModuleName = "Soft"
            Case 3
                ModuleName = "Photo"
            Case 5
                ModuleName = "Product"
            Case 6
                ModuleName = "Supply"
            End Select
            Set trs = Conn.Execute("select Count(" & ModuleName & "ID) from PE_" & ModuleName & " where ChannelID=" & rsChannel("ChannelID") & " and Deleted=" & PE_False & "")
            ItemCount = trs(0)
            Set trs = Nothing

            If ModuleName = "Product" Then
                Set trs = Conn.Execute("select Count(" & ModuleName & "ID) from PE_" & ModuleName & " where ChannelID=" & rsChannel("ChannelID") & " and EnableSale=" & PE_True & " and Deleted=" & PE_False & "")
            Else
                Set trs = Conn.Execute("select Count(" & ModuleName & "ID) from PE_" & ModuleName & " where ChannelID=" & rsChannel("ChannelID") & " and Status=3 and Deleted=" & PE_False & "")
            End If
            ItemChecked = trs(0)
            Set trs = Nothing

            Set trs = Conn.Execute("select Count(CommentID) from PE_Comment C inner join PE_" & ModuleName & " I on C.InfoID=I." & ModuleName & "ID where I.ChannelID=" & rsChannel("ChannelID") & "")
            CommentCount = trs(0)
            Set trs = Nothing

            Set trs = Conn.Execute("select Count(SpecialID) from PE_Special where ChannelID=" & rsChannel("ChannelID") & "")
            SpecialCount = trs(0)
            Set trs = Nothing

            rsChannel("ItemCount") = ItemCount
            rsChannel("ItemChecked") = ItemChecked
            rsChannel("CommentCount") = CommentCount
            rsChannel("SpecialCount") = SpecialCount
        End Select
        rsChannel.Update
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
End Sub

Sub UpdateUserData(UserType, UserName, BeginID, EndID)
    Dim sqlUser, rsUser, trs, PostItems, PassedItems
    If UserType = 0 Then
        If InStr(UserName, ",") > 0 Then
            sqlUser = "select * from PE_User where UserName in ('" & Replace(UserName, ",", "','") & "')"
        Else
            sqlUser = "select * from PE_User where UserName='" & UserName & "'"
        End If
    Else
        sqlUser = "select * from PE_User where UserID>=" & BeginID & " and UserID<=" & EndID
    End If
    Set rsUser = Server.CreateObject("Adodb.RecordSet")
    rsUser.Open sqlUser, Conn, 1, 3
    Do While Not rsUser.EOF
        Set trs = Conn.Execute("select count(ArticleID) from PE_Article where Deleted=" & PE_False & " and  Inputer='" & rsUser("UserName") & "'")
        If IsNull(trs(0)) Then
            PostItems = 0
        Else
            PostItems = trs(0)
        End If
        Set trs = Nothing
        Set trs = Conn.Execute("select count(ArticleID) from PE_Article where Deleted=" & PE_False & " and Status=3 and Inputer='" & rsUser("UserName") & "'")
        If IsNull(trs(0)) Then
            PassedItems = 0
        Else
            PassedItems = trs(0)
        End If
        Set trs = Nothing
        Set trs = Conn.Execute("select count(SoftID) from PE_Soft where Deleted=" & PE_False & " and  Inputer='" & rsUser("UserName") & "'")
        If Not IsNull(trs(0)) Then PostItems = PostItems + trs(0)
        Set trs = Nothing
        Set trs = Conn.Execute("select count(SoftID) from PE_Soft where Deleted=" & PE_False & " and  Status=3 and Inputer='" & rsUser("UserName") & "'")
        If Not IsNull(trs(0)) Then PassedItems = PassedItems + trs(0)
        Set trs = Nothing
        Set trs = Conn.Execute("select count(PhotoID) from PE_Photo where Deleted=" & PE_False & " and  Inputer='" & rsUser("UserName") & "'")
        If Not IsNull(trs(0)) Then PostItems = PostItems + trs(0)
        Set trs = Nothing
        Set trs = Conn.Execute("select count(PhotoID) from PE_Photo where Deleted=" & PE_False & " and  Status=3 and Inputer='" & rsUser("UserName") & "'")
        If Not IsNull(trs(0)) Then PassedItems = PassedItems + trs(0)
        Set trs = Nothing
        
        rsUser("PostItems") = PostItems
        rsUser("PassedItems") = PassedItems
        rsUser.Update
        rsUser.MoveNext
    Loop
    rsUser.Close
    Set rsUser = Nothing
End Sub

%>
