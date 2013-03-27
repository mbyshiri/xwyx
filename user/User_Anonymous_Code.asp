<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.Cache.asp"-->
<!--#include file="../Include/PowerEasy.Channel.asp"-->
<!--#include file="../Include/PowerEasy.Common.Front.asp"-->
<!--#include file="../Include/PowerEasy.Common.Manage.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************
If ShowAnonymous = False Then
    Call WriteErrMsg("��վδ��������Ͷ�幦��",ComeUrl)
    response.end
End If
Dim IsUpload
'���û�����Ӧ������ֵ
Sub GetUserAmon()
    If ShowAnonymous = False Then
        Call WriteErrMsg("��վδ��������Ͷ�幦��",ComeUrl)
        Exit Sub    
    End If	
    Dim  rsGroup
    Set rsGroup = Conn.Execute("select * from PE_UserGroup where GroupID=-1")
    GroupName = rsGroup("GroupName")
    GroupType = rsGroup("GroupType")
    arrClass_Browse = Trim(rsGroup("arrClass_Browse"))
    arrClass_View = Trim(rsGroup("arrClass_View"))
    arrClass_Input = Trim(rsGroup("arrClass_Input"))
    UserSetting = Split(Trim(rsGroup("GroupSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
    rsGroup.Close
    Set rsGroup = Nothing
    NeedlessCheck = PE_CLng(UserSetting(1))
    EnableModifyDelete = PE_CLng(UserSetting(2))
    MaxPerDay = PE_CLng(UserSetting(3))
    PresentExpTimes = PE_CDbl(UserSetting(4))
    MaxSendNum = PE_CLng(UserSetting(7))
    MaxFavorite = PE_CLng(UserSetting(8))
    Discount_Member = PE_CDbl(UserSetting(11))
    UserEnableComment = PE_CBool(UserSetting(5))
    UserCheckComment = PE_CBool(UserSetting(6))
    If PE_CBool(PE_CLng(UserSetting(9))) = True and  ShowAnonymous = True Then
        IsUpload = True
    Else
 	    IsUpload = False	
    End If	       
    UserChargeType = PE_CLng(UserSetting(14))
End Sub

Sub GetClass()
	ClassName = ""
	RootID = 0
	ParentID = 0
	Depth = 0
	ParentPath = "0"
	Child = 0
	arrChildID = ""
    If ClassID > 0 Then
        Dim tClass
        Set tClass = Conn.Execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,arrChildID from PE_Class where ClassID=" & ClassID)
        If tClass.BOF And tClass.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�Ҳ���ָ������Ŀ��</li>"
        Else
            ClassName = tClass(0)
            RootID = tClass(1)
            ParentID = tClass(2)
            Depth = tClass(3)
            ParentPath = tClass(4)
            Child = tClass(5)
            arrChildID = tClass(6)
        End If
        Set tClass = Nothing
    End If
End Sub
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
'��������User_GetClass_Option
'��  �ã���ʾ�û���Ŀ�����˵�
'��  ����ShowType ----��ʾ����
'        CurrentID ----��ǰ��ĿID
'����ֵ���û���Ŀ�����˵�
'**************************************************
Function User_GetClass_Option(ShowType, CurrentID)
    Dim rsClass, sqlClass, strClass_Option, tmpDepth, i, ClassNum
    Dim arrShowLine(20)
    Dim CheckParentPath, PurviewChecked
    
    ClassNum = 1
    CurrentID = PE_CLng(CurrentID)
    
    sqlClass = "Select * from PE_Class where ChannelID=" & ChannelID & " And ClassType=1 order by RootID,OrderID"
    Set rsClass = Conn.Execute(sqlClass)
    If rsClass.BOF And rsClass.EOF Then
        strClass_Option = strClass_Option & "<option value=''>���������Ŀ</option>"
    Else
        Do While Not rsClass.EOF
            ClassNum = ClassNum + 1
            tmpDepth = rsClass("Depth")
            If rsClass("NextID") > 0 Then
                arrShowLine(tmpDepth) = True
            Else
                arrShowLine(tmpDepth) = False
            End If
            If ShowType = 1 Then
                strClass_Option = strClass_Option & "<option value='" & rsClass("ClassID") & "'"
            Else
                If rsClass("ParentID") > 0 Then
                    CheckParentPath = ChannelDir & "all," & rsClass("ParentPath") & "," & rsClass("ClassID") & "," & rsClass("arrChildID")
                Else
                    CheckParentPath = ChannelDir & "all," & rsClass("ClassID") & "," & rsClass("arrChildID")
                End If

                If CheckPurview_Class(arrClass_Input, CheckParentPath) = True Then
                    PurviewChecked = True
                    If rsClass("Child") > 0 And rsClass("EnableAdd") = False And rsClass("ClassID") <> CurrentID Then
                        strClass_Option = strClass_Option & "<option value='0'"
                    Else
                        strClass_Option = strClass_Option & "<option value='" & rsClass("ClassID") & "'"
                    End If
                Else
                    PurviewChecked = False
                End If
            End If
            If ShowType = 1 Or PurviewChecked = True Then
                If CurrentID = 0 Then
                    If ClassNum = 1 Then
                        strClass_Option = strClass_Option & " selected"
                    End If
                Else
                    If rsClass("ClassID") = CurrentID Then
                        strClass_Option = strClass_Option & " selected"
                    End If
                End If
                strClass_Option = strClass_Option & ">"
                
                If tmpDepth > 0 Then
                    For i = 1 To tmpDepth
                        strClass_Option = strClass_Option & "&nbsp;&nbsp;"
                    Next
                End If
                strClass_Option = strClass_Option & rsClass("ClassName")
                strClass_Option = strClass_Option & "</option>"
            
                ClassNum = ClassNum + 1
            End If
            rsClass.MoveNext
        Loop
    End If
    rsClass.Close
    Set rsClass = Nothing
    User_GetClass_Option = strClass_Option	
End Function


'**************************************************
'��������User_GetChannel_Option
'��  �ã���ʾ�û�Ƶ�������˵�
'��  ����ShowType ----��ʾ����
'        CurrentID ----��ǰ��ĿID
'����ֵ���û���Ŀ�����˵�
'**************************************************

Function User_GetChannel_Option()
    arrClass_Input = Conn.Execute("SELECT arrClass_Input from PE_UserGroup where GroupID=-1")(0)
    Dim strChannel_Option,rsChannel
    Set rsChannel = Conn.Execute("SELECT ChannelID,ChannelName,ChannelShortName,ChannelDir FROM PE_Channel WHERE ChannelType<=1 AND ModuleType=1 AND Disabled=" & PE_False & " ORDER BY OrderID")	
    If PE_CLng(ChannelID) = 0 Then
        strChannel_Option = strChannel_Option & "<option value='0' selected>��ѡ��Ƶ��</option>"
    End If
    Do While not rsChannel.Eof 
	    If FoundInArr(arrClass_Input, rsChannel("ChannelDir") & "none", ",") = True Then
        Else
            strChannel_Option = strChannel_Option & "<option value='" & rsChannel("ChannelID") &"'"
            If rsChannel("ChannelID") = ChannelID then strChannel_Option = strChannel_Option & " selected"			   
            strChannel_Option = strChannel_Option &  ">" &rsChannel(1)&"</option>"		
		End If	
	rsChannel.MoveNext
	Loop
    User_GetChannel_Option = strChannel_Option
End Function 
%>
<!--#include file="../Include/PowerEasy.Common.Manage.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<%
Dim ArticleID, AuthorName, Status, ManageType
Dim IncludePic, UploadFiles, DefaultPicUrl
Dim ArticlePro1, ArticlePro2, ArticlePro3, ArticlePro4
Dim ClassID, SpecialID, OnTop, IsElite, IsHot, Created
Dim tClass, ClassName, RootID, ParentID, Depth, ParentPath, Child, arrChildID, ParentDir, ClassDir, ClassPurview
Dim ChannelUpload
Sub Execute()
    ChannelID = PE_CLng(Request("ChannelID"))
    Call GetUserAmon
    If ChannelID > 0 Then
        Call GetChannel(ChannelID)
        ChannelUpload = Conn.Execute("Select EnableUploadFile From PE_Channel Where ChannelID = "& ChannelID)(0)
        If ChannelUpload = False Then IsUpload = False		
    'Else
    '   FoundErr = True
    '   ErrMsg = ErrMsg & "<li>��ָ��Ҫ�鿴��Ƶ��ID��</li>"
    '   Response.Write ErrMsg
    '   Exit Sub
    Else
        ChannelShortName = "����"	
        IsUpload = False			
    End If

    ArticleID = Trim(Request("ArticleID"))
    ClassID = PE_CLng(Trim(Request("ClassID")))
    Status = Trim(Request("Status"))
    AuthorName = Trim(Request("AuthorName"))
    If Status = "" Then
        Status = 9
    Else
        Status = PE_CLng(Status)
    End If
    If IsValidID(ArticleID) = False Then
        ArticleID = ""
    End If
    ManageType = Trim(Request("ManageType"))

    If Action = "" Then Action = "Manage"
    FileName = "User_Anonymous.asp?ChannelID=" & ChannelID & "&Action=" & Action & "&ManageType=" & ManageType
    strFileName = FileName & "&ClassID=" & ClassID & "&Field=" & strField & "&keyword=" & Keyword
    If AuthorName <> "" Then
        AuthorName = ReplaceBadChar(AuthorName)
        strFileName = strFileName & "&AuthorName=" & AuthorName
    End If


    If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))
    ArticlePro1 = XmlText("Article", "ArticlePro1", "[ͼ��]")
    ArticlePro2 = XmlText("Article", "ArticlePro2", "[��ͼ]")
    ArticlePro3 = XmlText("Article", "ArticlePro3", "[�Ƽ�]")
    ArticlePro4 = XmlText("Article", "ArticlePro4", "[ע��]")


    Select Case Action
    Case "Add"
        Call Add
    Case "SaveAdd"
        Call SaveArticle
    Case "Preview"
        Call Preview
    Case "Del"
        Call Del
    Case "Show"
        Call Show	
    Case Else
        Call Add
    End Select
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
End Sub

Function GetStrNoItem(iClassID, iStatus)
    Dim strNoItem
    strNoItem = ""
    If ClassID > 0 Then
        strNoItem = strNoItem & "����Ŀ��������Ŀ��û���κ�"
    Else
        strNoItem = strNoItem & "û���κ�"
    End If
    Select Case Status
    Case -2
        strNoItem = strNoItem & "δ�����õ�" & ChannelShortName
    Case -1
        strNoItem = strNoItem & "�ݸ�"
    Case 0
        strNoItem = strNoItem & "<font color=blue>�����</font>��" & ChannelShortName & "��"
    Case 3
        strNoItem = strNoItem & "<font color=green>�����</font>��" & ChannelShortName & "��"
    Case Else
        strNoItem = strNoItem & "" & ChannelShortName & "��"
    End Select
    GetStrNoItem = strNoItem
End Function

Function GetInfoIncludePic(IncludePic)
    Dim strInfoIncludePic
    Select Case PE_CLng(IncludePic)
        Case 1
            strInfoIncludePic = "<font color=blue>" & ArticlePro1 & "</font>"
        Case 2
            strInfoIncludePic = "<font color=blue>" & ArticlePro2 & "</font>"
        Case 3
            strInfoIncludePic = "<font color=blue>" & ArticlePro3 & "</font>"
        Case 4
            strInfoIncludePic = "<font color=blue>" & ArticlePro4 & "</font>"
    End Select
    GetInfoIncludePic = strInfoIncludePic
End Function

Function GetLinkTips(Title, Author, CopyFrom, UpdateTime, Hits, Keyword, Stars, PaginationType, InfoPoint)
    Dim strLinkTips
    strLinkTips = ""
    strLinkTips = strLinkTips & "��&nbsp;&nbsp;&nbsp;&nbsp;�⣺" & Title & vbCrLf
    strLinkTips = strLinkTips & "��&nbsp;&nbsp;&nbsp;&nbsp;�ߣ�" & Author & vbCrLf
    strLinkTips = strLinkTips & "ת �� �ԣ�" & CopyFrom & vbCrLf
    strLinkTips = strLinkTips & "����ʱ�䣺" & UpdateTime & vbCrLf
    strLinkTips = strLinkTips & "�� �� ����" & Hits & vbCrLf
    strLinkTips = strLinkTips & "�� �� �֣�" & Mid(Keyword, 2, Len(Keyword) - 2) & vbCrLf
    strLinkTips = strLinkTips & "�Ƽ��ȼ���"
    If Stars = 0 Then
        strLinkTips = strLinkTips & "��"
    Else
        strLinkTips = strLinkTips & String(Stars, "��")
    End If
    strLinkTips = strLinkTips & vbCrLf
    strLinkTips = strLinkTips & "��ҳ��ʽ��"
    Select Case PaginationType
    Case 0
        strLinkTips = strLinkTips & "����ҳ"
    Case 1
        strLinkTips = strLinkTips & "�Զ���ҳ"
    Case 2
        strLinkTips = strLinkTips & "�ֶ���ҳ"
    End Select
    strLinkTips = strLinkTips & vbCrLf
    strLinkTips = strLinkTips & "�Ķ�������" & InfoPoint
    GetLinkTips = strLinkTips
End Function

Function GetInfoStatus(iStatus)
    Dim strInfoStatus
    Select Case iStatus
    Case -2
        strInfoStatus = "<font color=gray>�˸�</font>"
    Case -1
        strInfoStatus = "<font color=gray>�ݸ�</font>"
    Case 0
        strInfoStatus = "<font color=red>�����</font>"
    Case 1
        strInfoStatus = "<font color=red>һ��ͨ��</font>"
    Case 2
        strInfoStatus = "<font color=red>����ͨ��</font>"
    Case 3
        strInfoStatus = "<font color=black>����ͨ��</font>"
    End Select
    GetInfoStatus = strInfoStatus
End Function

Function GetInfoProperty(OnTop, Hits, Elite, DefaultPicUrl)
    Dim strInfoProperty
    strInfoProperty = ""
    If OnTop = True Then
        strInfoProperty = strInfoProperty & "<font color=blue>��</font>"
    Else
        strInfoProperty = strInfoProperty & "&nbsp;&nbsp;"
    End If
    strInfoProperty = strInfoProperty & "&nbsp;"
    If Hits >= HitsOfHot Then
        strInfoProperty = strInfoProperty & "<font color=red>��</font>"
    Else
        strInfoProperty = strInfoProperty & "&nbsp;&nbsp;"
    End If
    strInfoProperty = strInfoProperty & "&nbsp;"
    If Elite = True Then
        strInfoProperty = strInfoProperty & "<font color=green>��</font>"
    Else
        strInfoProperty = strInfoProperty & "&nbsp;&nbsp;"
    End If
    strInfoProperty = strInfoProperty & "&nbsp;"
    If DefaultPicUrl <> "" Then
        strInfoProperty = strInfoProperty & "<font color=blue>ͼ</font>"
    Else
        strInfoProperty = strInfoProperty & "&nbsp;&nbsp;"
    End If
    GetInfoProperty = strInfoProperty
End Function

Sub ShowJS_Article()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function AddItem(strFileName){" & vbCrLf
    Response.Write "  var arrName=strFileName.split('.');" & vbCrLf
    Response.Write "  var FileExt=arrName[1];" & vbCrLf
    Response.Write "  if (FileExt=='gif'||FileExt=='jpg'||FileExt=='jpeg'||FileExt=='jpe'||FileExt=='bmp'||FileExt=='png'){" & vbCrLf
    
    If PE_CLng(UserSetting(21)) = 1 Then
        Response.Write "      if(document.myform.IncludePic.selectedIndex<2){" & vbCrLf
        Response.Write "        document.myform.IncludePic.selectedIndex+=1;" & vbCrLf
        Response.Write "      }" & vbCrLf
    End If

    Response.Write "  document.myform.DefaultPicUrl.value=strFileName;}" & vbCrLf
    Response.Write "  document.myform.DefaultPicList.options[document.myform.DefaultPicList.length]=new Option(strFileName,strFileName);" & vbCrLf
    Response.Write "  document.myform.DefaultPicList.selectedIndex+=1;" & vbCrLf
    Response.Write "  if(document.myform.UploadFiles.value==''){" & vbCrLf
    Response.Write "    document.myform.UploadFiles.value=strFileName;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    document.myform.UploadFiles.value=document.myform.UploadFiles.value+'|'+strFileName;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function selectPaginationType(){" & vbCrLf
    Response.Write "  document.myform.PaginationType.value=2;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function rUseLinkUrl(){" & vbCrLf
    Response.Write "  if(document.myform.UseLinkUrl.checked==true){" & vbCrLf
    Response.Write "    document.myform.LinkUrl.disabled=false;" & vbCrLf
    Response.Write "     ArticleContent.style.display='none';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    document.myform.LinkUrl.disabled=true;" & vbCrLf
    Response.Write "    ArticleContent.style.display='';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  var CurrentMode=editor.CurrentMode;" & vbCrLf
    Response.Write "  if (CurrentMode==0){" & vbCrLf
    Response.Write "    document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else if(CurrentMode==1){" & vbCrLf
    Response.Write "   document.myform.Content.value=editor.HtmlEdit.document.body.innerText;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "    alert('Ԥ��״̬���ܱ��棡���Ȼص��༭״̬���ٱ���');" & vbCrLf
    Response.Write "    editor.HtmlEdit.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.curChannelID.value=='0'){" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "��ָ��Ƶ����');" & vbCrLf
    Response.Write "    document.myform.curChannelID.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf	
    Response.Write "  if (document.myform.ClassID.value==''){" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "������Ŀ����ָ��Ϊ�ⲿ��Ŀ��');" & vbCrLf
    Response.Write "    document.myform.ClassID.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.ClassID.value=='0'){" & vbCrLf
    Response.Write "    alert('ָ������Ŀ���������" & ChannelShortName & "��ֻ������������Ŀ�����" & ChannelShortName & "��');" & vbCrLf
    Response.Write "    document.myform.ClassID.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.ClassID.value=='-1'){" & vbCrLf
    Response.Write "    alert('��û���ڴ���Ŀ����" & ChannelShortName & "��Ȩ�ޣ���ѡ��������Ŀ��');" & vbCrLf
    Response.Write "    document.myform.ClassID.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Title.value==''){" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "���ⲻ��Ϊ�գ�');" & vbCrLf
    Response.Write "    document.myform.Title.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Keyword.value==''){" & vbCrLf
    Response.Write "    alert('�ؼ��ֲ���Ϊ�գ�');" & vbCrLf
    Response.Write "    document.myform.Keyword.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    If PE_CLng(UserSetting(23)) = 1 Then
        Response.Write "  if(document.myform.UseLinkUrl.checked==true){" & vbCrLf
        Response.Write "    if (document.myform.LinkUrl.value==''||document.myform.LinkUrl.value=='http://'){" & vbCrLf
        Response.Write "      alert('������ת�����ӵĵ�ַ��');" & vbCrLf
        Response.Write "      document.myform.LinkUrl.focus();" & vbCrLf
        Response.Write "      return false;" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "  }" & vbCrLf
        Response.Write "  else{" & vbCrLf
        Response.Write "    if (document.myform.Content.value==''){" & vbCrLf
        Response.Write "      alert('" & ChannelShortName & "���ݲ���Ϊ�գ�');" & vbCrLf
        Response.Write "      editor.HtmlEdit.focus();" & vbCrLf
        Response.Write "      return false;" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "  }" & vbCrLf
    Else
        Response.Write "  if (document.myform.Content.value==''){" & vbCrLf
        Response.Write "    alert('" & ChannelShortName & "���ݲ���Ϊ�գ�');" & vbCrLf
        Response.Write "    editor.HtmlEdit.focus();" & vbCrLf
        Response.Write "    return false;" & vbCrLf
        Response.Write "  }" & vbCrLf
    End If
    Response.Write "  return true;  " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub Add()
    If FoundInArr(arrClass_Input, ChannelDir & "none", ",") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�Բ�����û����" & ChannelName & "���" & ChannelShortName & "��Ȩ�ޣ�</li><br><br>"
        Exit Sub
    End If
    Dim trs
    If MaxPerDay > 0 Then
        Set trs = Conn.Execute("select count(ArticleID) from PE_Article where Inputer='" & UserName & "' and Deleted=" & PE_False & " and DateDiff(" & PE_DatePart_D & "," & PE_Now & ",UpdateTime)=0")
        If trs(0) >= MaxPerDay Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�����췢���" & ChannelShortName & "�Ѿ��ﵽ�����ޣ�</li>"
        End If
        Set trs = Nothing
        If FoundErr = True Then Exit Sub
    End If
    
    Call ShowJS_Article
    
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='User_Anonymous.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center' colspan='2'><b>���" & ChannelShortName & "</td>"
    Response.Write "    </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>����Ƶ����</strong></td>"
    Response.Write "          <td><select  onchange=""window.location.href='User_Anonymous.asp?ChannelID='+this.options[this.selectedIndex].value"" name='curChannelID'>" & User_GetChannel_Option() & "</select></td>"
    Response.Write "        </tr>"	
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>������Ŀ��</strong></td>"
    Response.Write "          <td><select name='ClassID'>" & User_GetClass_Option(4, ClassID) & "</select></td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>����ר�⣺</strong></td>"
    Response.Write "          <td><select name='SpecialID'><option value='0'>�������κ�ר��</option>" & GetSpecial_Option(0) & "</select></td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "���⣺</strong></td>"
    Response.Write "          <td>"
    If PE_CLng(UserSetting(21)) = 1 Then
        Response.Write "<select name='IncludePic'><option  value='0' selected> </option><option value='1'>" & ArticlePro1 & "</option><option value='2'>" & ArticlePro2 & "</option><option value='3'>" & ArticlePro3 & "</option><option value='4'>" & ArticlePro4 & "</option></select>"
    Else
       Response.Write "<Input TYPE='hidden' Name='IncludePic' value=''>"
    End If
    Response.Write "          <input name='Title' type='text' id='Title' value='' size='45' maxlength='255' class='bginput'> <font color='#FF0000'>*</font>"
    If PE_CLng(UserSetting(22)) = 1 Then
        Response.Write "<input name='ShowCommentLink' type='checkbox' id='ShowCommentLink' value='Yes'> ��ʾ" & ChannelShortName & "�б�ʱ�ڱ�������ʾ��������"
    End If
    Response.Write "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>�ؼ��֣�</strong></td>"
    Response.Write "          <td><input name='Keyword' type='text' id='Keyword' value='" & Trim(Session("Keyword")) & "' size='50' maxlength='255'> <font color='#FF0000'>*</font>" & GetKeywordList("User", ChannelID)
    Response.Write "<br><font color='#0000FF'>�����������" & ChannelShortName & "�����������ؼ��֣��м���<font color='#FF0000'>��|��</font>���������ܳ���&quot;'&?;:()���ַ���</font></td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "���ߣ�</strong></td>"
    Response.Write "          <td>"
    Response.Write "              <input name='Author' type='text' id='Author' value='" & Trim(Session("Author")) & "' size='50' maxlength='100'>" & GetAuthorList("User", ChannelID, UserName)
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "��Դ��</strong></td>"
    Response.Write "          <td>"
    Response.Write "              <input name='CopyFrom' type='text' id='CopyFrom' value='" & Trim(Session("CopyFrom")) & "' size='50' maxlength='100'>" & GetCopyFromList("User", ChannelID)
    Response.Write "          </td>"
    Response.Write "        </tr>"
    If PE_CLng(UserSetting(23)) = 1 Then
        Response.Write "        <tr class='tdbg'>"
        Response.Write "          <td width='120' align='right' class='tdbg5'><strong><font color='#FF0000'>ת�����ӣ�</font></strong></td>"
        Response.Write "          <td>"
        Response.Write "            <input name='LinkUrl' type='text' id='LinkUrl' value='http://' size='50' maxlength='255' disabled>"
        Response.Write "            <input name='UseLinkUrl' type='checkbox' id='UseLinkUrl' value='Yes' onClick='rUseLinkUrl();'>"
        Response.Write "            <font color='#FF0000'>ʹ��ת������</font></td>"
        Response.Write "        </tr>"
    End If
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "��飺</strong></td>"
    Response.Write "            <td ><textarea name='Intro' cols='80' rows='4'></textarea></td>"
    Response.Write "          </tr>"
    Response.Write "        <tr class='tdbg' id='ArticleContent' style=""display:''"">"
    Response.Write "          <td width='120' align='right' class='tdbg5'><p><strong>" & ChannelShortName & "���ݣ�</strong></p>"
    Response.Write "<br><br><font color='red'>�����밴Shift+Enter<br><br>����һ���밴Enter</font></div>"
    Response.Write "         </td>"
    Response.Write "         <td><textarea name='Content' style='display:none'>" & XmlText("Article", "DefaultAddTemplate", "") & "</textarea>"
    
    If PE_CLng(UserSetting(24)) = 1 Then
        Response.Write "            <iframe id='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=0&tContentid=Content&Anonymous=1' frameborder=1 scrolling=no width='600' height='405'></iframe>"
    Else
        Response.Write "            <iframe id='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=2&tContentid=Content&Anonymous=1' frameborder=1 scrolling=no width='600' height='405'></iframe>"
    End If
    Response.Write "         </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong><font color='#FF0000'>��ҳͼƬ��</font></strong></td>"
    Response.Write "          <td><input name='DefaultPicUrl' type='text' id='DefaultPicUrl' size='56' maxlength='200'>"
    Response.Write "      ��������ҳ��ͼƬ" & ChannelShortName & "����ʾ <br>ֱ�Ӵ��ϴ�ͼƬ��ѡ��"
    Response.Write "      <select name='DefaultPicList' id='select' onChange='DefaultPicUrl.value=this.value;'>"
    Response.Write "        <option selected>��ָ����ҳͼƬ</option>"
    Response.Write "      </select><input name='UploadFiles' type='hidden' id='UploadFiles'>"
    Response.Write "          </td>"
    Response.Write "          </tr>"
    '�Զ����ֶ�
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & ChannelID & " or ChannelID=-1")
    Do While Not rsField.EOF
        IF rsField("ShowOnForm") = True then
            Call WriteFieldHTML(rsField("FieldName"), rsField("Title"), rsField("Tips"), rsField("FieldType"), rsField("DefaultValue"), rsField("Options"), rsField("EnableNull"))
        End If
        rsField.MoveNext
    Loop
    Set rsField = Nothing
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "״̬��</strong></td>"
    Response.Write "            <td><input name='Status' type='radio' id='Status' value='-1'>�ݸ�&nbsp;&nbsp;<input Name='Status' Type='Radio' Id='Status' Value='0' checked>Ͷ��</td>"
    Response.Write "          </tr>"
    Response.Write "  </table>"
    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='SaveAdd'><input name='PaginationType' type='hidden' value='0'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='add' type='submit'  id='Add' value=' �� �� ' onClick=""document.myform.Action.value='SaveAdd';document.myform.target='_self';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Preview' type='submit'  id='Preview' value=' Ԥ �� ' onClick=""document.myform.Action.value='Preview';document.myform.target='_blank';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='User_Anonymous.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
End Sub

Sub SaveArticle()
    If FoundInArr(arrClass_Input, ChannelDir & "none", ",") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�Բ�����û����" & ChannelName & "���" & ChannelShortName & "��Ȩ�ޣ�</li><br><br>"
        Exit Sub
    End If
    Dim rsArticle, sql, i
    Dim trs
    Dim ArticleID, ClassID, SpecialID, Title, ShowCommentLink, Keyword, UseLinkUrl, LinkUrl, Content, tAuthor, Intro
    Dim Author, CopyFrom, Inputer
    Dim arrUploadFiles, SaveRemotePic
    Dim PresentExp, DefaultItemPoint, DefaultItemChargeType, DefaultItemPitchTime, DefaultItemReadTimes, DefaultItemDividePercent

    ArticleID = PE_CLng(Trim(Request.Form("ArticleID")))
    ClassID = PE_CLng(Trim(Request.Form("ClassID")))
    SpecialID = PE_CLng(Trim(Request.Form("SpecialID")))
    Title = PE_HTMLEncode(Trim(Request.Form("Title")))
    ShowCommentLink = Trim(Request.Form("ShowCommentLink"))
    Keyword = Trim(Request.Form("Keyword"))
    UseLinkUrl = PE_HTMLEncode(Trim(Request.Form("UseLinkUrl")))
    LinkUrl = PE_HTMLEncode(Trim(Request.Form("LinkUrl")))
    Intro = PE_HTMLEncode(Trim(Request.Form("Intro")))
    For i = 1 To Request.Form("Content").Count
        Content = Content & FilterJS(Request.Form("Content")(i))
    Next
    Author = PE_HTMLEncode(Trim(Request.Form("Author")))
    CopyFrom = PE_HTMLEncode(Trim(Request.Form("CopyFrom")))
    IncludePic = PE_CLng(Trim(Request.Form("IncludePic")))
    DefaultPicUrl = PE_HTMLEncode(Trim(Request.Form("DefaultPicUrl")))
    UploadFiles = PE_HTMLEncode(Trim(Request.Form("UploadFiles")))
    SaveRemotePic = PE_HTMLEncode(Trim(Request.Form("SaveRemotePic")))
    Inputer = UserName
    Status = PE_CLng(Trim(Request.Form("Status")))
    If ClassID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>δָ��������Ŀ������ָ������Ŀ������˲�����</li>"
    Else
        Dim tClass
        Set tClass = Conn.Execute("select ClassName,ClassType,Depth,ParentID,ParentPath,Child,EnableAdd,PresentExp,DefaultItemPoint,DefaultItemChargeType,DefaultItemPitchTime,DefaultItemReadTimes,DefaultItemDividePercent from PE_Class where ClassID=" & ClassID)
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
            DefaultItemPoint = tClass("DefaultItemPoint")
            DefaultItemChargeType = tClass("DefaultItemChargeType")
            DefaultItemPitchTime = tClass("DefaultItemPitchTime")
            DefaultItemReadTimes = tClass("DefaultItemReadTimes")
            DefaultItemDividePercent = tClass("DefaultItemDividePercent")

            If Child > 0 And tClass("EnableAdd") = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>ָ������Ŀ���������" & ChannelShortName & "</li>"
            End If
            If tClass("ClassType") = 2 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>����ָ��Ϊ�ⲿ��Ŀ</li>"
            End If
            Dim CheckParentPath
            If ParentID > 0 Then
                CheckParentPath = ChannelDir & "all," & ParentPath & "," & ClassID
            Else
                CheckParentPath = ChannelDir & "all," & ClassID
            End If
            If CheckPurview_Class(arrClass_Input, CheckParentPath) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�Բ�����û�д���Ŀ����Ӧ����Ȩ�ޣ�</li>"
            End If
        End If
        Set tClass = Nothing
    End If

    If Title = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "���ⲻ��Ϊ��</li>"
    Else
        Title = ReplaceText(Title, 2)
    End If

    If Author = "" Then Author = XmlText("BaseText", "DefAuthor", "����")
    If CopyFrom = "" Then CopyFrom = XmlText("BaseText", "DefCopyFrom", "��վԭ��")
    Keyword = ReplaceBadChar(Keyword)
    If Keyword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������" & ChannelShortName & "�ؼ���</li>"
    Else
        Call SaveKeyword(Keyword)
    End If
    If UseLinkUrl = "Yes" Then
        If LinkUrl = "" Or LCase(LinkUrl) = "http://" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���ӵ�ַ����Ϊ��</li>"
        ElseIf Left(LCase(LinkUrl), 7) <> "http://" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���ӵ�ַ������ http:// ��ͷ</li>"
        End If
    Else
        If Content = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & ChannelShortName & "���ݲ���Ϊ��</li>"
        End If
    End If

    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & ChannelID & " or ChannelID=-1")
    Do While Not rsField.EOF
        If rsField("EnableNull") = False Then
            If Trim(Request(rsField("FieldName"))) = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>������" & rsField("Title") & "��</li>"
            End If
        End If
        rsField.MoveNext
    Loop
    
    If FoundErr = True Then
        Exit Sub
    End If

    If Status < 0 Then
        Status = -1
    Else
        If CheckLevel = 0 Or NeedlessCheck = 1 Then
            Status = 3
        Else
            Status = 0
        End If
    End If

    Keyword = "|" & ReplaceBadChar(Keyword) & "|"

    '�����Ե�ַת��Ϊ��Ե�ַ
    Dim strSiteUrl
    strSiteUrl = Request.ServerVariables("HTTP_REFERER")
    strSiteUrl = LCase(Left(strSiteUrl, InStrRev(strSiteUrl, "/") - 1))
    strSiteUrl = Left(strSiteUrl, InStrRev(strSiteUrl, "/")) & ChannelDir & "/"
    Content = ReplaceBadUrl(Replace(Content, strSiteUrl, "[InstallDir_ChannelDir]"))
    strSiteUrl = InstallDir & ChannelDir & "/"
    Content = Replace(Content, strSiteUrl, "[InstallDir_ChannelDir]")

    Set rsArticle = Server.CreateObject("adodb.recordset")
    If Action = "SaveAdd" Then
        If Session("Title") = Title And DateDiff("S", Session("AddTime"), Now()) < 100 Then
            FoundErr = True
            ErrMsg = "<li>�벻Ҫ�ظ����ͬһƪ����</li>"
            Exit Sub
        Else
            Session("Title") = Title
            Session("AddTime") = Now()
            If MaxPerDay > 0 Then
                Set trs = Conn.Execute("select count(ArticleID) from PE_Article where Inputer='" & UserName & "' and DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ")=0")
                If trs(0) >= MaxPerDay Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�����췢���" & ChannelShortName & "�Ѿ��ﵽ�����ޣ�</li>"
                End If
                Set trs = Nothing
                If FoundErr = True Then Exit Sub
            End If
            
            sql = "select top 1 * from PE_Article"
            rsArticle.Open sql, Conn, 1, 3
            rsArticle.addnew
            ArticleID = PE_CLng(Conn.Execute("select max(ArticleID) from PE_Article")(0)) + 1
            Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (1," & ArticleID & "," & SpecialID & ")")
            rsArticle("ArticleID") = ArticleID
            rsArticle("ChannelID") = ChannelID
            rsArticle("ClassID") = ClassID
            rsArticle("Title") = Title
            rsArticle("Intro") = Intro
            rsArticle("Content") = Content
            rsArticle("Keyword") = Keyword
            rsArticle("Hits") = 0
            rsArticle("Author") = Author
            rsArticle("CopyFrom") = CopyFrom
            rsArticle("LinkUrl") = LinkUrl
            rsArticle("Inputer") = "����Ͷ��"
            rsArticle("Editor") = "����Ͷ��"
            rsArticle("IncludePic") = IncludePic
            If ShowCommentLink = "Yes" Then
                rsArticle("ShowCommentLink") = True
            Else
                rsArticle("ShowCommentLink") = False
            End If
            rsArticle("Status") = Status
            rsArticle("OnTop") = False
            'rsArticle("Hot") = False
            rsArticle("Elite") = False
            rsArticle("Stars") = 0
            rsArticle("UpdateTime") = Now()
            rsArticle("PaginationType") = 0
            rsArticle("MaxCharPerPage") = 0
            rsArticle("SkinID") = 0
            rsArticle("TemplateID") = 0
            rsArticle("DefaultPicUrl") = DefaultPicUrl
            rsArticle("UploadFiles") = UploadFiles
            rsArticle("Deleted") = False
            PresentExp = CLng(PresentExp * PresentExpTimes)
            rsArticle("PresentExp") = PresentExp
            rsArticle("InfoPoint") = DefaultItemPoint
            rsArticle("VoteID") = 0
            rsArticle("InfoPurview") = 0
            rsArticle("arrGroupID") = ""
            rsArticle("ChargeType") = DefaultItemChargeType
            rsArticle("PitchTime") = DefaultItemPitchTime
            rsArticle("ReadTimes") = DefaultItemReadTimes
            rsArticle("DividePercent") = DefaultItemDividePercent
            rsArticle("Copymoney") = 0
            rsArticle("IsPayed") = False
            rsArticle("Beneficiary") = UserName
            
            If Not (rsField.BOF And rsField.EOF) Then
                rsField.MoveFirst
                Do While Not rsField.EOF
                    If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                        rsArticle(Trim(rsField("FieldName"))) = PE_HTMLEncode(FilterJS(Trim(Request(rsField("FieldName")))))
                    End If
                    rsField.MoveNext
                Loop
            End If
            Set rsField = Nothing         
            rsArticle.Update
        End If
    End If
    rsArticle.Close
    Set rsArticle = Nothing
    
    If FoundErr = True Then Exit Sub

    Call UpdateChannelData(ChannelID)
  '  Call UpdateUserData(0, UserName, 0, 0)
    
    Response.Write "<br><br>"
    Response.Write "<table class='border' align=center width='400' border='0' cellpadding='0' cellspacing='0' bordercolor='#999999'>"
    Response.Write "  <tr align=center> "
    Response.Write "    <td  height='22' align='center' class='title'> "
    If Action = "SaveAdd" Then
        Response.Write "<b>���" & ChannelShortName & "�ɹ�</b>"
    Else
        Response.Write "<b>�޸�" & ChannelShortName & "�ɹ�</b>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    If Status = 0 Then
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td height='60'><font color='#0000FF'>ע�⣺</font><br>&nbsp;&nbsp;&nbsp;&nbsp;����" & ChannelShortName & "��δ��������ֻ�еȹ���Ա��˲�ͨ��������" & ChannelShortName & "��������ӵ�" & ChannelShortName & "�Żᷢ��</td>"
        Response.Write "  </tr>"
    End If
    Response.Write "  <tr>"
    Response.Write "    <td><table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right'><strong>������Ŀ��</strong></td>"
    Response.Write "          <td>" & ShowClassPath() & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right'><strong>" & ChannelShortName & "���⣺</strong></td>"
    Response.Write "          <td>" & Title & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right'><strong>��&nbsp;&nbsp;&nbsp;&nbsp;�ߣ�</strong></td>"
    Response.Write "          <td>" & Author & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right'><strong>ת �� �ԣ�</strong></td>"
    Response.Write "          <td>" & CopyFrom & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'> "
    Response.Write "          <td width='100' align='right'><strong>�� �� �֣�</strong></td>"
    Response.Write "          <td>" & Mid(Keyword, 2, Len(Keyword) - 2) & "</td>"
    Response.Write "        </tr>"
    Response.Write "      </table></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='30' align='center'>"
    Response.Write "��<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "&Action=Modify&ArticleID=" & ArticleID & "'>�޸ı���</a>��&nbsp;"
    Response.Write "��<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>�������" & ChannelShortName & "</a>��&nbsp;"
    Response.Write "��<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & ArticleID & "'>Ԥ��" & ChannelShortName & "����</a>��"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf

    Session("Keyword") = Trim(Request("Keyword"))
    Session("Author") = Author
    Session("CopyFrom") = CopyFrom
    Call ClearSiteCache(0)
   ' Call CreateAllJS_User
End Sub


Sub Show()
    Dim rsArticle, sql, i
    ArticleID = PE_CLng(ArticleID)
    sql = "select * from PE_Article where Deleted=" & PE_False & " and ArticleID=" & ArticleID & ""
    Set rsArticle = Server.CreateObject("ADODB.Recordset")
    rsArticle.Open sql, Conn, 1, 1
    If rsArticle.BOF And rsArticle.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���" & ChannelShortName & "</li>"
    Else
        If rsArticle("Inputer") <> UserName And FoundInArr(rsArticle("ReceiveUser"), UserName, ",") = False And rsArticle("Inputer")<>"����Ͷ��" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�鿴" & ChannelShortName & "ʧ�ܣ���" & ChannelShortName & "����������ӵġ�</li>"
        End If
        ClassID = rsArticle("ClassID")
        Call GetClass
    End If
    If FoundErr = True Then
        rsArticle.Close
        Set rsArticle = Nothing
        Exit Sub
    End If

    Response.Write "<SCRIPT language='javascript'>" & vbCrLf
    Response.Write "function resizepic(thispic){" & vbCrLf
    Response.Write "  if(thispic.width>600) thispic.width=600;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function bbimg(o){" & vbCrLf
    Response.Write "  var zoom=parseInt(o.style.zoom, 10)||100;" & vbCrLf
    Response.Write "  zoom+=event.wheelDelta/12;" & vbCrLf
    Response.Write "  if (zoom>0) o.style.zoom=zoom+'%';" & vbCrLf
    Response.Write "  return false;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf


    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' width='82%'>"
    Response.Write "�����ڵ�λ�ã�&nbsp;<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "'>" & ChannelShortName & "����</a>&nbsp;&gt;&gt;&nbsp;"
    If ParentID > 0 Then
        Dim sqlPath, rsPath
        sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
        Set rsPath = Server.CreateObject("adodb.recordset")
        rsPath.Open sqlPath, Conn, 1, 1
        Do While Not rsPath.EOF
            Response.Write "<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "&ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    Response.Write "<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "&ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
    Response.Write "<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & rsArticle("ArticleID") & "'>"
    
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

    Response.Write "" & rsArticle("Title") & "</a>"
    Response.Write " </td>"
    Response.Write "    <td width='18%' height='22' align='right'>"

    If rsArticle("OnTop") = True Then
        Response.Write "<font color=blue>��</font>&nbsp;"
    Else
        Response.Write "&nbsp;&nbsp;&nbsp;"
    End If
    If rsArticle("Hits") >= HitsOfHot Then
        Response.Write "<font color=red>��</font>&nbsp;"
    Else
        Response.Write "&nbsp;&nbsp;&nbsp;"
    End If
    If rsArticle("Elite") = True Then
        Response.Write "<font color=green>��</font>"
    Else
        Response.Write "&nbsp;&nbsp;"
    End If
    Response.Write "&nbsp;&nbsp;<font color='#009900'>" & String(rsArticle("Stars"), "��") & "</font>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr align='center' class='tdbg'>"
    Response.Write "    <td colspan='2' height='40' valign='bottom'>"
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
    Response.Write "���ߣ�" & Author & "&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "��Դ��"
    If InStr(CopyFrom, "|") > 0 Then
        Response.Write "<a href='" & Right(CopyFrom, Len(CopyFrom) - InStr(CopyFrom, "|")) & "' target='_blank'>" & Left(CopyFrom, InStr(CopyFrom, "|") - 1) & "</a>"
    Else
        Response.Write "" & CopyFrom
    End If
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;�������" & rsArticle("Hits") & "&nbsp;&nbsp;&nbsp;&nbsp;����ʱ�䣺" & FormatDateTime(rsArticle("UpdateTime"), 2)
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td colspan='2'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='5'>"
    Response.Write "        <tr>"
    Response.Write "          <td height='200' valign='top'>"
    If Trim(rsArticle("LinkUrl")) <> "" Then
        Response.Write "<p align='center'><br><br><br><font color=red>��" & ChannelShortName & "�������ⲿ" & ChannelShortName & "���ݡ����ӵ�ַΪ��<a href='" & rsArticle("LinkUrl") & "' target='_blank'>" & rsArticle("LinkUrl") & "</a></font></p>"
    Else
        Response.Write "<p>" & Replace(Replace(FilterJS(rsArticle("Content")), "[InstallDir_ChannelDir]", InstallDir & ChannelDir & "/"), "{$UploadDir}", UploadDir) & "</p>"
    End If
    Response.Write "       </td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr  align='right' class='tdbg'>"
    Response.Write "    <td colspan='2'>"
    Response.Write "" & ChannelShortName & "¼�룺<a href='User_Anonymous.asp?ChannelID=" & ChannelID & "&Field=Inputer&Keyword=" & rsArticle("Inputer") & "'>" & rsArticle("Inputer") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;���α༭��"
    If rsArticle("Status") > 0 Then
        Response.Write "" & rsArticle("Editor")
    Else
        Response.Write "��"
    End If
    Response.Write " </td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf
    Response.Write "<form name='formA' method='get' action='User_Anonymous.asp'><p align='center'> "
    Response.Write "<input type='hidden' name='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input type='hidden' name='ArticleID' value='" & ArticleID & "'>"
    Response.Write "<input type='hidden' name='Action' value=''>" & vbCrLf
    rsArticle.Close
    Set rsArticle = Nothing
    Response.Write "</Form></p>"
End Sub


Sub Preview()
    Response.Write "<br><table width='760' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td width='400' height='22'>"

    If ClassID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��������Ŀ</li>"
        Exit Sub
    End If

    Call GetClass
    If FoundErr = True Then Exit Sub

    If ParentID > 0 Then
        Dim sqlPath, rsPath
        sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
        Set rsPath = Server.CreateObject("adodb.recordset")
        rsPath.Open sqlPath, Conn, 1, 1
        Do While Not rsPath.EOF
            Response.Write "" & rsPath(1) & "&nbsp;&gt;&gt;&nbsp;"
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    Response.Write "" & ClassName & "&nbsp;&gt;&gt;&nbsp;" & GetInfoIncludePic(Trim(Request("IncludePic"))) & PE_HTMLEncode(Request("Title"))
    Response.Write " </td>"
    Response.Write "    <td width='50' height='22' align='right'>"
    If LCase(Request("OnTop")) = "yes" Then
        Response.Write "��&nbsp;"
    Else
        Response.Write "&nbsp;&nbsp;&nbsp;"
    End If
  
    If LCase(Request("Elite")) = "yes" Then
        Response.Write "��"
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
    Response.Write "  <tr class='tdbg'><td colspan='3' align='center'>���ߣ�" & PE_HTMLEncode(Request("Author")) & "&nbsp;&nbsp;&nbsp;&nbsp;ת���ԣ�" & PE_HTMLEncode(Request("CopyFrom")) & "&nbsp;&nbsp;&nbsp;&nbsp;�������0&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "¼�룺" & UserName & "</td></tr>"
    Response.Write "  <tr class='tdbg'><td colspan='3'><p>" & FilterJS(Request("Content")) & "</p></td></tr>"
    Response.Write "</table>"
    Response.Write "<p align='center'>��<a href='javascript:window.close();'>�رմ���</a>��</p>"
End Sub


Sub WriteFieldHTML(FieldName, Title, Tips, FieldType, strValue, Options, EnableNull)
    Dim strEnableNull
    If EnableNull = False Then
        strEnableNull = " <font color='#FF0000'>*</font>"
    End If
    Response.Write "<tr class='tdbg'><td width='120' align='right' class='tdbg5'>" & Title & "��</b><td colspan='5'>"
    Select Case FieldType
    Case 1,8    '�����ı���
        Response.Write "<input type='text' name='" & FieldName & "' size='80' maxlength='255' value='" & strValue & "'>" & strEnableNull
    Case 2,9    '�����ı���
        Response.Write "<textarea name='" & FieldName & "' cols='80' rows='10'>" & strValue & "</textarea>" & strEnableNull
    Case 3    '�����б�
        Response.Write "<select name='" & FieldName & "'>"
        Dim arrOptions, i
        arrOptions = Split(Options, vbCrLf)
        For i = 0 To UBound(arrOptions)
            Response.Write "<option value='" & arrOptions(i) & "'"
            If arrOptions(i) = strValue Then Response.Write " selected"
            Response.Write ">" & arrOptions(i) & "</option>"
        Next
        Response.Write "</select>" & strEnableNull
    Case 4   'ͼƬ
        If strValue = "" Then
            Response.Write "<input type='text' id='"&FieldName&"' name='"&FieldName&"'  size='45' maxlength='255' value='http://'><br>" & strEnableNull
        Else
           ' Response.Write "<input type='text' name='" & FieldName & "' size='40' maxlength='255' value='" & strValue & "'>" & strEnableNull
            Response.Write "<input type='text' name='" & FieldName & "' id='" & FieldName & "' size='45' maxlength='255' value='" & strValue & "'><br>" & strEnableNull
		End If
        If IsUpload = True Then		
            Response.Write "<iframe style='top:2px;' id='uploadPhoto' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=fieldpic&Anonymous=1&FieldName="& FieldName &"' frameborder=0 scrolling=no width='650' height='25'></iframe>"		
        End If		
    Case 5   '�ļ�
        If strValue = "" Then
            Response.Write "<input type='text' id='"&FieldName&"' name='"&FieldName&"'  size='45' maxlength='255' value='http://'><br>" & strEnableNull
        Else
            Response.Write "<input type='text' name='" & FieldName & "' id='" & FieldName & "' size='45' maxlength='255' value='" & strValue & "'><br>" & strEnableNull
		End If
        If IsUpload = True Then				
            Response.Write "            <iframe style='top:2px' id='uploadsoft' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=fieldsoft&Anonymous=1&FieldName="& FieldName &"' frameborder=0 scrolling=no width='650' height='25'></iframe>"	
        End If						
    Case 6    '����
        If strValue = "" Then
            Response.Write "<input type='text' name='" & FieldName & "' size='20' maxlength='20' value='" & Now() & "'>" & strEnableNull
        Else
            Response.Write "<input type='text' name='" & FieldName & "' size='20' maxlength='20' value='" & strValue & "'>" & strEnableNull
        End If
    Case 7    '����
        If strValue = "" Then
            Response.Write "<input type='text' name='" & FieldName & "'  onkeyup=""value=value.replace(/[^\d]/g,'')"" size='20' maxlength='20' value='0'>" & strEnableNull
        Else
            Response.Write "<input type='text' name='" & FieldName & "'  onkeyup=""value=value.replace(/[^\d]/g,'')""  size='20' maxlength='20' value='" & PE_Clng(strValue) & "'>" & strEnableNull
        End If		
    End Select
    If IsNull(Tips) = False And Tips <> ""  and (FieldType <> 4 and FieldType <> 5) Then
        Response.Write "<br>" & PE_HTMLEncode(Tips)
    End If
    Response.Write "</td></tr>"
End Sub

%>
