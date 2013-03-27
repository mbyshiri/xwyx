<!--#include file="CommonCode.asp"-->
<!--#include file="../Include/PowerEasy.Common.Purview.asp"-->
<!--#include file="../Include/PowerEasy.Common.Manage.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Dim ArticleID, AuthorName, Status, ManageType
Dim IncludePic, UploadFiles, DefaultPicUrl
Dim ArticlePro1, ArticlePro2, ArticlePro3, ArticlePro4
Dim ClassID, SpecialID, OnTop, IsElite, IsHot, Created

Dim tClass, ClassName, RootID, ParentID, Depth, ParentPath, Child, arrChildID, ParentDir, ClassDir, ClassPurview


Sub Execute()
    ChannelID = PE_CLng(Trim(Request("ChannelID")))
    If ChannelID > 0 Then
        Call GetChannel(ChannelID)
    'Else
    '   FoundErr = True
    '   ErrMsg = ErrMsg & "<li>��ָ��Ҫ�鿴��Ƶ��ID��</li>"
    '   Response.Write ErrMsg
    '   Exit Sub
    Else
        ChannelShortName = "����"		
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
    FileName = "User_Article.asp?ChannelID=" & ChannelID & "&Action=" & Action & "&ManageType=" & ManageType
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

    Response.Write "<table align='center'><tr align='center' valign='top'>"
    If CheckUser_ChannelInput() = True Then
        Response.Write "<td width='90'><a href='User_Article.asp?ChannelID=" & ChannelID & "&Action=Add'><img src='images/article_add.gif' border='0' align='absmiddle'><br>���" & ChannelShortName & "</a></td>"
    End If
    Response.Write "<td width='90'><a href='User_Article.asp?ChannelID=" & ChannelID & "&Status=9'><img src='images/article_all.gif' border='0' align='absmiddle'><br>����" & ChannelShortName & "</a></td>"
    Response.Write "<td width='90'><a href='User_Article.asp?ChannelID=" & ChannelID & "&Status=-1'><img src='images/article_draft.gif' border='0' align='absmiddle'><br>�� ��</a></td>"
    Response.Write "<td width='90'><a href='User_Article.asp?ChannelID=" & ChannelID & "&Status=0'><img src='images/article_unpassed.gif' border='0' align='absmiddle'><br>����˵�" & ChannelShortName & "</a></td>"
    Response.Write "<td width='90'><a href='User_Article.asp?ChannelID=" & ChannelID & "&Status=3'><img src='images/article_passed.gif' border='0' align='absmiddle'><br>����˵�" & ChannelShortName & "</a></td>"
    Response.Write "<td width='90'><a href='User_Article.asp?ChannelID=" & ChannelID & "&Status=-2'><img src='images/article_reject.gif' border='0' align='absmiddle'><br>δ�����õ�" & ChannelShortName & "</a></td>"
    Response.Write "</tr></table>" & vbCrLf

    Select Case Action
    Case "Add"
        Call Add
    Case "Modify"
        Call Modify
    Case "SaveAdd", "SaveModify"
        Call SaveArticle
    Case "Preview"
        Call Preview
    Case "Show"
        Call Show
    Case "Del"
        Call Del
    Case "Receive"
        Call Receive
    Case "Manage"
        Call main
    End Select
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
End Sub


Sub main()
    Call GetClass
    If FoundErr = True Then Exit Sub

    Call ShowJS_Main(ChannelShortName)
    If ChannelID > 0 Then
        Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Response.Write "  <tr class='title'>"
        Response.Write "    <td height='22'>" & GetRootClass() & "</td>"
        Response.Write "  </tr>" & GetChild_Root() & ""
        Response.Write "</table><br>"
    End If

    If ManageType = "Receive" Then
        Call ShowContentManagePath(ChannelShortName & "ǩ�չ���")
    Else
        Call ShowContentManagePath(ChannelShortName & "����")
    End If

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "    <form name='myform' method='Post' action='User_Article.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "     <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "          <tr class='title' height='22'> "
    Response.Write "            <td height='22' width='30' align='center'><strong>ѡ��</strong></td>"
    Response.Write "            <td width='25' align='center'><strong>ID</strong></td>"
    Response.Write "            <td align='center' ><strong>" & ChannelShortName & "����</strong></td>"
    Response.Write "            <td width='60' align='center' ><strong>¼��</strong></td>"
    Response.Write "            <td width='40' align='center' ><strong>�����</strong></td>"
    Response.Write "            <td width='80' align='center' ><strong>" & ChannelShortName & "����</strong></td>"
    If ManageType = "Receive" Then
        Response.Write "            <td width='60' align='center' ><strong>ǩ��״̬</strong></td>"
        Response.Write "            <td width='140' align='center' ><strong>ǩ�ղ���</strong></td>"
    Else
        Response.Write "            <td width='60' align='center' ><strong>���״̬</strong></td>"
        Response.Write "            <td width='80' align='center' ><strong>�������</strong></td>"
    End If
    Response.Write "          </tr>"

    Dim rsArticleList, sql, tmpChannelID
    sql = "select A.ArticleID,A.ChannelID,A.ClassID,C.ClassName,C.ParentDir,C.ClassDir,C.ClassPurview,A.Title,A.Keyword,A.Author,A.CopyFrom,A.UpdateTime,A.DefaultPicUrl,A.Inputer,A.ReceiveType,"
    If ManageType = "Receive" Then
        sql = sql & "A.Receive,A.Received,"
    End If
    sql = sql & "A.Hits,A.OnTop,A.Elite,A.Status,A.IncludePic,A.Stars,A.PaginationType,A.InfoPoint from PE_Article A"
    sql = sql & " left join PE_Class C on A.ClassID=C.ClassID where A.Deleted=" & PE_False
    If ChannelID > 0 Then
        sql = sql & " and A.ChannelID=" & ChannelID & " "
    End If
    If ManageType = "Receive" Then
        sql = sql & " and A.ArticleID in (" & GetReceiveArticleID() & ") "
    Else
        sql = sql & " and A.Inputer='" & UserName & "' "
    End If
    Select Case Status
    Case 3
        sql = sql & " and A.Status=3"
    Case 0
        sql = sql & " and (A.Status=0 Or A.Status=1 Or A.Status=2)"
    Case -1
        sql = sql & " and A.Status=-1"
    Case -2
        sql = sql & " and A.Status=-2"
    End Select
    If ClassID > 0 Then
        If Child > 0 Then
            sql = sql & " and A.ClassID in (" & arrChildID & ")"
        Else
            sql = sql & " and A.ClassID=" & ClassID
        End If
    End If

    If Keyword <> "" Then
        Select Case strField
        Case "Title"
            sql = sql & " and A.Title like '%" & Keyword & "%' "
        Case "Content"
            sql = sql & " and A.Content like '%" & Keyword & "%' "
        Case "Author"
            sql = sql & " and A.Author like '%" & Keyword & "%' "
        Case "Inputer"
            sql = sql & " and A.Inputer='" & Keyword & "' "
        Case Else
            sql = sql & " and A.Title like '%" & Keyword & "%' "
        End Select
    End If
    sql = sql & " order by A.ArticleID desc"

    Set rsArticleList = Server.CreateObject("ADODB.Recordset")
    rsArticleList.Open sql, Conn, 1, 1
    If rsArticleList.BOF And rsArticleList.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>" & GetStrNoItem(ClassID, Status) & "<br><br></td></tr>"
    Else
        totalPut = rsArticleList.RecordCount
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
                rsArticleList.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Dim ArticleNum, ArticlePath
        ArticleNum = 0
        Do While Not rsArticleList.EOF
            If ChannelID = 0 Then
                If rsArticleList("ChannelID") <> tmpChannelID Then
                    ChannelID = rsArticleList("ChannelID")
                    Call GetChannel(ChannelID)
                    tmpChannelID = rsArticleList("ChannelID")
                End If
            End If
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "        <td width='30' align='center'><input name='ArticleID' type='checkbox' onclick='unselectall()' id='ArticleID' value='" & rsArticleList("ArticleID") & "'></td>"
            Response.Write "        <td width='25' align='center'>" & rsArticleList("ArticleID") & "</td>"
            Response.Write "        <td>"
            If rsArticleList("ClassID") <> ClassID Then
                Response.Write "<a href='User_Article.asp?ChannelID=" & rsArticleList("ChannelID") & "&Action=Manage&ManageType=" & ManageType & "&ClassID=" & rsArticleList("ClassID") & "'>[" & rsArticleList("ClassName") & "]</a>&nbsp;"
            End If
            Response.Write GetInfoIncludePic(IncludePic)
            Response.Write "<a href='User_Article.asp?ChannelID=" & rsArticleList("ChannelID") & "&Action=Show&ArticleID=" & rsArticleList("ArticleID") & "'"
            Response.Write " title='" & GetLinkTips(rsArticleList("Title"), rsArticleList("Author"), rsArticleList("CopyFrom"), rsArticleList("UpdateTime"), rsArticleList("Hits"), rsArticleList("Keyword"), rsArticleList("Stars"), rsArticleList("PaginationType"), rsArticleList("InfoPoint")) & "'>" & rsArticleList("title") & "</a>"
            Response.Write "</td>"
            Response.Write "      <td width='60' align='center'><a href='" & FileName & "&field=Inputer&keyword=" & rsArticleList("Inputer") & "' title='������鿴���û�¼�������" & ChannelShortName & "'>" & rsArticleList("Inputer") & "</a></td>"
            Response.Write "      <td width='40' align='center'>" & rsArticleList("Hits") & "</td>"
            Response.Write "      <td width='80' align='center'>" & GetInfoProperty(rsArticleList("OnTop"), rsArticleList("Hits"), rsArticleList("Elite"), rsArticleList("DefaultPicUrl")) & "</td>"
            If ManageType = "Receive" Then
                Response.Write "            <td width='60' align='center' >"
                If FoundInArr(rsArticleList("Received"), UserName, "|") = False Then
                    Response.Write "<font color=red>δǩ��</font>"
                Else
                    Response.Write "<font color=green>��ǩ��</font>"
                End If
                Response.Write "</td>"
                Response.Write "            <td width='140' align='center' >"
                Response.Write "<a href='User_Article.asp?ChannelID=" & rsArticleList("ChannelID") & "&Action=Show&ArticleID=" & rsArticleList("ArticleID") & "'"
                Response.Write " title='" & GetLinkTips(rsArticleList("Title"), rsArticleList("Author"), rsArticleList("CopyFrom"), rsArticleList("UpdateTime"), rsArticleList("Hits"), rsArticleList("Keyword"), rsArticleList("Stars"), rsArticleList("PaginationType"), rsArticleList("InfoPoint")) & "' target='_blank' >�鿴�ļ�</a>"
                If FoundInArr(rsArticleList("Received"), UserName, "|") = False Then
                    Response.Write "&nbsp;&nbsp;<a href='User_Article.asp?ChannelID=" & rsArticleList("ChannelID") & "&Action=Receive&ArticleID=" & rsArticleList("ArticleID") & "'>ǩ���ļ�</a>"
                End If
                Response.Write "</td>"
            Else
                Response.Write "    <td width='60' align='center'>" & GetInfoStatus(rsArticleList("Status")) & "</td>"
                Response.Write "    <td width='80' align='center'>"
                If rsArticleList("Inputer") = UserName And (rsArticleList("Status") <= 0 Or EnableModifyDelete = 1) Then
                    Response.Write "<a href='User_Article.asp?ChannelID=" & rsArticleList("ChannelID") & "&Action=Modify&ArticleID=" & rsArticleList("ArticleID") & "'>�޸�</a>&nbsp;"
                    Response.Write "<a href='User_Article.asp?ChannelID=" & rsArticleList("ChannelID") & "&Action=Del&ArticleID=" & rsArticleList("ArticleID") & "' onclick=""return confirm('ȷ��Ҫɾ����" & ChannelShortName & "��һ��ɾ�������ָܻ���');"">ɾ��</a>"
                End If
                Response.Write "</td>"
            End If
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
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>ѡ�б�ҳ��ʾ������" & ChannelShortName & "</td><td>"
    If ManageType = "Receive" Then
        Response.Write "<input name='submit1' type='submit' value='ǩ��ѡ����" & ChannelShortName & "' onClick=""document.myform.Action.value='Receive'"">"
    Else
        Response.Write "<input name='submit1' type='submit' value='ɾ��ѡ����" & ChannelShortName & "' onClick=""document.myform.Action.value='Del'"">"
    End If
    Response.Write "<input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName & "", True)
    End If

    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='80' align='right'><strong>" & ChannelShortName & "������</strong></td>"
    Response.Write "   <td>"
    Response.Write "<table border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "<form method='Get' name='SearchForm' action='" & FileName & "'>"
    Response.Write "<tr><td height='28' align='center'>"
    Response.Write "<select name='Field' size='1'>"
    Response.Write "<option value='Title' selected>" & ChannelShortName & "����</option>"
    Response.Write "<option value='Content'>" & ChannelShortName & "����</option>"
    Response.Write "<option value='Author'>" & ChannelShortName & "����</option>"
    Response.Write "</select>"
    Response.Write "<select name='ClassID'><option value=''>������Ŀ</option>" & User_GetClass_Option(1, 0) & "</select>"
    Response.Write "<input type='text' name='keyword'  size='20' value='�ؼ���' maxlength='50' onFocus='this.select();'>"
    Response.Write "<input type='submit' name='Submit'  value='����'>"
    Response.Write "<input name='ManageType' type='hidden' id='ManageType' value='" & ManageType & "'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "</td></tr></form></table>"
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "<br><b>˵����</b><br>&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "�����еĸ���壺<font color=blue>��</font>----�̶�" & ChannelShortName & "��<font color=red>��</font>----����" & ChannelShortName & "��<font color=green>��</font>----�Ƽ�" & ChannelShortName & "��<font color=blue>ͼ</font>----��ҳͼƬ" & ChannelShortName & "<br><br>"
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
    
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='User_Article.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center' colspan='2'><b>���" & ChannelShortName & "</td>"
    Response.Write "    </tr>"
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
        Response.Write "            <iframe id='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=0&tContentid=Content' frameborder=1 scrolling=no width='600' height='405'></iframe>"
    Else
        Response.Write "            <iframe id='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=2&tContentid=Content' frameborder=1 scrolling=no width='600' height='405'></iframe>"
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
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='User_Article.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
End Sub

Sub Modify()
    If FoundInArr(arrClass_Input, ChannelDir & "none", ",") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�Բ�����û����" & ChannelName & "���" & ChannelShortName & "��Ȩ�ޣ�</li><br><br>"
        Exit Sub
    End If
    Dim rsArticle, sql, tmpAuthor, tmpCopyFrom, SpecialID
    
    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�" & ChannelShortName & "ID</li>"
        Exit Sub
    Else
        ArticleID = PE_CLng(ArticleID)
    End If
    sql = "select * from PE_Article where Inputer='" & UserName & "' and Deleted=" & PE_False & " and ArticleID=" & ArticleID & ""
    Set rsArticle = Server.CreateObject("ADODB.Recordset")
    rsArticle.Open sql, Conn, 1, 1
    If rsArticle.BOF And rsArticle.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���" & ChannelShortName & "</li>"
    Else
        If rsArticle("Status") > 0 And EnableModifyDelete = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & ChannelShortName & "�Ѿ������ͨ����������У��������ٽ����޸ģ�</li>"
        End If
    End If
    If FoundErr = True Then
        rsArticle.Close
        Set rsArticle = Nothing
        Exit Sub
    End If
    SpecialID = PE_CLng(Conn.Execute("select top 1 SpecialID from PE_InfoS where ModuleType=1 and ItemID=" & ArticleID & "")(0))

    tmpAuthor = rsArticle("Author")
    tmpCopyFrom = rsArticle("CopyFrom")

    Call ShowJS_Article

    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='User_Article.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center' colspan='2'><b>�޸�" & ChannelShortName & "</b></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "        <td width='120' align='right' class='tdbg5'><strong>������Ŀ��</strong></td>"
    Response.Write "        <td><select name='ClassID'>" & User_GetClass_Option(4, rsArticle("ClassID")) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "        <td width='120' align='right' class='tdbg5'><strong>����ר�⣺</strong></td>"
    Response.Write "        <td><select name='SpecialID'><option value='0'>�������κ�ר��</option>" & GetSpecial_Option(SpecialID) & "</select></td>"
    Response.Write "    </tr>"

    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "���⣺</strong></td>"
    Response.Write "           <td>"

    If PE_CLng(UserSetting(21)) = 1 Then
        Response.Write "             <select name='IncludePic'>"
        Response.Write "               <option " & OptionValue(rsArticle("IncludePic"), 0) & "> </option>"
        Response.Write "               <option " & OptionValue(rsArticle("IncludePic"), 1) & ">" & ArticlePro1 & "</option>"
        Response.Write "               <option " & OptionValue(rsArticle("IncludePic"), 2) & ">" & ArticlePro2 & "</option>"
        Response.Write "               <option " & OptionValue(rsArticle("IncludePic"), 3) & ">" & ArticlePro3 & "</option>"
        Response.Write "               <option " & OptionValue(rsArticle("IncludePic"), 4) & ">" & ArticlePro4 & "</option>"
        Response.Write "             </select>"
    End If

    Response.Write "          <input name='Title' type='text' id='Title' value='" & rsArticle("Title") & "' size='45' maxlength='255' class='bginput'> <font color='#FF0000'>*</font>"
    If PE_CLng(UserSetting(22)) = 1 Then
        Response.Write "<input name='ShowCommentLink' type='checkbox' id='ShowCommentLink' value='Yes'"
        If rsArticle("ShowCommentLink") = True Then Response.Write " checked"
        Response.Write "> ��ʾ" & ChannelShortName & "�б�ʱ�ڱ�������ʾ�������� </td>"
    End If
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>�ؼ��֣�</strong></td>"
    Response.Write "          <td><input name='Keyword' type='text' id='Keyword' value='" & Mid(rsArticle("Keyword"), 2, Len(rsArticle("Keyword")) - 2) & "' size='50' maxlength='255'> <font color='#FF0000'>*</font>" & GetKeywordList("User", ChannelID)
    Response.Write "<br><font color='#0000FF'>�����������" & ChannelShortName & "�����������ؼ��֣��м���<font color='#FF0000'>��|��</font>���������ܳ���&quot;'&?;:()���ַ���</font></td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "���ߣ�</strong></td>"
    Response.Write "          <td>"
    Response.Write "              <input name='Author' type='text' id='Author' value='" & tmpAuthor & "' size='50' maxlength='100'>" & GetAuthorList("User", ChannelID, UserName)
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "��Դ��</strong></td>"
    Response.Write "          <td>"
    Response.Write "              <input name='CopyFrom' type='text' id='CopyFrom' value='" & tmpCopyFrom & "' size='50' maxlength='100'>" & GetCopyFromList("User", ChannelID)
    Response.Write "          </td>"
    Response.Write "        </tr>"
    If PE_CLng(UserSetting(23)) = 1 Then
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong><font color='#FF0000'>ת�����ӣ�</font></strong></td>"
        Response.Write "            <td><input name='LinkUrl' type='text' id='LinkUrl' value='" & rsArticle("LinkUrl") & "' size='60' maxlength='255'"
        If rsArticle("LinkUrl") = "" Or rsArticle("LinkUrl") = "http://" Then Response.Write " disabled"
        Response.Write "> <input name='UseLinkUrl' type='checkbox' id='UseLinkUrl' value='Yes' onClick='rUseLinkUrl();'"
        If rsArticle("LinkUrl") <> "" And rsArticle("LinkUrl") <> "http://" Then Response.Write " checked"
        Response.Write "><font color='#FF0000'>ʹ��ת������</font></td>"
        Response.Write "          </tr>"
    End If
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "��飺</strong></td>"
    Response.Write "            <td><textarea name='Intro' cols='80' rows='4'>" & PE_ConvertBR(rsArticle("Intro")) & "</textarea></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg' id='ArticleContent' style=""display:'"
    If rsArticle("LinkUrl") <> "" And rsArticle("LinkUrl") <> "http://" Then Response.Write "none"
    Response.Write "'"">"
    Response.Write "            <td width='120' align='right' class='tdbg5'><p><strong>" & ChannelShortName & "���ݣ�</strong></p>"
    Response.Write "<br><br><font color='red'>�����밴Shift+Enter<br><br>����һ���밴Enter</font></div>"
    Response.Write "            </td>"
    Response.Write "            <td><textarea name='Content' style='display:none'>" & Replace(Replace(Server.HTMLEncode(FilterJS(rsArticle("Content"))), "[InstallDir_ChannelDir]", InstallDir & ChannelDir & "/"), "{$UploadDir}", UploadDir) & "</textarea>"
    If PE_CLng(UserSetting(24)) = 1 Then
        Response.Write "            <iframe id='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=0&tContentid=Content' frameborder=1 scrolling=no width='600' height='405'></iframe>"
    Else
        Response.Write "            <iframe id='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=2&tContentid=Content' frameborder=1 scrolling=no width='600' height='405'></iframe>"
    End If
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'> "
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>��ҳͼƬ��</strong></td>"
    Response.Write "            <td><input name='DefaultPicUrl' type='text' id='DefaultPicUrl' value='" & rsArticle("DefaultPicUrl") & "' size='56' maxlength='200'>��������ҳ��ͼƬ" & ChannelShortName & "����ʾ <br>"
    Response.Write "              ֱ�Ӵ��ϴ�ͼƬ��ѡ��<select name='DefaultPicList' id='DefaultPicList' onChange='DefaultPicUrl.value=this.value;'>"
    Response.Write "                <option value=''"
    If rsArticle("DefaultPicUrl") = "" Then Response.Write "selected"
    Response.Write ">��ָ����ҳͼƬ</option>"
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
    Response.Write "              </select><input name='UploadFiles' type='hidden' id='UploadFiles' value='" & rsArticle("UploadFiles") & "'> "
    Response.Write "            </td>"
    Response.Write "          </tr>"
    '�Զ����ֶ�
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & ChannelID & " or ChannelID=-1")
    Do While Not rsField.EOF
        IF rsField("ShowOnForm") = True then
            Call WriteFieldHTML(rsField("FieldName"), rsField("Title"), rsField("Tips"), rsField("FieldType"), rsArticle(Trim(rsField("FieldName"))), rsField("Options"), rsField("EnableNull"))
        End If	
        rsField.MoveNext
    Loop
    Set rsField = Nothing
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "״̬��</td>"
    Response.Write "            <td>"
    If rsArticle("Status") <= 0 Then
        Response.Write "<Input Name='Status' Type='radio' Id='Status' Value='-1'"
        If rsArticle("Status") = -1 Then
            Response.Write " checked"
        End If
        Response.Write "> �ݸ�&nbsp;&nbsp;"
        Response.Write "<Input Name='Status' Type='radio' Id='Status' Value='0'"
        If rsArticle("Status") = 0 Then
            Response.Write "checked"
        End If
        Response.Write "> Ͷ��"
    Else
        If rsArticle("Status") < 3 Then
            Response.Write "�����"
        Else
            Response.Write "�Ѿ�����"
        End If
    End If
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "  </table>"
    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "   <input name='ArticleID' type='hidden' id='ArticleID' value='" & rsArticle("ArticleID") & "'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='Save' type='submit' value='�����޸Ľ��' style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='User_Article.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
    rsArticle.Close
    Set rsArticle = Nothing

End Sub

Sub WriteFieldHTML(FieldName, Title, Tips, FieldType, strValue, Options, EnableNull)
    Dim FieldUpload, ChannelUpload, UserUpload,rsFieldUpload,sqlFieldUpload   
    Select Case FieldType
    Case 4,5
        FieldUpload = True		
        ChannelUpload = Conn.Execute("Select EnableUploadFile from PE_Channel where ChannelID="&ChannelID)(0) 
        If  ChannelUpload = False Then FieldUpload = False
        If UserName<>"" Then   
            sqlFieldUpload = "SELECT U.UserID,U.SpecialPermission,U.UserSetting,G.GroupSetting FROM PE_User U inner join PE_UserGroup G on U.GroupID=G.GroupID WHERE"
            sqlFieldUpload = sqlFieldUpload & " UserName='" & UserName & "'" 
            Set rsFieldUpload = Conn.Execute(sqlFieldUpload)
            If rsFieldUpload.BOF And rsFieldUpload.EOF Then
                FieldUpload = False
            Else
                If rsFieldUpload("SpecialPermission") = True Then
                    UserSetting = Split(Trim(rsFieldUpload("UserSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
                Else
                    UserSetting = Split(Trim(rsFieldUpload("GroupSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
                End If
                If CBool(PE_CLng(UserSetting(9))) = False Then
                    FieldUpload = False
                End If
            End If
            Set rsFieldUpload = nothing			 
        End If			               			
    End Select	
    Dim strEnableNull
    If EnableNull = False Then
        strEnableNull = " <font color='#FF0000'>*</font>"
    End If
    Response.Write "<tr class='tdbg'><td width='120' align='right' class='tdbg5'><b>" & Title & "��</b></td><td colspan='5'>"
    Select Case FieldType
    Case 1, 8   '�����ı���
        Response.Write "<input type='text' name='" & FieldName & "' size='80' maxlength='255' value='" & strValue & "'>" & strEnableNull
    Case 2, 9    '�����ı���
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
            Response.Write "<input type='text' name='" & FieldName & "' id='" & FieldName & "' size='45' maxlength='255' value='" & strValue & "'><br>" & strEnableNull
        End If
        If PE_CBool(FieldUpload) = True Then		
            Response.Write "<iframe style='top:2px;' id='uploadPhoto' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=fieldpic&FieldName="& FieldName &"' frameborder=0 scrolling=no width='650' height='25'></iframe>"
        End If				
    Case 5   '�ļ�
        If strValue = "" Then
            Response.Write "<input type='text' id='"&FieldName&"' name='"&FieldName&"'  size='45' maxlength='255' value='http://'><br>" & strEnableNull
        Else
            Response.Write "<input type='text' name='" & FieldName & "' id='" & FieldName & "' size='45' maxlength='255' value='" & strValue & "'><br>" & strEnableNull
        End If
        If PE_CBool(FieldUpload) = True Then			
            Response.Write "<iframe style='top:2px' id='uploadsoft' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=fieldsoft&FieldName="& FieldName &"' frameborder=0 scrolling=no width='650' height='25'></iframe>"	
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
            Response.Write "<input type='text' name='" & FieldName & "' onkeyup=""value=value.replace(/[^\d]/g,'')"" size='20' maxlength='20' value='" & PE_Clng(strValue) & "'>" & strEnableNull
        End If		
    End Select
    If IsNull(Tips) = False And Tips <> "" Then
        Response.Write "<br>" & PE_HTMLEncode(Tips)
    End If
    Response.Write "</td></tr>"
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
            rsArticle("Inputer") = Inputer
            rsArticle("Editor") = Inputer
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

            If BlogFlag = True Then 'д��BLOGID
                Dim blogid
                Set blogid = Conn.Execute("select top 1 ID from PE_Space where Type=1 and UserID=" & UserID)
                If blogid.BOF And blogid.EOF Then
                    rsArticle("BlogID") = 0
                Else
                    rsArticle("BlogID") = blogid("ID")
                End If
                Set blogid = Nothing
            End If
            
            rsArticle.Update
            If rsArticle("Status") = 3 Then
                Conn.Execute ("update PE_User set UserExp=UserExp+" & PresentExp & " where UserName='" & UserName & "'")
            End If
        End If
    ElseIf Action = "SaveModify" Then
        If ArticleID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����ȷ��" & ChannelShortName & "ID��ֵ</li>"
        Else
            sql = "select * from PE_Article where Inputer='" & UserName & "' and Deleted=" & PE_False & " and ArticleID=" & ArticleID
            rsArticle.Open sql, Conn, 1, 3
            If rsArticle.BOF And rsArticle.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�Ҳ�����" & ChannelShortName & "�������Ѿ���������ɾ����</li>"
            Else
                If rsArticle("Status") > 0 And EnableModifyDelete = 0 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>" & ChannelShortName & "�Ѿ������ͨ�����������ٽ����޸ģ�</li>"
                Else
                    Conn.Execute ("delete from PE_InfoS where ModuleType=1 and ItemID=" & ArticleID)
                    Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (1," & ArticleID & "," & SpecialID & ")")
                    rsArticle("ClassID") = ClassID
                    rsArticle("Title") = Title
                    rsArticle("Intro") = Intro
                    rsArticle("Content") = Content
                    rsArticle("Keyword") = Keyword
                    rsArticle("Author") = Author
                    rsArticle("CopyFrom") = CopyFrom
                    rsArticle("LinkUrl") = LinkUrl
                    rsArticle("IncludePic") = IncludePic
                    rsArticle("Status") = Status
                    If ShowCommentLink = "Yes" Then
                        rsArticle("ShowCommentLink") = True
                    Else
                        rsArticle("ShowCommentLink") = False
                    End If
                    rsArticle("UpdateTime") = Now()
                    rsArticle("DefaultPicUrl") = DefaultPicUrl
                    rsArticle("UploadFiles") = UploadFiles
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
        End If
    End If
    rsArticle.Close
    Set rsArticle = Nothing
    
    If FoundErr = True Then Exit Sub

    Call UpdateChannelData(ChannelID)
    Call UpdateUserData(0, UserName, 0, 0)
    
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
    Response.Write "��<a href='User_Article.asp?ChannelID=" & ChannelID & "&Action=Modify&ArticleID=" & ArticleID & "'>�޸ı���</a>��&nbsp;"
    Response.Write "��<a href='User_Article.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>�������" & ChannelShortName & "</a>��&nbsp;"
    Response.Write "��<a href='User_Article.asp?ChannelID=" & ChannelID & "&Action=Manage&ClassID=" & ClassID & "'>" & ChannelShortName & "����</a>��&nbsp;"
    Response.Write "��<a href='User_Article.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & ArticleID & "'>Ԥ��" & ChannelShortName & "����</a>��"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf

    Session("Keyword") = Trim(Request("Keyword"))
    Session("Author") = Author
    Session("CopyFrom") = CopyFrom
    Call ClearSiteCache(0)
    Call CreateAllJS_User
End Sub

Sub Del()
    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ѡ��" & ChannelShortName & "��</li>"
        Exit Sub
    End If

    Dim sqlDel, rsDel, NeedUpdateCache
    NeedUpdateCache = False

    sqlDel = "select * from PE_Article where Inputer='" & UserName & "' and Deleted=" & PE_False & " and "
    If InStr(ArticleID, ",") > 0 Then
        sqlDel = sqlDel & " ArticleID in (" & ArticleID & ") order by ArticleID"
    Else
        sqlDel = sqlDel & " ArticleID=" & ArticleID
    End If
    Set rsDel = Server.CreateObject("ADODB.Recordset")
    rsDel.Open sqlDel, Conn, 1, 3
    Do While Not rsDel.EOF
        If rsDel("Status") > 0 Then
            If EnableModifyDelete = 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>ɾ��" & ChannelShortName & "��" & rsDel("Title") & "��ʧ�ܡ�ԭ�򣺴�" & ChannelShortName & "�Ѿ������ͨ������������ɾ����</li>"
            Else
                Conn.Execute ("update PE_User set UserExp=UserExp-" & rsDel("PresentExp") & " where UserName='" & UserName & "'")
                rsDel("Deleted") = True
                rsDel.Update
                NeedUpdateCache = True
            End If
        Else
            rsDel("Deleted") = True
            rsDel.Update
        End If
        rsDel.MoveNext
    Loop
    rsDel.Close
    Set rsDel = Nothing
    
    If NeedUpdateCache = True Then
        Call ClearSiteCache(0)
        Call CreateAllJS_User
    End If

    Call CloseConn
    If FoundErr = False Then
        Response.Redirect ComeUrl
    End If
End Sub

Sub Receive()

    Dim sqlReceive, rsReceive

    If ArticleID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ѡ��" & ChannelShortName & "��</li>"
        Exit Sub
    End If

    sqlReceive = "select * from PE_Article "
    If InStr(ArticleID, ",") > 0 Then
        sqlReceive = sqlReceive & " where ArticleID in (" & ArticleID & ") order by ArticleID"
    Else
        sqlReceive = sqlReceive & " where ArticleID=" & ArticleID
    End If
    Set rsReceive = Server.CreateObject("ADODB.Recordset")
    rsReceive.Open sqlReceive, Conn, 1, 3
    Do While Not rsReceive.EOF
        If FoundInArr(rsReceive("ReceiveUser"), UserName, ",") = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ǩ��" & ChannelShortName & "��" & rsReceive("Title") & "��ʧ�ܡ�ԭ�򣺴�" & ChannelShortName & "����Ҫ��ǩ�գ�</li>"
        End If
        If FoundInArr(rsReceive("Received"), UserName, "|") = True Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ǩ��" & ChannelShortName & "��" & rsReceive("Title") & "��ʧ�ܡ�ԭ�򣺴�" & ChannelShortName & "���Ѿ�ǩ�չ���</li>"
        End If
        If FoundErr = True Then
            rsReceive.Close
            Set rsReceive = Nothing
            Exit Sub
        End If
        If rsReceive("Received") = "" Or IsNull(rsReceive("Received")) Then
            rsReceive("Received") = UserName
        Else
            rsReceive("Received") = rsReceive("Received") & "|" & UserName
        End If
        rsReceive.Update
        rsReceive.MoveNext
    Loop
    rsReceive.Close
    Set rsReceive = Nothing

    Dim sqlUser, rsUser, i, tmpUnsignedItems, tmpArticleID, arrID
    Set rsUser = Server.CreateObject("adodb.recordset")
    sqlUser = "select UserID,UserName,UnsignedItems from PE_User where UserName='" & UserName & "'"
    rsUser.Open sqlUser, Conn, 1, 3
    If Not rsUser.EOF Then
        arrID = Split(ArticleID, ",")
        For i = 0 To UBound(arrID)
            If FoundInArr(rsUser("UnsignedItems"), CStr(arrID(i)), ",") = True Then
                tmpUnsignedItems = "," & rsUser("UnsignedItems") & ","
                tmpArticleID = "," & PE_CLng(Trim(arrID(i))) & ","
                tmpUnsignedItems = Replace(tmpUnsignedItems, tmpArticleID, ",")
                If tmpUnsignedItems = "," Then
                    rsUser("UnsignedItems") = ""
                Else
                    rsUser("UnsignedItems") = Mid(tmpUnsignedItems, 2, Len(tmpUnsignedItems) - 2)
                End If
                rsUser.Update
            End If
        Next
    End If
    rsUser.Close
    Set rsUser = Nothing

    Call CloseConn
    Response.Redirect ComeUrl
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
        If rsArticle("Inputer") <> UserName And FoundInArr(rsArticle("ReceiveUser"), UserName, ",") = False Then
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
    Response.Write "�����ڵ�λ�ã�&nbsp;<a href='User_Article.asp?ChannelID=" & ChannelID & "'>" & ChannelShortName & "����</a>&nbsp;&gt;&gt;&nbsp;"
    If ParentID > 0 Then
        Dim sqlPath, rsPath
        sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
        Set rsPath = Server.CreateObject("adodb.recordset")
        rsPath.Open sqlPath, Conn, 1, 1
        Do While Not rsPath.EOF
            Response.Write "<a href='User_Article.asp?ChannelID=" & ChannelID & "&ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    Response.Write "<a href='User_Article.asp?ChannelID=" & ChannelID & "&ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
    Response.Write "<a href='User_Article.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & rsArticle("ArticleID") & "'>"
    
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
    If FoundInArr(rsArticle("ReceiveUser"), UserName, ",") = True And FoundInArr(rsArticle("Received"), UserName, "|") = True Then
        Response.Write "&nbsp;&nbsp;<span id='ReceiveState' style='color:green'>����ǩ�ա�</font>"
    End If
    If FoundInArr(rsArticle("ReceiveUser"), UserName, ",") = True And FoundInArr(rsArticle("Received"), UserName, "|") = False Then
        Response.Write "&nbsp;&nbsp;<span id='ReceiveState' style='color:red'>��δǩ�ա�</font>"
    End If
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
    Response.Write "" & ChannelShortName & "¼�룺<a href='User_Article.asp?ChannelID=" & ChannelID & "&Field=Inputer&Keyword=" & rsArticle("Inputer") & "'>" & rsArticle("Inputer") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;���α༭��"
    If rsArticle("Status") > 0 Then
        Response.Write "" & rsArticle("Editor")
    Else
        Response.Write "��"
    End If
    Response.Write " </td>"
    Response.Write "  </tr>"
    If rsArticle("Receive") = True Then
        Response.Write "  <tr class='tdbg'> "
        Response.Write "    <td colspan='2' height='40'>"
        Response.Write "<table width='100%'><tr><td width='100'>Ҫ��ǩ�յ��û���</td><td align='left' style='width:600; word-wrap:break-word;'>" & Replace(rsArticle("ReceiveUser"), ",", ",") & "</td></tr><tr><td>�Ѿ�ǩ�յ��û���</td><td style='width:600; word-wrap:break-word;'>" & Replace(rsArticle("Received"), "|", ",") & "</td></tr><tr><td>��δǩ�յ��û���</td><td style='width:600; word-wrap:break-word;'>"
        Dim NotReceiveUser, arrUser
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
        Response.Write "" & NotReceiveUser & "</td></tr></table>"
        Response.Write "</td>"
        Response.Write "  </tr>"
    End If
    Response.Write "</table>" & vbCrLf
    Response.Write "<form name='formA' method='get' action='User_Article.asp'><p align='center'> "
    Response.Write "<input type='hidden' name='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input type='hidden' name='ArticleID' value='" & ArticleID & "'>"
    Response.Write "<input type='hidden' name='Action' value=''>" & vbCrLf
    If rsArticle("Inputer") = UserName And (rsArticle("Status") <= 0 Or UserSetting(2) = 1) Then
        Response.Write "<input type='Submit' name='button1' value=' �� �� ' onclick=""document.formA.Action.value='Modify'"">&nbsp;&nbsp;"
        Response.Write "<input type='Submit' name='button2' value=' ɾ �� ' onclick=""if(confirm('ȷ��Ҫɾ����" & ChannelShortName & "��')==true){document.formA.Action.value='Del';}"">"
    End If
    If FoundInArr(rsArticle("ReceiveUser"), UserName, ",") = True And FoundInArr(rsArticle("Received"), UserName, "|") = False Then
        Response.Write "&nbsp;&nbsp;<input type='Submit' name='ReceiveButton' id='ReceiveButton' value=' ǩ �� ' onclick=""document.formA.Action.value='Receive'"" style='display:'>"
        If rsArticle("AutoReceiveTime") > 0 Then
            Call ShowJS_SignIn(ArticleID, rsArticle("AutoReceiveTime"))
        End If
    End If
    rsArticle.Close
    Set rsArticle = Nothing
    Response.Write "</Form></p>"

    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='0'><tr class='tdbg'><td>"
    Response.Write "<li>��һ" & ChannelItemUnit & ChannelShortName & "��"
    Dim rsPrev
    sql = "Select Top 1 A.ArticleID,A.Title,C.ClassID,C.ClassName from PE_Article A left join PE_Class C On C.ClassID=A.ClassID Where A.ChannelID=" & ChannelID & " and A.Deleted=" & PE_False & " and A.Inputer='" & UserName & "' and ArticleID<" & ArticleID & " order by ArticleID desc"
    Set rsPrev = Server.CreateObject("ADODB.Recordset")
    rsPrev.Open sql, Conn, 1, 1
    If rsPrev.EOF Then
        Response.Write "û����"
    Else
        Response.Write "[<a href='User_Article.asp?ChannelID=" & ChannelID & "&ClassID=" & rsPrev("ClassID") & "'>" & rsPrev("ClassName") & "</a>] <a href='User_Article.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & rsPrev("ArticleID") & "'>" & rsPrev("Title") & "</a>"
    End If
    rsPrev.Close
    Set rsPrev = Nothing
    Response.Write "</li></td</tr><tr class='tdbg'><td><li>��һ" & ChannelItemUnit & ChannelShortName & "��"
    Dim rsNext
    sql = "Select Top 1 A.ArticleID,A.Title,C.ClassID,C.ClassName from PE_Article A left join PE_Class C On C.ClassID=A.ClassID Where A.ChannelID=" & ChannelID & " and A.Deleted=" & PE_False & " and A.Inputer='" & UserName & "' and ArticleID>" & ArticleID & " order by ArticleID asc"
    Set rsNext = Server.CreateObject("ADODB.Recordset")
    rsNext.Open sql, Conn, 1, 1
    If rsNext.EOF Then
        Response.Write "û����"
    Else
        Response.Write "[<a href='User_Article.asp?ChannelID=" & ChannelID & "&ClassID=" & rsNext("ClassID") & "'>" & rsNext("ClassName") & "</a>] <a href='User_Article.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & rsNext("ArticleID") & "'>" & rsNext("Title") & "</a>"
    End If
    rsNext.Close
    Set rsNext = Nothing
    Response.Write "</li></td></tr></table><br>" & vbCrLf

End Sub

Sub ShowJS_SignIn(iArticleID, iAutoReceiveTime)
    Response.Write "<script type='text/javascript' language='javascript'>" & vbCrLf
    Response.Write "var secs;" & vbCrLf
    Response.Write "var timerID = null;" & vbCrLf
    Response.Write "var timerRunning = false;" & vbCrLf
    Response.Write "var delay = 1000;" & vbCrLf
    Response.Write "function InitializeTimer()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    secs = " & iAutoReceiveTime & ";" & vbCrLf
    Response.Write "    StopTheClock();" & vbCrLf
    Response.Write "    StartTheTimer();" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function StopTheClock()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if(timerRunning)" & vbCrLf
    Response.Write "        clearTimeout(timerID);" & vbCrLf
    Response.Write "    timerRunning = false;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function StartTheTimer()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if (secs==0)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        StopTheClock();" & vbCrLf
    Response.Write "        self.status = '';" & vbCrLf
    Response.Write "        makeRequest('User_ArticleReceive.asp?ArticleID=" & iArticleID & "');" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    else" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        self.status = secs;" & vbCrLf
    Response.Write "        secs = secs - 1;" & vbCrLf
    Response.Write "        timerRunning = true;" & vbCrLf
    Response.Write "        timerID = self.setTimeout('StartTheTimer()', delay);" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "var http_request = false;" & vbCrLf
    Response.Write "function makeRequest(url) {" & vbCrLf
    Response.Write "  http_request = false;" & vbCrLf
    Response.Write "  if (window.XMLHttpRequest) { // Mozilla, Safari,..." & vbCrLf
    Response.Write "    http_request = new XMLHttpRequest();" & vbCrLf
    Response.Write "    if (http_request.overrideMimeType) {" & vbCrLf
    Response.Write "      http_request.overrideMimeType('text/xml');" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  } else if (window.ActiveXObject) { // IE" & vbCrLf
    Response.Write "    try {" & vbCrLf
    Response.Write "      http_request = new ActiveXObject('Msxml2.XMLHTTP');" & vbCrLf
    Response.Write "    } catch (e) {" & vbCrLf
    Response.Write "      try {" & vbCrLf
    Response.Write "        http_request = new ActiveXObject('Microsoft.XMLHTTP');" & vbCrLf
    Response.Write "      } catch (e) {}" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (!http_request) {" & vbCrLf
    Response.Write "    alert('Giving up :( Cannot create an XMLHTTP instance');" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  http_request.onreadystatechange = SignIn;" & vbCrLf
    Response.Write "  http_request.open('GET', url, true);" & vbCrLf
    Response.Write "  http_request.send(null);" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function SignIn() {" & vbCrLf
    Response.Write "  if (http_request.readyState == 4) {" & vbCrLf
    Response.Write "    if (http_request.status == 200) {" & vbCrLf
    Response.Write "      if (http_request.responseText == 'OK') {" & vbCrLf
    Response.Write "        document.getElementById('ReceiveButton').style.display='none';" & vbCrLf
    Response.Write "        document.getElementById('ReceiveState').style.color='green';" & vbCrLf
    Response.Write "        document.getElementById('ReceiveState').innerHTML='�����Զ�ǩ�ա�';" & vbCrLf
    Response.Write "      }" & vbCrLf
    Response.Write "    } else {" & vbCrLf
    Response.Write "      alert('There was a problem with the request.');" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "InitializeTimer();" & vbCrLf
    Response.Write "</script>" & vbCrLf
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

Function GetReceiveArticleID()
    Dim rsReceive, sqlReceive, strArticleID
    Set rsReceive = Server.CreateObject("ADODB.Recordset")
    sqlReceive = "select ArticleID,ReceiveUser from PE_Article where Receive=" & PE_True & " and Status=3 order by ArticleID desc"
    rsReceive.Open sqlReceive, Conn, 1, 1
    Do While Not rsReceive.EOF
        If FoundInArr(rsReceive("ReceiveUser"), UserName, ",") = True Then
            If strArticleID = "" Or IsNull(strArticleID) Then
                strArticleID = rsReceive("ArticleID")
            Else
                strArticleID = strArticleID & "," & rsReceive("ArticleID")
            End If
        End If
        rsReceive.MoveNext
    Loop
    rsReceive.Close
    Set rsReceive = Nothing
    If strArticleID = "" Or IsNull(strArticleID) Then
        GetReceiveArticleID = "0"
    Else
        GetReceiveArticleID = strArticleID
    End If
End Function

%>
