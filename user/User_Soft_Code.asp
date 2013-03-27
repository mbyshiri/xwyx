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

Dim SoftID, AuthorName, Status, ManageType
Dim ClassID, SpecialID, OnTop, IsElite, IsHot, Created

Dim tClass, ClassName, RootID, ParentID, Depth, ParentPath, Child, arrChildID, ParentDir, ClassDir, ClassPurview
Dim arrFields_Options, arrSoftType, arrSoftLanguage, arrCopyrightType, arrOperatingSystem
    

Sub Execute()
    ChannelID = PE_CLng(Trim(Request("ChannelID")))
    If ChannelID > 0 Then
        Call GetChannel(ChannelID)
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�鿴��Ƶ��ID��</li>"
        Response.Write ErrMsg
        Exit Sub
    End If
    SoftID = Trim(Request("SoftID"))
    ClassID = PE_CLng(Trim(Request("ClassID")))
    Status = Trim(Request("Status"))
    AuthorName = Trim(Request("AuthorName"))
    strField = Trim(Request("Field"))
    If Status = "" Then
        Status = 9
    Else
        Status = PE_CLng(Status)
    End If
    If IsValidID(SoftID) = False Then
        SoftID = ""
    End If

    arrFields_Options = Split(",,,", ",")
    arrSoftType = ""
    arrSoftLanguage = ""
    arrCopyrightType = ""
    arrOperatingSystem = ""
    If Fields_Options & "" <> "" Then
        arrFields_Options = Split(Fields_Options, "$$$")
        If UBound(arrFields_Options) = 3 Then
            arrSoftType = Split(arrFields_Options(0), vbCrLf)
            arrSoftLanguage = Split(arrFields_Options(1), vbCrLf)
            arrCopyrightType = Split(arrFields_Options(2), vbCrLf)
            arrOperatingSystem = Split(arrFields_Options(3), vbCrLf)
        End If
    End If

    If Action = "" Then Action = "Manage"
    FileName = "User_Soft.asp?ChannelID=" & ChannelID & "&Action=" & Action & "&ManageType=" & ManageType
    strFileName = FileName & "&ClassID=" & ClassID & "&Field=" & strField & "&keyword=" & Keyword
    If AuthorName <> "" Then
        AuthorName = ReplaceBadChar(AuthorName)
        strFileName = strFileName & "&AuthorName=" & AuthorName
    End If

    If XmlDoc.Load(Server.MapPath(InstallDir & "Language/Gb2312_Channel_" & ChannelID & ".xml")) = False Then XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))

    Response.Write "<table align='center'><tr align='center' valign='top'>"
    If CheckUser_ChannelInput() Then
        Response.Write "<td width='90'><a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Add'><img src='images/Soft_add.gif' border='0' align='absmiddle'><br>���" & ChannelShortName & "</a></td>"
    End If
    Response.Write "<td width='90'><a href='User_Soft.asp?ChannelID=" & ChannelID & "&Status=9'><img src='images/Soft_all.gif' border='0' align='absmiddle'><br>����" & ChannelShortName & "</a></td>"
    Response.Write "<td width='90'><a href='User_Soft.asp?ChannelID=" & ChannelID & "&Status=-1'><img src='images/Soft_draft.gif' border='0' align='absmiddle'><br>�� ��</a></td>"
    Response.Write "<td width='90'><a href='User_Soft.asp?ChannelID=" & ChannelID & "&Status=0'><img src='images/Soft_unpassed.gif' border='0' align='absmiddle'><br>����˵�" & ChannelShortName & "</a></td>"
    Response.Write "<td width='90'><a href='User_Soft.asp?ChannelID=" & ChannelID & "&Status=3'><img src='images/Soft_passed.gif' border='0' align='absmiddle'><br>����˵�" & ChannelShortName & "</a></td>"
    Response.Write "<td width='90'><a href='User_Soft.asp?ChannelID=" & ChannelID & "&Status=-2'><img src='images/Soft_reject.gif' border='0' align='absmiddle'><br>δ�����õ�" & ChannelShortName & "</a></td>"
    Response.Write "</tr></table>" & vbCrLf

    Select Case Action
    Case "Add"
        Call Add
    Case "Modify"
        Call Modify
    Case "SaveAdd", "SaveModify"
        Call SaveSoft
    Case "Preview"
        Call Preview
    Case "Show"
        Call Show
    Case "Del"
        Call Del
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
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22'>" & GetRootClass() & "</td>"
    Response.Write "  </tr>" & GetChild_Root() & ""
    Response.Write "</table><br>"

    Call ShowContentManagePath(ChannelShortName & "����")

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "    <form name='myform' method='Post' action='User_Soft.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "     <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "          <tr class='title' height='22'> "
    Response.Write "            <td height='22' width='30' align='center'><strong>ѡ��</strong></td>"
    Response.Write "            <td width='25' align='center'><strong>ID</strong></td>"
    Response.Write "            <td align='center' ><strong>" & ChannelShortName & "����</strong></td>"
    Response.Write "            <td width='60' align='center' ><strong>¼��</strong></td>"
    Response.Write "            <td width='40' align='center' ><strong>������</strong></td>"
    Response.Write "            <td width='80' align='center' ><strong>" & ChannelShortName & "����</strong></td>"
    Response.Write "            <td width='60' align='center' ><strong>���״̬</strong></td>"
    Response.Write "            <td width='80' align='center' ><strong>�������</strong></td>"
    Response.Write "          </tr>"

    Dim rsSoftList, sql
    sql = "select S.SoftID,S.ClassID,C.ClassName,C.ParentDir,C.ClassDir,S.SoftName,S.SoftVersion,S.Author,S.Keyword,S.UpdateTime,S.Inputer,S.Editor,S.Hits,S.SoftSize,S.OnTop,S.Elite,S.Status,S.Stars,S.InfoPoint from PE_Soft S"
    sql = sql & " left join PE_Class C on S.ClassID=C.ClassID where S.ChannelID=" & ChannelID & " and S.Deleted=" & PE_False & " and S.Inputer='" & UserName & "' "
    If AuthorName <> "" Then
        sql = sql & " and S.Author='" & AuthorName & "|' "
    End If
    Select Case Status
    Case 3
        sql = sql & " and S.Status=3"
    Case 0
        sql = sql & " and (S.Status=0 Or S.Status=1 Or S.Status=2)"
    Case -1
        sql = sql & " and S.Status=-1"
    Case -2
        sql = sql & " and S.Status=-2"
    End Select
    If ClassID > 0 Then
        If Child > 0 Then
            sql = sql & " and S.ClassID in (" & arrChildID & ")"
        Else
            sql = sql & " and S.ClassID=" & ClassID
        End If
    End If

    If Keyword <> "" Then
        Select Case strField
        Case "SoftName"
            sql = sql & " and S.SoftName like '%" & Keyword & "%' "
        Case "SoftIntro"
            sql = sql & " and S.SoftIntro like '%" & Keyword & "%' "
        Case "Author"
            sql = sql & " and S.Author like '%" & Keyword & "%' "
        Case "Inputer"
            sql = sql & " and S.Inputer='" & Keyword & "' "
        Case Else
            sql = sql & " and S.SoftName like '%" & Keyword & "%' "
        End Select
    End If
    sql = sql & " order by S.SoftID desc"

    Set rsSoftList = Server.CreateObject("ADODB.Recordset")
    rsSoftList.Open sql, Conn, 1, 1
    If rsSoftList.BOF And rsSoftList.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>" & GetStrNoItem(ClassID, Status) & "<br><br></td></tr>"
    Else
        totalPut = rsSoftList.RecordCount
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
                rsSoftList.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Dim SoftNum
        SoftNum = 0
        Do While Not rsSoftList.EOF
            Response.Write "      <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "        <td width='30' align='center'><input name='SoftID' type='checkbox' onclick='unselectall()' id='SoftID' value='" & rsSoftList("SoftID") & "'></td>"
            Response.Write "        <td width='25' align='center'>" & rsSoftList("SoftID") & "</td>"
            Response.Write "        <td>"
            If rsSoftList("ClassID") <> ClassID Then
                Response.Write "<a href='" & FileName & "&ClassID=" & rsSoftList("ClassID") & "'>[" & rsSoftList("ClassName") & "]</a>&nbsp;"
            End If
            Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Show&SoftID=" & rsSoftList("SoftID") & "' title='" & GetLinkTips(rsSoftList("SoftName"), rsSoftList("SoftVersion"), rsSoftList("Author"), rsSoftList("UpdateTime"), rsSoftList("Hits"), rsSoftList("Keyword"), rsSoftList("Stars"), rsSoftList("InfoPoint")) & "'>" & rsSoftList("SoftName")
            If rsSoftList("SoftVersion") <> "" Then
                Response.Write "&nbsp;&nbsp;" & rsSoftList("SoftVersion")
            End If
            Response.Write "</a></td>"
            Response.Write "            <td width='60' align='center'><a href='" & FileName & "&field=Inputer&keyword=" & rsSoftList("Inputer") & "' title='������鿴���û�¼�������" & ChannelShortName & "'>" & rsSoftList("Inputer") & "</a></td>"
            Response.Write "            <td width='40' align='center'>" & rsSoftList("Hits") & "</td>"
            Response.Write "            <td width='80' align='center'>" & GetInfoProperty(rsSoftList("OnTop"), rsSoftList("Hits"), rsSoftList("Elite")) & "</td>"
            Response.Write "            <td width='60' align='center'>" & GetInfoStatus(rsSoftList("Status")) & "</td>"
            Response.Write "    <td width='80' align='center'>"
            If rsSoftList("Inputer") = UserName And (rsSoftList("Status") <= 0 Or EnableModifyDelete = 1) Then
                Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Modify&SoftID=" & rsSoftList("SoftID") & "'>�޸�</a>&nbsp;"
                Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Del&SoftID=" & rsSoftList("SoftID") & "' onclick=""return confirm('ȷ��Ҫɾ����" & ChannelShortName & "��һ��ɾ�������ָܻ���');"">ɾ��</a>"
            End If
            Response.Write "</td>"
            Response.Write "</tr>"

            SoftNum = SoftNum + 1
            If SoftNum >= MaxPerPage Then Exit Do
            rsSoftList.MoveNext
        Loop
    End If
    rsSoftList.Close
    Set rsSoftList = Nothing
    Response.Write "</table>"

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>ѡ�б�ҳ��ʾ������" & ChannelShortName & "</td><td>"
    Response.Write "<input name='submit1' type='submit' value='ɾ��ѡ����" & ChannelShortName & "' onClick=""document.myform.Action.value='Del'"" >"
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
    Response.Write "<option value='SoftName' selected>" & ChannelShortName & "����</option>"
    Response.Write "<option value='SoftIntro'>" & ChannelShortName & "���</option>"
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
    Response.Write "<br><b>˵����</b><br>&nbsp;&nbsp;&nbsp;&nbsp;" & ChannelShortName & "�����еĸ���壺<font color=blue>��</font>----�̶�" & ChannelShortName & "��<font color=red>��</font>----����" & ChannelShortName & "��<font color=green>��</font>----�Ƽ�" & ChannelShortName & "<br><br>"
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

Function GetLinkTips(SoftName, SoftVersion, Author, UpdateTime, Hits, Keyword, Stars, InfoPoint)
    Dim strLinkTips
    strLinkTips = ""
    strLinkTips = strLinkTips & "��&nbsp;&nbsp;&nbsp;&nbsp;�ƣ�" & SoftName & vbCrLf
    strLinkTips = strLinkTips & "��&nbsp;&nbsp;&nbsp;&nbsp;����" & SoftVersion & vbCrLf
    strLinkTips = strLinkTips & "��&nbsp;&nbsp;&nbsp;&nbsp;�ߣ�" & Author & vbCrLf
    strLinkTips = strLinkTips & "����ʱ�䣺" & UpdateTime & vbCrLf
    strLinkTips = strLinkTips & "���ش�����" & Hits & vbCrLf
    strLinkTips = strLinkTips & "�� �� �֣�" & Mid(Keyword, 2, Len(Keyword) - 2) & vbCrLf
    strLinkTips = strLinkTips & "�Ƽ��ȼ���"
    If Stars = 0 Then
        strLinkTips = strLinkTips & "��"
    Else
        strLinkTips = strLinkTips & String(Stars, "��")
    End If
    strLinkTips = strLinkTips & vbCrLf
    strLinkTips = strLinkTips & "���ص�����" & InfoPoint
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

Function GetInfoProperty(OnTop, Hits, Elite)
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
    GetInfoProperty = strInfoProperty
End Function

Sub ShowJS_Soft()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function AddUrl(){" & vbCrLf
    Response.Write "  var thisurl='" & XmlText("Soft", "DownloadUrlTip", "���ص�ַ") & "'+(document.myform.DownloadUrl.length+1)+'|http://'; " & vbCrLf
    Response.Write "  var url=prompt('���������ص�ַ���ƺ����ӣ��м��á�|��������',thisurl);" & vbCrLf
    Response.Write "  if(url!=null&&url!=''){document.myform.DownloadUrl.options[document.myform.DownloadUrl.length]=new Option(url,url);}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function ModifyUrl(){" & vbCrLf
    Response.Write "  if(document.myform.DownloadUrl.length==0) return false;" & vbCrLf
    Response.Write "  var thisurl=document.myform.DownloadUrl.value; " & vbCrLf
    Response.Write "  if (thisurl=='') {alert('����ѡ��һ�����ص�ַ���ٵ��޸İ�ť��');return false;}" & vbCrLf
    Response.Write "  var url=prompt('���������ص�ַ���ƺ����ӣ��м��á�|��������',thisurl);" & vbCrLf
    Response.Write "  if(url!=thisurl&&url!=null&&url!=''){document.myform.DownloadUrl.options[document.myform.DownloadUrl.selectedIndex]=new Option(url,url);}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function DelUrl(){" & vbCrLf
    Response.Write "  if(document.myform.DownloadUrl.length==0) return false;" & vbCrLf
    Response.Write "  var thisurl=document.myform.DownloadUrl.value; " & vbCrLf
    Response.Write "  if (thisurl=='') {alert('����ѡ��һ�����ص�ַ���ٵ�ɾ����ť��');return false;}" & vbCrLf
    Response.Write "  document.myform.DownloadUrl.options[document.myform.DownloadUrl.selectedIndex]=null;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckForm()" & vbCrLf
    Response.Write "{" & vbCrLf
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
    Response.Write "  if (document.myform.SoftName.value==''){" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "���Ʋ���Ϊ�գ�');" & vbCrLf
    Response.Write "    document.myform.SoftName.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Keyword.value==''){" & vbCrLf
    Response.Write "    alert('�ؼ��ֲ���Ϊ�գ�');" & vbCrLf
    Response.Write "    document.myform.Keyword.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  document.myform.SoftIntro.value=editor.HtmlEdit.document.body.innerHTML; " & vbCrLf
    Response.Write "  if (document.myform.SoftIntro.value==''){" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "��鲻��Ϊ�գ�');" & vbCrLf
    Response.Write "    editor.HtmlEdit.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.DownloadUrl.length==0){" & vbCrLf
    Response.Write "    alert('" & ChannelShortName & "���ص�ַ����Ϊ�գ�');" & vbCrLf
    Response.Write "    document.myform.DownloadUrl.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  document.myform.DownloadUrls.value=''" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.DownloadUrl.length;i++){" & vbCrLf
    Response.Write "    if (document.myform.DownloadUrls.value=='') document.myform.DownloadUrls.value=document.myform.DownloadUrl.options[i].value;" & vbCrLf
    Response.Write "    else document.myform.DownloadUrls.value+='$$$'+document.myform.DownloadUrl.options[i].value;" & vbCrLf
    Response.Write "  }" & vbCrLf
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
        Set trs = Conn.Execute("select count(SoftID) from PE_Soft where Inputer='" & UserName & "' and Deleted=" & PE_False & " and DateDiff(" & PE_DatePart_D & "," & PE_Now & ",UpdateTime)=0")
        If trs(0) >= MaxPerDay Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�����췢���" & ChannelShortName & "�Ѿ��ﵽ�����ޣ�</li>"
        End If
        Set trs = Nothing
        If FoundErr = True Then Exit Sub
    End If
    
    Call ShowJS_Soft
    
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='User_Soft.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center' colspan='2'><b>���" & ChannelShortName & "</td>"
    Response.Write "    </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>������Ŀ��</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <select name='ClassID'>" & User_GetClass_Option(4, ClassID) & "</select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>����ר�⣺</strong></td>"
    Response.Write "            <td><select name='SpecialID'><option value='0'>�������κ�ר��</option>" & GetSpecial_Option(0) & "</select></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "���ƣ�</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='SoftName' type='text' value='' size='50' maxlength='255'> <font color='#FF0000'>*</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>�ؼ��֣�</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Keyword' type='text' id='Keyword' size='50' maxlength='255'> <font color='#FF0000'>*</font> " & GetKeywordList("User", ChannelID)
    Response.Write "              <br><font color='#0000FF'>�����������" & ChannelShortName & "�����������ؼ��֣��м���<font color='#FF0000'>��|��</font>���������ܳ���&quot;'&?;:()���ַ���</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>����/�����̣�</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Author' type='text' id='Author' value='" & Trim(Session("Author")) & "' size='50' maxlength='30'>" & GetAuthorList("User", ChannelID, UserName)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    If FoundInArr(arrEnabledTabs, "SoftParameter", ",") = True Then
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "�汾��</strong></td>"
        Response.Write "            <td><input name='SoftVersion' type='text' size='15' maxlength='100'></td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "��Դ��</strong></td>"
        Response.Write "            <td>"
        Response.Write "              <input name='CopyFrom' type='text' id='CopyFrom' value='" & Trim(Session("CopyFrom")) & "' size='50' maxlength='100'>" & GetCopyFromList("User", ChannelID)
        Response.Write "            </td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "���</strong></td>"
        Response.Write "            <td>"
        Response.Write "              <select name='SoftType' id='SoftType'>" & GetSoftType(0) & "</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "              <strong>" & ChannelShortName & "���ԣ�</strong> <select name='SoftLanguage' id='SoftLanguage'>" & GetSoftLanguage(0) & "</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "              <strong>��Ȩ��ʽ��</strong> <select name='CopyrightType' id='CopyrightType'>" & GetCopyrightType(0) & "</select>"
        Response.Write "            </td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "ƽ̨��</strong></td>"
        Response.Write "            <td>"
        Response.Write "              <input name='OperatingSystem' type='text' value='" & XmlText("Soft", "OperatingSystem", "Win9x/NT/2000/XP/") & "' size='80' maxlength='200'> <br>" & GetOperatingSystemList
        Response.Write "            </td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "��ʾ��ַ��</strong></td>"
        Response.Write "            <td><input name='DemoUrl' type='text' value='http://' size='80' maxlength='200'></td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "ע���ַ��</strong></td>"
        Response.Write "            <td><input name='RegUrl' type='text' value='http://' size='80' maxlength='200'></td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>��ѹ���룺</strong></td>"
        Response.Write "            <td><input name='DecompressPassword' type='text' id='DecompressPassword' size='30' maxlength='30'></td>"
        Response.Write "          </tr>"
    End If
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "ͼƬ��</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='SoftPicUrl' type='text' id='SoftPicUrl' size='80' maxlength='200'>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'></td>"
    Response.Write "            <td><table><tr><td>�ϴ�" & ChannelShortName & "ͼƬ��</td><td><iframe style='top:2px' id='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=softpic' frameborder=0 scrolling=no width='450' height='25'></iframe></td></tr></table></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5' valign='middle'><strong>" & ChannelShortName & "��飺</strong></td>"
    Response.Write "            <td><textarea name='SoftIntro' cols='80' rows='10' id='SoftIntro' style='display:none'></textarea>"
    Response.Write "                <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=3&tContentid=SoftIntro' frameborder='1' scrolling='no' width='650' height='200' ></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>�ϴ�" & ChannelShortName & "��</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <iframe style='top:2px' ID='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=soft' frameborder=0 scrolling=no width='450' height='25'></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"	
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "��ַ��</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "                <tr>"
    Response.Write "                  <td>"
    Response.Write "                    <input type='hidden' name='DownloadUrls' value=''>"
    Response.Write "                    <select name='DownloadUrl' style='width:400;height:100' size='2' ondblclick='return ModifyUrl();'></select>"
    Response.Write "                  </td>"
    Response.Write "                  <td>"
    Response.Write "                    <input type='button' name='addurl' value='����ⲿ��ַ' onclick='AddUrl();'><br>"
    Response.Write "                    <input type='button' name='modifyurl' value='�޸ĵ�ǰ��ַ' onclick='return ModifyUrl();'><br>"
    Response.Write "                    <input type='button' name='delurl' value='ɾ����ǰ��ַ' onclick='DelUrl();'>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "              <tr><td  colspan='3'>ϵͳ�ṩ���ϴ�����ֻ�ʺ��ϴ��Ƚ�С��" & ChannelShortName & "����ASPԴ����ѹ�����������" & ChannelShortName & "�Ƚϴ�" & MaxFileSize \ 1024 & "M���ϣ�������ʹ��FTP�ϴ�������Ҫʹ��ϵͳ�ṩ���ϴ����ܣ������ϴ���������ռ�÷�������CPU��Դ��FTP�ϴ����뽫��ַ���Ƶ�����ĵ�ַ���С�</td></tr>"		
    Response.Write "              </table>"
    Response.Write "            </td>"
    Response.Write "          </tr>"

    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "��С��</strong></td>"
    Response.Write "            <td><input name='SoftSize' type='text' id='SoftSize' size='10' maxlength='10'> K</strong></td>"
    Response.Write "          </tr>"
    
    '�Զ����ֶ�
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & ChannelID & " or ChannelID=-2")
    Do While Not rsField.EOF
        IF rsField("ShowOnForm") = True then
            Call WriteFieldHTML(rsField("FieldName"), rsField("Title"), rsField("Tips"), rsField("FieldType"), rsField("DefaultValue"), rsField("Options"), rsField("EnableNull"))
        End If
        rsField.MoveNext
    Loop
    Set rsField = Nothing
    Response.Write "          <tr class='tdbg' class='tdbg5'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "״̬��</strong></td>"
    Response.Write "            <td><input name='Status' type='radio' id='Status' value='-1'>�ݸ�&nbsp;&nbsp;<input Name='Status' Type='Radio' Id='Status' Value='0' checked>Ͷ��</td>"
    Response.Write "          </tr>"
    Response.Write "  </table>"
    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='add' type='submit'  id='Add' value=' �� �� ' onClick=""document.myform.Action.value='SaveAdd';document.myform.target='_self';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Preview' type='submit'  id='Preview' value=' Ԥ �� ' onClick=""document.myform.Action.value='Preview';document.myform.target='_blank';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
End Sub

Sub Modify()
    If FoundInArr(arrClass_Input, ChannelDir & "none", ",") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�Բ�����û����" & ChannelName & "���" & ChannelShortName & "��Ȩ�ޣ�</li><br><br>"
        Exit Sub
    End If
    Dim rsSoft, sql, tmpAuthor, tmpCopyFrom, SpecialID
    
    If SoftID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵ�" & ChannelShortName & "ID</li>"
        Exit Sub
    Else
        SoftID = PE_CLng(SoftID)
    End If
    sql = "select * from PE_Soft where Inputer='" & UserName & "' and Deleted=" & PE_False & " and SoftID=" & SoftID & ""
    Set rsSoft = Server.CreateObject("ADODB.Recordset")
    rsSoft.Open sql, Conn, 1, 1
    If rsSoft.BOF And rsSoft.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���" & ChannelShortName & "</li>"
    Else
        If rsSoft("Status") > 0 And EnableModifyDelete = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & ChannelShortName & "�Ѿ������ͨ����������У��������ٽ����޸ģ�</li>"
        End If
    End If
    If FoundErr = True Then
        rsSoft.Close
        Set rsSoft = Nothing
        Exit Sub
    End If
    SpecialID = PE_CLng(Conn.Execute("select top 1 SpecialID from PE_InfoS where ModuleType=2 and ItemID=" & SoftID & "")(0))

    If Right(rsSoft("Author"), 1) = "|" Then
        tmpAuthor = Left(rsSoft("Author"), Len(rsSoft("Author")) - 1)
    Else
        tmpAuthor = rsSoft("Author")
    End If
    If Right(rsSoft("CopyFrom"), 1) = "|" Then
        tmpCopyFrom = Left(rsSoft("CopyFrom"), Len(rsSoft("CopyFrom")) - 1)
    Else
        tmpCopyFrom = rsSoft("CopyFrom")
    End If
    Call ShowJS_Soft

    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='User_Soft.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center' colspan='2'><b>�޸�" & ChannelShortName & "</b></td>"
    Response.Write "    </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>������Ŀ��</strong></td>"
    Response.Write "            <td colspan='2'>"
    Response.Write "              <select name='ClassID'>" & User_GetClass_Option(4, rsSoft("ClassID")) & "</select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>����ר�⣺</strong></td>"
    Response.Write "            <td><select name='SpecialID'><option value='0'>�������κ�ר��</option>" & GetSpecial_Option(SpecialID) & "</select></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "���ƣ�</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='SoftName' type='text' value='" & rsSoft("SoftName") & "' size='50' maxlength='255'> <font color='#FF0000'>*</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>�ؼ��֣�</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Keyword' type='text' id='Keyword' value='" & Mid(rsSoft("Keyword"), 2, Len(rsSoft("Keyword")) - 2) & "' size='50' maxlength='255'> <font color='#FF0000'>*</font> " & GetKeywordList("User", ChannelID)
    Response.Write "              <br><font color='#0000FF'>�����������" & ChannelShortName & "�����������ؼ��֣��м���<font color='#FF0000'>��|��</font>���������ܳ���&quot;'&?;:()���ַ���</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>����/�����̣�</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Author' type='text' id='Author' value='" & tmpAuthor & "' size='50' maxlength='30'>" & GetAuthorList("User", ChannelID, UserName)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    If FoundInArr(arrEnabledTabs, "SoftParameter", ",") = True Then
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "�汾��</strong></td>"
        Response.Write "            <td><input name='SoftVersion' type='text' size='15' maxlength='100' value='" & rsSoft("SoftVersion") & "'></td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "��Դ��</strong></td>"
        Response.Write "            <td>"
        Response.Write "              <input name='CopyFrom' type='text' id='CopyFrom' value='" & tmpCopyFrom & "' size='50' maxlength='100'>" & GetCopyFromList("User", ChannelID)
        Response.Write "            </td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "���</strong></td>"
        Response.Write "            <td>"
        Response.Write "              <select name='SoftType' id='SoftType'>" & GetSoftType(rsSoft("SoftType")) & "</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "              <strong>" & ChannelShortName & "���ԣ�</strong> <select name='SoftLanguage' id='SoftLanguage'>" & GetSoftLanguage(rsSoft("SoftLanguage")) & "</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "              <strong>��Ȩ��ʽ��</strong> <select name='CopyrightType' id='CopyrightType'>" & GetCopyrightType(rsSoft("CopyrightType")) & "</select>"
        Response.Write "            </td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "ƽ̨��</strong></td>"
        Response.Write "            <td>"
        Response.Write "              <input name='OperatingSystem' type='text' value='" & rsSoft("OperatingSystem") & "' size='80' maxlength='200'> <br>" & GetOperatingSystemList
        Response.Write "            </td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "��ʾ��ַ��</strong></td>"
        Response.Write "            <td><input name='DemoUrl' type='text' value='" & rsSoft("DemoUrl") & "' size='80' maxlength='200'></td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "ע���ַ��</strong></td>"
        Response.Write "            <td><input name='RegUrl' type='text' value='" & rsSoft("RegUrl") & "' size='80' maxlength='200'></td>"
        Response.Write "          </tr>"
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'><strong>��ѹ���룺</strong></td>"
        Response.Write "            <td><input name='DecompressPassword' type='text' id='DecompressPassword' value='" & rsSoft("DecompressPassword") & "' size='30' maxlength='30'></td>"
        Response.Write "          </tr>"
    End If
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "ͼƬ��</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='SoftPicUrl' type='text' id='SoftPicUrl' value='" & rsSoft("SoftPicUrl") & "' size='80' maxlength='200'>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'></td>"
    Response.Write "            <td><table><tr><td>�ϴ�" & ChannelShortName & "ͼƬ��</td><td><iframe style='top:2px' id='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=softpic' frameborder=0 scrolling=no width='450' height='25'></iframe></td></tr></table></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5' valign='middle'><strong>" & ChannelShortName & "��飺</strong></td>"
    Response.Write "            <td><textarea name='SoftIntro' cols='80' rows='10' id='SoftIntro' style='display:none'>" & Server.HTMLEncode(FilterJS(rsSoft("SoftIntro"))) & "</textarea>"
    Response.Write "                <iframe ID='editor' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=3&tContentid=SoftIntro' frameborder='1' scrolling='no' width='650' height='200' ></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>�ϴ�" & ChannelShortName & "��</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <iframe style='top:2px' ID='UploadFiles' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=soft' frameborder=0 scrolling=no width='450' height='25'></iframe>"
    Response.Write "            </td>"
    Response.Write "          </tr>"	
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "��ַ��</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "                <tr>"
    Response.Write "                  <td>"
    Response.Write "                    <input type='hidden' name='DownloadUrls' value=''>"
    Response.Write "                    <select name='DownloadUrl' style='width:400;height:100' size='2' ondblclick='return ModifyUrl();'>"
    Dim DownloadUrls, arrDownloadUrls, iTemp
    DownloadUrls = rsSoft("DownloadUrl")
    If InStr(DownloadUrls, "$$$") > 1 Then
        arrDownloadUrls = Split(DownloadUrls, "$$$")
        For iTemp = 0 To UBound(arrDownloadUrls)
            Response.Write "<option value='" & arrDownloadUrls(iTemp) & "'>" & arrDownloadUrls(iTemp) & "</option>"
        Next
    Else
        Response.Write "<option value='" & DownloadUrls & "'>" & DownloadUrls & "</option>"
    End If
    Response.Write "                    </select>"
    Response.Write "                  </td>"
    Response.Write "                  <td>"
    Response.Write "                    <input type='button' name='addurl' value='����ⲿ��ַ' onclick='AddUrl();'><br>"
    Response.Write "                    <input type='button' name='modifyurl' value='�޸ĵ�ǰ��ַ' onclick='return ModifyUrl();'><br>"
    Response.Write "                    <input type='button' name='delurl' value='ɾ����ǰ��ַ' onclick='DelUrl();'>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
        Response.Write "              <tr><td  colspan='3'>ϵͳ�ṩ���ϴ�����ֻ�ʺ��ϴ��Ƚ�С��" & ChannelShortName & "����ASPԴ����ѹ�����������" & ChannelShortName & "�Ƚϴ�" & MaxFileSize \ 1024 & "M���ϣ�������ʹ��FTP�ϴ�������Ҫʹ��ϵͳ�ṩ���ϴ����ܣ������ϴ���������ռ�÷�������CPU��Դ��FTP�ϴ����뽫��ַ���Ƶ�����ĵ�ַ���С�</td></tr>"		
    Response.Write "              </table>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'><strong>" & ChannelShortName & "��С��</strong></td>"
    Response.Write "            <td><input name='SoftSize' type='text' id='SoftSize' value='" & rsSoft("SoftSize") & "' size='10' maxlength='10'> K</td>"
    Response.Write "          </tr>"
    '�Զ����ֶ�
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & ChannelID & " or ChannelID=-2")
    Do While Not rsField.EOF
        If rsField("ShowOnForm") = True then
            Call WriteFieldHTML(rsField("FieldName"), rsField("Title"), rsField("Tips"), rsField("FieldType"), rsSoft(Trim(rsField("FieldName"))), rsField("Options"), rsField("EnableNull"))
        End If
        rsField.MoveNext
    Loop
    Set rsField = Nothing
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "״̬��</td>"
    Response.Write "            <td>"
    If rsSoft("Status") <= 0 Then
        Response.Write "<Input Name='Status' Type='radio' Id='Status' Value='-1'"
        If rsSoft("Status") = -1 Then
            Response.Write " checked"
        End If
        Response.Write "> �ݸ�&nbsp;&nbsp;"
        Response.Write "<Input Name='Status' Type='radio' Id='Status' Value='0'"
        If rsSoft("Status") = 0 Then
            Response.Write "checked"
        End If
        Response.Write "> Ͷ��"
    Else
        If rsSoft("Status") < 3 Then
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
    Response.Write "   <input name='SoftID' type='hidden' id='SoftID' value='" & rsSoft("SoftID") & "'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='Save' type='submit' value='�����޸Ľ��' style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
    rsSoft.Close
    Set rsSoft = Nothing

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
    Case 1 ,8   '�����ı���
        Response.Write "<input type='text' name='" & FieldName & "' size='80' maxlength='255' value='" & strValue & "'>" & strEnableNull
    Case 2 ,9   '�����ı���
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

Sub SaveSoft()
    If FoundInArr(arrClass_Input, ChannelDir & "none", ",") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�Բ�����û����" & ChannelName & "���" & ChannelShortName & "��Ȩ�ޣ�</li><br><br>"
        Exit Sub
    End If
    Dim rsSoft, sql
    Dim trs, tAuthor
    Dim SoftID, ClassID, SpecialID, SoftName, SoftVersion, SoftType, SoftLanguage, CopyrightType, OperatingSystem, Author, CopyFrom
    Dim DemoUrl, RegUrl, SoftPicUrl, SoftIntro, Keyword, DecompressPassword, SoftSize, DownloadUrls, Inputer
    Dim PresentExp, DefaultItemPoint, DefaultItemChargeType, DefaultItemPitchTime, DefaultItemReadTimes, DefaultItemDividePercent

    
    SoftID = PE_CLng(Trim(Request.Form("SoftID")))
    ClassID = PE_CLng(Trim(Request.Form("ClassID")))
    SpecialID = PE_CLng(Trim(Request.Form("SpecialID")))
    SoftName = Trim(Request.Form("SoftName"))
    SoftVersion = Trim(Request.Form("SoftVersion"))
    Keyword = Trim(Request.Form("Keyword"))
    SoftType = PE_HTMLEncode(Trim(Request.Form("SoftType")))
    SoftLanguage = PE_HTMLEncode(Trim(Request.Form("SoftLanguage")))
    CopyrightType = PE_HTMLEncode(Trim(Request.Form("CopyrightType")))
    OperatingSystem = PE_HTMLEncode(Trim(Request.Form("OperatingSystem")))
    Author = PE_HTMLEncode(Trim(Request.Form("Author")))
    CopyFrom = PE_HTMLEncode(Trim(Request.Form("CopyFrom")))
    DemoUrl = PE_HTMLEncode(Trim(Request.Form("DemoUrl")))
    RegUrl = PE_HTMLEncode(Trim(Request.Form("RegUrl")))
    SoftPicUrl = PE_HTMLEncode(Trim(Request.Form("SoftPicUrl")))
    SoftIntro = ReplaceBadUrl(FilterJS(Trim(Request.Form("SoftIntro"))))
    DecompressPassword = PE_HTMLEncode(Trim(Request.Form("DecompressPassword")))
    SoftSize = PE_CLng(Trim(Request.Form("SoftSize")))
    DownloadUrls = PE_HTMLEncode(Trim(Request.Form("DownloadUrls")))
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
    
    If SoftName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "���Ʋ���Ϊ��</li>"
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
    If FoundInArr(arrEnabledTabs, "SoftParameter", ",") = True Then
        If SoftType = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & ChannelShortName & "�����Ϊ��</li>"
        End If
        If SoftLanguage = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & ChannelShortName & "���Բ���Ϊ��</li>"
        End If
        If CopyrightType = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��Ȩ��ʽ����Ϊ��</li>"
        End If
        If OperatingSystem = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>" & ChannelShortName & "ƽ̨����Ϊ��</li>"
        End If
    End If
    If DownloadUrls = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>" & ChannelShortName & "���ص�ַ����Ϊ��</li>"
    End If

    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & ChannelID & " or ChannelID=-2")
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
    
    SoftName = PE_HTMLEncode(SoftName)
    SoftVersion = PE_HTMLEncode(SoftVersion)
    SoftType = PE_HTMLEncode(SoftType)
    SoftLanguage = PE_HTMLEncode(SoftLanguage)
    CopyrightType = PE_HTMLEncode(CopyrightType)
    OperatingSystem = PE_HTMLEncode(OperatingSystem)
    DemoUrl = PE_HTMLEncode(DemoUrl)
    RegUrl = PE_HTMLEncode(RegUrl)
    SoftPicUrl = PE_HTMLEncode(SoftPicUrl)
    Keyword = "|" & ReplaceBadChar(Keyword) & "|"
    DecompressPassword = PE_HTMLEncode(DecompressPassword)

    Set rsSoft = Server.CreateObject("adodb.recordset")
    If Action = "SaveAdd" Then
        If Session("SoftName") = SoftName And DateDiff("S", Session("AddTime"), Now()) < 100 Then
            FoundErr = True
            ErrMsg = "<li>�벻Ҫ�ظ����ͬһ" & ChannelItemUnit & ChannelShortName & "</li>"
            Exit Sub
        Else
            Session("SoftName") = SoftName
            Session("AddTime") = Now()
            If MaxPerDay > 0 Then
                Set trs = Conn.Execute("select count(SoftID) from PE_Soft where Inputer='" & UserName & "' and DateDiff(" & PE_DatePart_D & ",UpdateTime," & PE_Now & ")=0")
                If trs(0) >= MaxPerDay Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�����췢���" & ChannelShortName & "�Ѿ��ﵽ�����ޣ�</li>"
                End If
                Set trs = Nothing
                If FoundErr = True Then Exit Sub
            End If
            
            sql = "select top 1 * from PE_Soft"
            rsSoft.Open sql, Conn, 1, 3
            rsSoft.addnew
            SoftID = PE_CLng(Conn.Execute("select max(SoftID) from PE_Soft")(0)) + 1
            Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (2," & SoftID & "," & SpecialID & ")")
            rsSoft("SoftID") = SoftID
            rsSoft("ChannelID") = ChannelID
            rsSoft("ClassID") = ClassID
            rsSoft("SoftName") = SoftName
            rsSoft("SoftVersion") = SoftVersion
            rsSoft("SoftType") = SoftType
            rsSoft("SoftLanguage") = SoftLanguage
            rsSoft("CopyrightType") = CopyrightType
            rsSoft("OperatingSystem") = OperatingSystem
            rsSoft("Author") = Author
            rsSoft("CopyFrom") = CopyFrom
            rsSoft("DemoUrl") = DemoUrl
            rsSoft("RegUrl") = RegUrl
            rsSoft("SoftPicUrl") = SoftPicUrl
            rsSoft("SoftIntro") = SoftIntro
            rsSoft("Keyword") = Keyword
            rsSoft("Hits") = 0
            rsSoft("DayHits") = 0
            rsSoft("WeekHits") = 0
            rsSoft("MonthHits") = 0
            rsSoft("Stars") = 0
            rsSoft("UpdateTime") = Now()
            rsSoft("Status") = Status
            rsSoft("OnTop") = False
            rsSoft("Elite") = False
            rsSoft("DecompressPassword") = DecompressPassword
            rsSoft("SoftSize") = SoftSize
            rsSoft("DownloadUrl") = DownloadUrls
            rsSoft("Inputer") = Inputer
            rsSoft("Editor") = Inputer
            rsSoft("SkinID") = 0
            rsSoft("TemplateID") = 0
            rsSoft("Deleted") = False
            PresentExp = CLng(PresentExp * PresentExpTimes)
            rsSoft("PresentExp") = PresentExp
            rsSoft("InfoPoint") = DefaultItemPoint
            rsSoft("VoteID") = 0
            rsSoft("InfoPurview") = 0
            rsSoft("arrGroupID") = ""
            rsSoft("ChargeType") = DefaultItemChargeType
            rsSoft("PitchTime") = DefaultItemPitchTime
            rsSoft("ReadTimes") = DefaultItemReadTimes
            rsSoft("DividePercent") = DefaultItemDividePercent
            
            If Not (rsField.BOF And rsField.EOF) Then
                rsField.MoveFirst
                Do While Not rsField.EOF
                    If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                        rsSoft(Trim(rsField("FieldName"))) = FilterJS(Trim(Request(rsField("FieldName"))))
                    End If
                    rsField.MoveNext
                Loop
            End If
            Set rsField = Nothing

            If BlogFlag = True Then 'д��BLOGID
                Dim blogid
                Set blogid = Conn.Execute("select top 1 ID from PE_Space where Type=1 and UserID=" & UserID)
                If blogid.BOF And blogid.EOF Then
                    rsSoft("BlogID") = 0
                Else
                    rsSoft("BlogID") = blogid("ID")
                End If
                Set blogid = Nothing
            End If

            rsSoft.Update
            If CheckLevel = 0 Or NeedlessCheck = 1 Then
                Conn.Execute ("update PE_Channel set ItemCount=ItemCount+1,ItemChecked=ItemChecked+1 where ChannelID=" & ChannelID & "")
                Conn.Execute ("update PE_Class set ItemCount=ItemCount+1 where ClassID=" & ClassID & "")
                If rsSoft("Status") = 3 Then
                    Conn.Execute ("update PE_User set PostItems=PostItems+1,PassedItems=PassedItems+1,UserExp=UserExp+" & PresentExp & " where UserName='" & UserName & "'")
                End If
            Else
                Conn.Execute ("update PE_Channel set ItemCount=ItemCount+1 where ChannelID=" & ChannelID & "")
                Conn.Execute ("update PE_User set PostItems=PostItems+1 where UserName='" & UserName & "'")
            End If
        End If
    ElseIf Action = "SaveModify" Then
        If SoftID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����ȷ��" & ChannelShortName & "ID��ֵ</li>"
        Else
            sql = "select * from PE_Soft where Inputer='" & UserName & "' and Deleted=" & PE_False & " and SoftID=" & SoftID
            rsSoft.Open sql, Conn, 1, 3
            If rsSoft.BOF And rsSoft.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�Ҳ�����" & ChannelShortName & "�������Ѿ���������ɾ����</li>"
            Else
                If rsSoft("Status") > 0 And EnableModifyDelete = 0 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>" & ChannelShortName & "�Ѿ������ͨ�����������ٽ����޸ģ�</li>"
                Else
                    Conn.Execute ("delete from PE_InfoS where ModuleType=2 and ItemID=" & SoftID)
                    Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (2," & SoftID & "," & SpecialID & ")")
                    rsSoft("ClassID") = ClassID
                    rsSoft("SoftName") = SoftName
                    rsSoft("SoftVersion") = SoftVersion
                    rsSoft("SoftType") = SoftType
                    rsSoft("SoftLanguage") = SoftLanguage
                    rsSoft("CopyrightType") = CopyrightType
                    rsSoft("OperatingSystem") = OperatingSystem
                    rsSoft("Author") = Author
                    rsSoft("CopyFrom") = CopyFrom
                    rsSoft("DemoUrl") = DemoUrl
                    rsSoft("RegUrl") = RegUrl
                    rsSoft("SoftPicUrl") = SoftPicUrl
                    rsSoft("SoftIntro") = SoftIntro
                    rsSoft("Keyword") = Keyword
                    rsSoft("UpdateTime") = Now()
                    rsSoft("DecompressPassword") = DecompressPassword
                    rsSoft("SoftSize") = SoftSize
                    rsSoft("DownloadUrl") = DownloadUrls
                    rsSoft("Status") = Status
                    
                    If Not (rsField.BOF And rsField.EOF) Then
                        rsField.MoveFirst
                        Do While Not rsField.EOF
                            If Trim(Request(rsField("FieldName"))) <> "" Or rsField("EnableNull") = True Then
                                rsSoft(Trim(rsField("FieldName"))) = PE_HTMLEncode(FilterJS(Trim(Request(rsField("FieldName")))))
                            End If
                            rsField.MoveNext
                        Loop
                    End If
                    Set rsField = Nothing

                    rsSoft.Update
                End If
            End If
        End If
    End If
    rsSoft.Close
    Set rsSoft = Nothing
    
    If FoundErr = True Then Exit Sub
    
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
    Response.Write "          <td width='100' align='right'><strong>" & ChannelShortName & "���ƣ�</strong></td>"
    Response.Write "          <td>" & SoftName & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right'><strong>" & ChannelShortName & "�汾��</strong></td>"
    Response.Write "          <td>" & SoftVersion & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='100' align='right'><strong>" & ChannelShortName & "���ߣ�</strong></td>"
    Response.Write "          <td>" & Author & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'> "
    Response.Write "          <td width='100' align='right'><strong>�� �� �֣�</strong></td>"
    Response.Write "          <td>" & Mid(Keyword, 2, Len(Keyword) - 2) & "</td>"
    Response.Write "        </tr>"
    Response.Write "      </table></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='30' align='center'>"
    Response.Write "��<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Modify&SoftID=" & SoftID & "'>�޸Ĵ�" & ChannelShortName & "</a>��&nbsp;"
    Response.Write "��<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "'>�������" & ChannelShortName & "</a>��&nbsp;"
    Response.Write "��<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Manage&ClassID=" & ClassID & "'>" & ChannelShortName & "����</a>��&nbsp;"
    Response.Write "��<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Show&SoftID=" & SoftID & "'>Ԥ��" & ChannelShortName & "����</a>��"
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
    If SoftID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����ѡ��" & ChannelShortName & "��</li>"
        Exit Sub
    End If

    Dim sqlDel, rsDel, NeedUpdateCache
    NeedUpdateCache = False

    sqlDel = "select * from PE_Soft where Inputer='" & UserName & "' and Deleted=" & PE_False & " and "
    If InStr(SoftID, ",") > 0 Then
        sqlDel = sqlDel & " SoftID in (" & SoftID & ") order by SoftID"
    Else
        sqlDel = sqlDel & " SoftID=" & SoftID
    End If
    Set rsDel = Server.CreateObject("ADODB.Recordset")
    rsDel.Open sqlDel, Conn, 1, 3
    Do While Not rsDel.EOF
        If rsDel("Status") > 0 Then
            If EnableModifyDelete = 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>ɾ��" & ChannelShortName & "��" & rsDel("SoftName") & "��ʧ�ܡ�ԭ�򣺴�" & ChannelShortName & "�Ѿ������ͨ������������ɾ����</li>"
            Else
                Conn.Execute ("update PE_User set PostItems=PostItems-1,PassedItems=PassedItems-1,UserExp=UserExp-" & rsDel("PresentExp") & " where UserName='" & UserName & "'")
                Conn.Execute ("update PE_Channel set ItemCount=ItemCount-1,ItemChecked=ItemChecked-1 where ChannelID=" & ChannelID & "")
                Conn.Execute ("update PE_Class set ItemCount=ItemCount-1 where ClassID=" & rsDel("ClassID") & "")
                rsDel("Deleted") = True
                rsDel.Update
                NeedUpdateCache = True
            End If
        Else
            Conn.Execute ("update PE_Channel set ItemCount=ItemCount-1 where ChannelID=" & ChannelID & "")
            Conn.Execute ("update PE_User set PostItems=PostItems-1 where UserName='" & UserName & "'")
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

Sub Show()
    Dim rs, sql
    SoftID = PE_CLng(SoftID)
    sql = "select * from PE_Soft where Inputer='" & UserName & "' and Deleted=" & PE_False & " and SoftID=" & SoftID & ""
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, Conn, 1, 1
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���" & ChannelShortName & "</li>"
    Else
        ClassID = rs("ClassID")
        Call GetClass
    End If
    If FoundErr = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "<tr class='title'>"
    Response.Write "  <td height='22' colspan='4'>"
    Response.Write "�����ڵ�λ�ã�&nbsp;<a href='User_Soft.asp?ChannelID=" & ChannelID & "'>" & ChannelShortName & "����</a>&nbsp;&gt;&gt;&nbsp;"
    If ParentID > 0 Then
        Dim sqlPath, rsPath
        sqlPath = "select ClassID,ClassName from PE_Class where ClassID in (" & ParentPath & ") order by Depth"
        Set rsPath = Server.CreateObject("adodb.recordset")
        rsPath.Open sqlPath, Conn, 1, 1
        Do While Not rsPath.EOF
            Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
    Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Show&SoftID=" & rs("SoftID") & "'>" & PE_HTMLEncode(rs("SoftName")) & "</a>"
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "���ƣ�</td>"
    Response.Write "  <td colspan='3'><strong>" & PE_HTMLEncode(rs("SoftName")) & "&nbsp;&nbsp;" & PE_HTMLEncode(rs("SoftVersion")) & "</strong></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>�ļ���С��</td>"
    Response.Write "  <td width='200'>" & rs("SoftSize") & " K" & "</td>"
    Response.Write "  <td colspan='2' rowspan='7' align=center valign='middle'>"
    If rs("SoftPicUrl") = "" Then
        Response.Write "���ͼƬ"
    Else
        Response.Write "<img src='" & GetSoftPicUrl(rs("SoftPicUrl")) & "' width='150'>"
    End If
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>���л�����</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(rs("OperatingSystem")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "���</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(rs("SoftType")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "���ԣ�</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(rs("SoftLanguage")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>��Ȩ��ʽ��</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(rs("CopyrightType")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>���ֵȼ���</td>"
    Response.Write "  <td width='200'>" & String(rs("Stars"), "��") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>��ѹ���룺</td>"
    Response.Write "  <td width='200'>" & rs("DecompressPassword") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>���ʱ�䣺</td>"
    Response.Write "  <td width='200'>" & rs("UpdateTime") & "</td>"
    Response.Write "  <td width='100' align=right valign='middle'>�� �� �̣�</td>"
    Response.Write "  <td valign='middle'>" & PE_HTMLEncode(rs("Author")) & ""
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align=right valign='middle'>���ص�����</td>"
    Response.Write "  <td width='200'><font color=red> " & rs("InfoPoint") & "</font> ��</td>"
    Response.Write "  <td></td>"
    Response.Write "  <td></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "��ӣ�</td>"
    Response.Write "  <td width='200'>" & rs("Inputer") & "</td>"
    Response.Write "  <td width='100' align=right valign='middle'>���α༭��</td>"
    Response.Write "  <td valign='middle'>"
    If rs("Status") > 0 Then
        Response.Write rs("Editor")
    Else
        Response.Write "��"
    End If
    Response.Write "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>������ӣ�</td>"
    Response.Write "  <td colspan='3'><a href='" & rs("DemoUrl") & "' target='_blank'>" & ChannelShortName & "��ʾ��ַ</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='" & rs("RegUrl") & "' target='_blank'>" & ChannelShortName & "ע���ַ</a></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>���ش�����</td>"
    Response.Write "  <td colspan='3'>���գ�" & rs("DayHits") & "&nbsp;&nbsp;&nbsp;&nbsp;���ܣ�" & rs("WeekHits") & "&nbsp;&nbsp;&nbsp;&nbsp;���£�" & rs("MonthHits") & "&nbsp;&nbsp;&nbsp;&nbsp;�ܼƣ�" & rs("Hits")
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>���ص�ַ��</td>"
    Response.Write "  <td colspan='3'>" & ShowDownloadUrls(rs("DownloadUrl")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td align='right'>&nbsp;</td>"
    Response.Write "  <td colspan='3' align='right'>"
    Response.Write "<strong>���ò�����</strong>"
    If rs("Inputer") = UserName And (rs("Status") <= 0 Or UserSetting(2) = 1) Then
        Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Modify&SoftID=" & rs("SoftID") & "'>�޸�</a>&nbsp;&nbsp;"
        Response.Write "<a href='User_Soft.asp?ChannelID=" & ChannelID & "&Action=Del&SoftID=" & rs("SoftID") & "' onclick=""return confirm('ȷ��Ҫɾ����" & ChannelShortName & "��');"">ɾ��</a>&nbsp;&nbsp;"
    End If
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "��飺</td>"
    Response.Write "  <td height='100' colspan='3'>" & FilterJS(rs("SoftIntro")) & "</td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    rs.Close
    Set rs = Nothing
End Sub

Sub Preview()
    Response.Write "<br><table width='100%' border=0 align=center cellPadding=2 cellSpacing=1 bgcolor='#FFFFFF' class='border' style='WORD-BREAK: break-all'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='4'>"

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
            Response.Write rsPath(1) & "&nbsp;&gt;&gt;&nbsp;"
            rsPath.MoveNext
        Loop
        rsPath.Close
        Set rsPath = Nothing
    End If
    Response.Write ClassName & "&nbsp;&gt;&gt;&nbsp;"

    Response.Write PE_HTMLEncode(Request("SoftName"))
    Response.Write " </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "���ƣ�</td>"
    Response.Write "  <td colspan='3'><strong>" & PE_HTMLEncode(Request("SoftName")) & "&nbsp;&nbsp;" & PE_HTMLEncode(Request("SoftVersion")) & "</strong></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>�ļ���С��</td>"
    Response.Write "  <td width='200'>" & Request("SoftSize") & " K" & "</td>"
    Response.Write "  <td colspan='2' rowspan='7' align=center valign='middle'>"
    If Request("SoftPicUrl") = "" Then
        Response.Write "���ͼƬ"
    Else
        Response.Write "<img src='" & GetSoftPicUrl(Request("SoftPicUrl")) & "' width='150'>"
    End If
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>���л�����</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(Request("OperatingSystem")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "���</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(Request("SoftType")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "���ԣ�</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(Request("SoftLanguage")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>��Ȩ��ʽ��</td>"
    Response.Write "  <td width='200'>" & PE_HTMLEncode(Request("CopyrightType")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>���ֵȼ���</td>"
    Response.Write "  <td width='200'>" & String(Request("Stars"), "��") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>��ѹ���룺</td>"
    Response.Write "  <td width='200'>" & Request("DecompressPassword") & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>���ʱ�䣺</td>"
    Response.Write "  <td width='200'>" & Now() & "</td>"
    Response.Write "  <td width='100' align=right valign='middle'>�� �� �̣�</td>"
    Response.Write "  <td valign='middle'>" & PE_HTMLEncode(Request("Author")) & ""
    Response.Write "  </td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>������ӣ�</td>"
    Response.Write "  <td colspan='3'><a href='" & Request("DemoUrl") & "' target='_blank'>" & ChannelShortName & "��ʾ��ַ</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='" & Request("RegUrl") & "' target='_blank'>" & ChannelShortName & "ע���ַ</a></td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "��ַ��</td>"
    Response.Write "  <td colspan='3'>" & ShowDownloadUrls(Request("DownloadUrls")) & "</td>"
    Response.Write "</tr>"
    Response.Write "<tr class='tdbg'>"
    Response.Write "  <td width='100' align='right'>" & ChannelShortName & "��飺</td>"
    Response.Write "  <td height='100' colspan='3'>" & FilterJS(Request("SoftIntro")) & "</td>"
    Response.Write "</tr>"
    Response.Write "</table>"
    Response.Write "<p align='center'>��<a href='javascript:window.close();'>�رմ���</a>��</p>"
End Sub

Function GetSoftPicUrl(SoftPicUrl)
    If Left(SoftPicUrl, Len("UploadSoftPic")) = "UploadSoftPic" Then
        GetSoftPicUrl = strInstallDir & ChannelDir & "/" & SoftPicUrl
    Else
        GetSoftPicUrl = SoftPicUrl
    End If
End Function

Function ShowDownloadUrls(DownloadUrls)
    Dim arrDownloadUrls, arrUrls, iTemp, strUrls
    strUrls = ""
    arrDownloadUrls = Split(DownloadUrls, "$$$")
    For iTemp = 0 To UBound(arrDownloadUrls)
        arrUrls = Split(arrDownloadUrls(iTemp), "|")
        If UBound(arrUrls) = 1 Then
            If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                strUrls = strUrls & arrUrls(0) & "��<a href='" & strInstallDir & ChannelDir & "/" & UploadDir & "/" & arrUrls(1) & "'>" & arrUrls(1) & "</a><br>"
            Else
                strUrls = strUrls & arrUrls(0) & "��<a href='" & arrUrls(1) & "'>" & arrUrls(1) & "</a><br>"
            End If
        End If
    Next
    ShowDownloadUrls = strUrls
End Function

Function GetSoftType(SoftType)
    If IsArray(arrSoftType) = False Then Exit Function
    
    Dim strTemp, i
    For i = 0 To UBound(arrSoftType)
        If Trim(arrSoftType(i)) <> "" Then
            strTemp = strTemp & "<option value='" & arrSoftType(i) & "'"
            If Trim(SoftType) = arrSoftType(i) Then strTemp = strTemp & " selected"
            strTemp = strTemp & ">" & arrSoftType(i) & "</option>"
        End If
    Next
    GetSoftType = strTemp
End Function

Function GetSoftLanguage(SoftLanguage)
    If IsArray(arrSoftLanguage) = False Then Exit Function
    
    Dim strTemp, i
    For i = 0 To UBound(arrSoftLanguage)
        If Trim(arrSoftLanguage(i)) <> "" Then
            strTemp = strTemp & "<option value='" & arrSoftLanguage(i) & "'"
            If Trim(SoftLanguage) = arrSoftLanguage(i) Then strTemp = strTemp & " selected"
            strTemp = strTemp & ">" & arrSoftLanguage(i) & "</option>"
        End If
    Next
    GetSoftLanguage = strTemp
End Function

Function GetCopyrightType(CopyrightType)
    If IsArray(arrCopyrightType) = False Then Exit Function
    
    Dim strTemp, i
    For i = 0 To UBound(arrCopyrightType)
        If Trim(arrCopyrightType(i)) <> "" Then
            strTemp = strTemp & "<option value='" & arrCopyrightType(i) & "'"
            If Trim(CopyrightType) = arrCopyrightType(i) Then strTemp = strTemp & " selected"
            strTemp = strTemp & ">" & arrCopyrightType(i) & "</option>"
        End If
    Next
    GetCopyrightType = strTemp
End Function

Function GetOperatingSystemList()
    Dim strOperatingSystemList, i
    
    strOperatingSystemList = "<script language = 'JavaScript'>" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "function ToSystem(addTitle){" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "    var str=document.myform.OperatingSystem.value;" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "    if (document.myform.OperatingSystem.value=="""") {" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "        document.myform.OperatingSystem.value=document.myform.OperatingSystem.value+addTitle;" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "    }else{" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "        if (str.substr(str.length-1,1)==""" & XmlText("Soft", "OperatingSystemEmblem", "/") & """){" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "            document.myform.OperatingSystem.value=document.myform.OperatingSystem.value+addTitle;" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "        }else{" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "            document.myform.OperatingSystem.value=document.myform.OperatingSystem.value+""" & XmlText("Soft", "OperatingSystemEmblem", "/") & """+addTitle;" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "        }" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "    }" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "    document.myform.OperatingSystem.focus();" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "}" & vbCrLf
    strOperatingSystemList = strOperatingSystemList & "</script>" & vbCrLf

    strOperatingSystemList = strOperatingSystemList & "<font color='#808080'>ƽ̨ѡ��"
    If IsArray(arrOperatingSystem) Then
        For i = 0 To UBound(arrOperatingSystem)
            If Trim(arrOperatingSystem(i)) <> "" Then
                strOperatingSystemList = strOperatingSystemList & "<a href=""javascript:ToSystem('" & arrOperatingSystem(i) & "')"">" & arrOperatingSystem(i) & "</a>" & XmlText("Soft", "OperatingSystemEmblem", "/")
            End If
        Next
    End If
    strOperatingSystemList = strOperatingSystemList & "</font>"
    GetOperatingSystemList = strOperatingSystemList
End Function
%>
