<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "Collection"   '����Ȩ��

Private rs, sql, rsItem, strsql, i 'ͨ�ñ���
Private SelectCollateItemID


strFileName = "Admin_CollectionHistory.asp?Action=" & Action
SelectCollateItemID = PE_CLng(Trim(Request("SelectCollateItemID")))

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title> �� ʷ �� ¼ �� �� </title>" & vbCrLf
Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""Admin_Style.css"">" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">" & vbCrLf
Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"">" & vbCrLf
Call ShowPageTitle(" �� �� �� ʷ �� ¼ �� �� ", 10054)
Response.Write "  <tr class='tdbg'> " & vbCrLf
Response.Write "    <td height='30'>&nbsp;&nbsp;���ٲ��ң�<select name='SelectCollateItemID' onchange=""javascript:window.location='Admin_CollectionHistory.asp?Action=History&SelectCollateItemID='+this.value;"" > " & vbCrLf

sql = "SELECT ItemID,ItemName FROM PE_Item"
Set rsItem = Conn.Execute(sql)
If rsItem.BOF And rsItem.EOF Then
    Response.Write "<option value='0' selected>��û�вɼ���Ŀ��</option> "
Else
    Response.Write "<option value='0'"
    If SelectCollateItemID = 0 Then Response.Write " selected"
    Response.Write ">������Ŀ��ʷ��¼</option>"
    Response.Write "<option value='-1' "
    If SelectCollateItemID = -1 Then Response.Write " selected"
    Response.Write ">������Ŀʧ�ܼ�¼</option>"
    Do While Not rsItem.EOF
        Response.Write "<option value=" & rsItem("ItemID") & " "
        If SelectCollateItemID = rsItem("ItemID") Then Response.Write "selected"
        Response.Write ">" & rsItem("ItemName") & "</option>"
        rsItem.MoveNext
    Loop
End If
rsItem.Close
Set rsItem = Nothing
Response.Write "</select></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case Action
    Case "main"
        Call main
        'Ԥ����չ����
    Case "Del"
        Call Del
    Case "DelFaild"
        Call DelFaild
    Case "DelItem"
        Call DelItem
    Case Else
        Call main
End Select
Response.Write "</body></html>"
Call CloseConn


'=================================================
'��������main
'��  �ã��ɼ���ʷ��¼����
'=================================================
Sub main()
    Dim ItemID, TitleRight, strFileName, HistrolyNewsID, ArticleID, NewsCollecDate, Result, Title
    Dim NewsUrl, rsHistroly, HistrolyResult
    Dim ClassID
    Dim MaxPerPage
    
    MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))
    If MaxPerPage <= 0 Then MaxPerPage = 20
    If Trim(Request("HistrolyResult")) = "false" Then
        HistrolyResult = PE_False
    Else
        HistrolyResult = PE_True
    End If

    If Request("page") <> "" Then
        CurrentPage = CInt(Request("page"))
    Else
        CurrentPage = 1
    End If
    strFileName = "Admin_CollectionHistory.asp?Action=History&SelectCollateItemID=" & SelectCollateItemID & "&HistrolyResult=" & Trim(Request("HistrolyResult"))

    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "function unselectall(thisform)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if(thisform.chkAll.checked)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        thisform.chkAll.checked = thisform.chkAll.checked&0;" & vbCrLf
    Response.Write "    }   " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckAll(thisform)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    for (var i=0;i<thisform.elements.length;i++)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "    var e = thisform.elements[i];" & vbCrLf
    Response.Write "    if (e.Name != ""chkAll""&&e.disabled!=true)" & vbCrLf
    Response.Write "        e.checked = thisform.chkAll.checked;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function ConfirmDel(thisform)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if(thisform.Action.value==""Del"")" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        if(confirm(""ȷ��Ҫɾ��ѡ�еļ�¼��""))" & vbCrLf
    Response.Write "            return true;" & vbCrLf
    Response.Write "        else" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<form name=""form1"" method=""POST"" action=""Admin_CollectionHistory.asp"" onsubmit=""return confirm('ȷ��Ҫɾ��ѡ���ļ�¼��');"">" & vbCrLf
    Response.Write "   <table class=""border"" border=""0"" cellspacing=""1"" width=""100%"" cellpadding=""0"">" & vbCrLf
    Response.Write "     <tr class=""title"" align=""center"">" & vbCrLf
    Response.Write "      <td width=""30"" height=""22""><strong>ѡ��</strong></td>        " & vbCrLf
    Response.Write "      <td width=""120""><strong>��Ŀ����</strong></td>" & vbCrLf
    Response.Write "      <td><strong>���ű���</strong></td>" & vbCrLf
    Response.Write "      <td width=""100"" height=""22""><strong>����Ƶ��</strong></td>" & vbCrLf
    Response.Write "      <td width=""100"" height=""22""><strong>������Ŀ</strong></td>     " & vbCrLf
    Response.Write "      <td width=""60""><strong>�ɼ�ҳ��</strong></td>        " & vbCrLf
    Response.Write "      <td width=""40""><strong>��    ��</strong></td>" & vbCrLf
    Response.Write "      <td width=""40"" height=""22""><strong>����</strong></td>" & vbCrLf
    Response.Write "     </tr>   " & vbCrLf

    sql = "SELECT H.*, C.ChannelName, CL.ClassName,I.ItemName"
    sql = sql & " FROM ((PE_HistrolyNews H left JOIN PE_Channel C ON H.ChannelID =C.ChannelID)"
    sql = sql & " left JOIN PE_Item I ON H.ItemID =I.ItemID) left JOIN PE_Class CL ON H.ClassID = CL.ClassID"
    If SelectCollateItemID = -1 Then
        sql = sql & " Where H.Result=" & PE_False
    ElseIf SelectCollateItemID > 0 Then
        sql = sql & " where H.ItemID=" & SelectCollateItemID & " and H.Result=" & HistrolyResult
    End If
    sql = sql & " ORDER BY H.NewsCollecDate DESC"
    
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 1
    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg'><td colspan='20' height='50' align='center'>����ʷ��¼��</td></tr></table>"
    Else
        totalPut = rs.RecordCount
        TitleRight = TitleRight & "�� <font color=red>" & totalPut & "</font> ����ʷ��¼"
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
                rs.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        
        Dim VisitorNum
        VisitorNum = 0

      
        Do While Not rs.EOF
            HistrolyNewsID = rs("HistrolyNewsID")
            ArticleID = rs("ArticleID")
            Title = rs("Title")
            NewsCollecDate = rs("NewsCollecDate")
            NewsUrl = rs("NewsUrl")
            Result = rs("Result")
            ItemID = rs("ItemID")
            ClassID = rs("ClassID")
            ChannelID = rs("ChannelID")
            
            Response.Write "<tr class=""tdbg"" onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"" style=""padding: 0px 2px;"">" & vbCrLf
            Response.Write "  <td width=""30"" align=""center"">" & vbCrLf
            Response.Write "    <input type=""checkbox"" value= " & HistrolyNewsID & " name=""HistrolyNewsID"" onclick=""unselectall(this.form)"" >" & vbCrLf
            Response.Write "  </td>" & vbCrLf
            Response.Write "  <td width=""120"" align=""center"">" & rs("ItemName") & "</td>" & vbCrLf
            Response.Write "  <td>"
            If Title = "" Or IsNull(Title) Then
                Response.Write "����" & vbCrLf
            Else
                Response.Write "      <a href=Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & ArticleID & " target=_bank title='�ɼ�ʱ�䣺" & NewsCollecDate & "'>" & Title & "</a>" & vbCrLf
            End If
            Response.Write "  </td>" & vbCrLf
            Response.Write "  <td width=""100"" align=""center"">" & rs("ChannelName") & "</td>" & vbCrLf
            Response.Write "  <td width=""100"" align=""center"">" & rs("ClassName") & "</td>" & vbCrLf
            Response.Write "  <td width=""60"" align=""center""><a href=" & NewsUrl & " target=_blank title=" & NewsUrl & ">������</a></td>" & vbCrLf
            Response.Write "  <td width=""40"" align=""center"">" & vbCrLf
            If Result = True Then
                Response.Write "�ɹ�" & vbCrLf
            ElseIf Result = False Then
                Response.Write "<font color=red>ʧ��</font>" & vbCrLf
            Else
                Response.Write "<font color=red>�쳣</font>" & vbCrLf
            End If
            Response.Write "  </td>" & vbCrLf
            Response.Write "  <td width=""40"" align=""center"">" & vbCrLf
            Response.Write "    <a href=Admin_CollectionHistory.asp?Action=Del&HistrolyNewsID=" & HistrolyNewsID & " onclick='return confirm(""ȷ��Ҫɾ���˼�¼��"");'>ɾ��</a>" & vbCrLf
            Response.Write "  </td>" & vbCrLf
            Response.Write "</tr>" & vbCrLf

            VisitorNum = VisitorNum + 1
            If VisitorNum >= MaxPerPage Then Exit Do
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        Response.Write "</table>" & vbCrLf
        Response.Write "<table border=""0"" cellspacing=""1"" width=""100%"" cellpadding=""0""><tr><td height=""30"">" & vbCrLf
        Response.Write "  <input name=""Action"" type=""hidden""  value=""History"">" & vbCrLf
        Response.Write "  <input name=""chkAll"" type=""checkbox"" id=""chkAll"" onclick=CheckAll(this.form) value=""checkbox"" >ѡ��������Ŀ" & vbCrLf
        Response.Write "  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
        Response.Write "  <input type=""submit"" value="" ����ɾ�� "" name=""submit"" onClick=""document.form1.Action.value='Del';return ConfirmDel(this.form);"" >&nbsp;&nbsp;" & vbCrLf
        Response.Write "  <input type=""submit"" value=""���ʧ�ܼ�¼"" name=""DelDefeat"" onClick=""document.form1.Action.value='DelFaild';return ConfirmDel(this.form);"" >&nbsp;&nbsp;" & vbCrLf
        sql = "SELECT DISTINCT H.ItemID, I.ItemName FROM PE_HistrolyNews H INNER JOIN PE_Item I ON H.ItemID = I.ItemID"
        Set rs = Conn.Execute(sql)
        Response.Write "<select name='SelectHistoryItemID'>"
        If rs.BOF And rs.EOF Then
            Response.Write "<option value="""" selected>��û�вɼ���Ŀ��</option> "
        Else
            Do While Not rs.EOF
                Response.Write "<option value=" & rs("ItemID") & ">" & rs("ItemName") & "</option>"
                rs.MoveNext
            Loop
        End If
        Response.Write "</select>"
        rs.Close
        Set rs = Nothing
        Response.Write "        <input type=""submit"" value=""ɾ��ѡ����Ŀ����ʷ��¼"" name=""DelItem"" onClick=""document.form1.Action.value='DelItem';return ConfirmDel(this.form);"" >&nbsp;&nbsp;" & vbCrLf
        Response.Write "      </td>" & vbCrLf
        Response.Write "    </tr> " & vbCrLf
        Response.Write "</table>" & vbCrLf
        Response.Write "</form>" & vbCrLf
        If totalPut > 0 Then
            Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "����ʷ��¼", True)
        End If
    End If
End Sub

Sub Del()
    Dim HistrolyNewsID
    HistrolyNewsID = Trim(Request("HistrolyNewsID"))
	If IsValidID(HistrolyNewsID) = False Then
		HistrolyNewsID = ""
	End If
    If HistrolyNewsID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ѡ��Ҫɾ���ļ�¼��</li>"
    Else
        If InStr(HistrolyNewsID, ",") > 0 Then
            Conn.Execute ("Delete From [PE_HistrolyNews] Where HistrolyNewsID In(" & HistrolyNewsID & ")")
        Else
            Conn.Execute ("Delete From [PE_HistrolyNews] Where HistrolyNewsID=" & HistrolyNewsID)
        End If
		Call WriteSuccessMsg("<li>�Ѿ��ɹ�ɾ��ָ������ʷ��¼!", "Admin_CollectionHistory.asp?Action=main")
    End If
End Sub

Sub DelFaild()
    Conn.Execute ("Delete From PE_HistrolyNews Where Result=" & PE_False & "")
    Call WriteSuccessMsg("<li>�Ѿ��ɹ�ɾ�������вɼ�ʧ�ܵ���ʷ��¼!", "Admin_CollectionHistory.asp?Action=main")
End Sub

Sub DelItem()
    Conn.Execute ("Delete From PE_HistrolyNews Where ItemID=" & CLng(Trim(Request("SelectHistoryItemID"))) & "")
    Call WriteSuccessMsg("<li>�Ѿ��ɹ�ɾ����ָ���Ĳɼ���Ŀ��ʷ��¼!", "Admin_CollectionHistory.asp?Action=main")
End Sub

%>
