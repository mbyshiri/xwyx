<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_Content.asp"-->
<!--#include file="Admin_CommonCode_Collection.asp"-->
<!--#include file="../Include/PowerEasy.XmlHttp.asp"-->
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


Dim rs, sql, i 'ͨ�ñ���
Dim rsItem, ItemID, ItemName, strsql, NeedSave

ItemID = PE_CLng(Trim(Request("ItemID")))
NeedSave = Trim(Request("NeedSave"))          '�ж���Ŀ�Ƿ�����Ҫ����
strFileName = "Admin_CollectionManage.asp?Action=" & Action

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>�ɼ�����</title>" & vbCrLf
Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""Admin_Style.css"">" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">" & vbCrLf

Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"">" & vbCrLf
Call ShowPageTitle(" �� �� ϵ ͳ �� Ŀ �� �� ", 10052)
Response.Write "  <tr class=""tdbg""> " & vbCrLf
Response.Write "    <td width=""70"" height=""30""><strong>��������</strong></td>" & vbCrLf
Response.Write "    <td height=""30""><a href=Admin_CollectionManage.asp?Action=ItemManage>������ҳ</a> | <a href=""Admin_CollectionManage.asp?Action=Step1"">�������Ŀ</a> | <a href=Admin_CollectionManage.asp?Action=Import>������Ŀ</a> | <a href=""Admin_CollectionManage.asp?Action=Export"">������Ŀ</a></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
If InStr(Action, "Step") > 0 Then
    Response.Write "<br>�ɼ���Ŀ���ò��裺<a href=""Admin_CollectionManage.asp?Action=Step1&ItemID=" & ItemID & """>"
    If Action = "Step1" Then
        Response.Write "<font color=red>��������</font>"
    Else
        Response.Write "��������"
    End If
    Response.Write "</a> >> <a href=""Admin_CollectionManage.asp?Action=Step2&ItemID=" & ItemID & """>"
    If Action = "Step2" Then
        Response.Write "<font color=red>�б�����</font>"
    Else
        Response.Write "�б�����"
    End If
    Response.Write "</a> >> <a href=""Admin_CollectionManage.asp?Action=Step3&ItemID=" & ItemID & """>"
    If Action = "Step3" Then
        Response.Write "<font color=red>��������</font>"
    Else
        Response.Write "��������"
    End If
    Response.Write "</a> >> <a href=""Admin_CollectionManage.asp?Action=Step4&ItemID=" & ItemID & """>"
    If Action = "Step4" Then
        Response.Write "<font color=red>��������</font>"
    Else
        Response.Write "��������"
    End If
    Response.Write "</a> >> <a href=""Admin_CollectionManage.asp?Action=Step5&ItemID=" & ItemID & """>"
    If Action = "Step5" Then
        Response.Write "<font color=red>��������</font>"
    Else
        Response.Write "��������"
    End If
    Response.Write "</a> >> "
    If Action = "Step6" Then
        Response.Write "<font color=red>�������</font>"
    Else
        Response.Write "�������"
    End If
End If

Select Case Action
Case "Step1"                    '��Ŀ��������
    Call Step1
Case "Step2"                    '�б�����
    Call Step2
Case "Step3"                    '��������
    Call Step3
Case "Step4"                    '��������
    Call Step4
Case "Step5"                    '��������
    Call Step5
Case "Step6"                    '�������
    Call Step6
Case "Import"                   '��Ŀ�����һ��
    Call Import
Case "Import2"                  '��Ŀ����ڶ���
    Call Import2
Case "DoImport"                 '������Ŀ����
    Call DoImport
Case "Export"                   '������Ŀ
    Call Export
Case "DoExport"                 '������Ŀ����
    Call DoExport
Case "ItemManage"               '�ɼ��༭���Թ���
    Call ItemManage
Case "ItemCopy"                 '������Ŀ����
    Call ItemCopy
Case "DoItemCopy"               '��Ŀ���ƴ���
    Call DoItemCopy
Case "Batch"                    '����������Ŀ����
    Call Batch
Case "DoBatch"                  '����������Ŀ����
    Call DoBatch
Case "DelItem"
    Call DelItem
Case Else
    Call ItemManage
End Select
Response.Write "</body></html>"
Call CloseConn


Sub DelItem()
    Dim ItemID
    ItemID = Trim(Request("ItemID"))
	If IsValidID(ItemID) = False Then
		ItemID = ""
	End If

    If ItemID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<tr><td colspan='7'><li>��ָ��Ҫɾ������Ŀ��</td></tr></table><br>"
    Else
        If InStr(ItemID, ",") > 0 Then
            Conn.Execute ("Delete From [PE_Item] Where ItemID In (" & ItemID & ")")
            Conn.Execute ("Delete From [PE_Filters] Where ItemID In (" & ItemID & ")")
            Conn.Execute ("Delete From [PE_HistrolyNews] Where ItemID In (" & ItemID & ")")
        Else
            Conn.Execute ("Delete From [PE_Item] Where ItemID=" & ItemID)
            Conn.Execute ("Delete From [PE_Filters] Where ItemID=" & ItemID)
            Conn.Execute ("Delete From [PE_HistrolyNews] Where ItemID=" & ItemID)
        End If
    End If
    Call ItemManage
End Sub
'=================================================
'��������ItemManage
'��  �ã��ɼ���Ŀ�༭
'=================================================
Sub ItemManage()

    Call DataBaseModify

    Dim SqlH, RsH, Flag
    Dim iChannelID, ClassID, SpecialID, ItemID, ItemName, ListUrl, WebName, NewsCollecDate
    Dim SkinID, LayoutID, SkinCount, LayoutCount, MaxPerPage

    If Request("page") <> "" Then
        CurrentPage = CInt(Request("page"))
    Else
        CurrentPage = 1
    End If
    MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))
    iChannelID = PE_CLng(Trim(Request("iChannelID")))
    If MaxPerPage <= 0 Then MaxPerPage = 10
            
    strFileName = "Admin_CollectionManage.asp?Action=ItemManage&iChannelID=" & iChannelID
    
    Response.Write "<br>"
    
    If IsObjInstalled("MSXML2.XMLHTTP") = False Then
        Call WriteErrMsg("<li>����ϵͳû�а�װXMLHTTP ���,�뵽΢����վ����MSXML 4.0", ComeUrl)
        Exit Sub
    End If

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "<tr class='title'><td colspan='2'> | "
    sql = "SELECT DISTINCT I.ChannelID, C.ChannelName,C.ModuleType FROM PE_Item I LEFT OUTER JOIN PE_Channel C ON I.ChannelID = C.ChannelID where C.ModuleType=1"
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
    Else
        Do While Not rs.EOF
            If IsNull(rs("ChannelName")) Then
            Else
                Response.Write "<a href='Admin_CollectionManage.asp?Action=ItemManage&iChannelID=" & rs("ChannelID") & "'><FONT style='font-size:12px'"
                If rs("ChannelID") = iChannelID Then Response.Write "color='red'"
                Response.Write "> " & rs("ChannelName") & "</FONT></a> | "
            End If
            rs.MoveNext
        Loop
        Response.Write "<a href='Admin_CollectionManage.asp?Action=ItemManage&iChannelID=0'><FONT style='font-size:12px'"
        If iChannelID = 0 Then Response.Write "color='red'"
        Response.Write "> ����Ƶ�� </FONT></a> | "
    End If
    Response.Write "</td></tr>"
    Response.Write "</table>"
    rs.Close
    Set rs = Nothing
    Response.Write "<br>"
    Response.Write GetManagePath(iChannelID)
    Response.Write "<br>"
    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "    function unselectall(thisform){" & vbCrLf
    Response.Write "        if(thisform.chkAll.checked){" & vbCrLf
    Response.Write "            thisform.chkAll.checked = thisform.chkAll.checked&0;" & vbCrLf
    Response.Write "        }   " & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function CheckAll(thisform){" & vbCrLf
    Response.Write "        for (var i=0;i<thisform.elements.length;i++){" & vbCrLf
    Response.Write "            var e = thisform.elements[i];" & vbCrLf
    Response.Write "            if (e.Name != ""chkAll""&&e.disabled!=true)" & vbCrLf
    Response.Write "                e.checked = thisform.chkAll.checked;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<table class=""border"" border=""0"" cellspacing=""1"" width=""100%"" cellpadding=""0"">" & vbCrLf
    Response.Write "<form name=""myform"" method=""POST"" action=""Admin_CollectionManage.asp"">" & vbCrLf
    Response.Write "  <tr class=""title"" style=""padding: 0px 2px;"">" & vbCrLf
    Response.Write "    <td width=""40"" height=""22"" align=""center""><strong>ѡ��</strong></td>        " & vbCrLf
    Response.Write "    <td width=""100"" align=""center""><strong>��Ŀ����</strong></td>" & vbCrLf
    Response.Write "    <td width=""100"" align=""center""><strong>�ɼ���ַ</strong></td>" & vbCrLf
    Response.Write "    <td width=""100"" height=""22"" align=""center""><strong>����Ƶ��</strong></td> " & vbCrLf
    Response.Write "    <td width=""100"" height=""22"" align=""center""><strong>������Ŀ</strong></td> " & vbCrLf
    Response.Write "    <td width=""40"" align=""center""><strong>������</strong></td>        " & vbCrLf
    Response.Write "    <td width=""120"" height=""22"" align=""center""><strong>�ϴβɼ�ʱ��</strong>" & vbCrLf
    Response.Write "    <td width=""120"" height=""22"" align=""center""><strong>����</strong></td>   " & vbCrLf
    Response.Write "  </tr>" & vbCrLf
        
    sql = "SELECT I.*,C.ChannelName,CL.ClassName,C.Disabled,C.ModuleType"
    sql = sql & " FROM (PE_Item I left JOIN PE_Channel C ON I.ChannelID =C.ChannelID)"
    sql = sql & " Left JOIN PE_Class CL ON I.ClassID = CL.ClassID"
    sql = sql & " where C.ModuleType=1"
    If iChannelID <> 0 Then sql = sql & " And I.ChannelID=" & iChannelID
    sql = sql & " ORDER BY I.ItemID DESC"

    
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 1
    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg'><td height='50' align='center' colspan='8'>ϵͳ�����޲ɼ���Ŀ��</td></tr></table>"
    Else
        totalPut = rs.RecordCount
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
            ChannelID = rs("ChannelID")
            ClassID = PE_CLng(rs("ClassID"))

            ItemID = rs("ItemID")
            ItemName = rs("ItemName")
            ListUrl = rs("ListStr")
            WebName = rs("WebName")
            NewsCollecDate = rs("NewsCollecDate")
            Flag = rs("Flag")

            Response.Write "<tr class=""tdbg"" onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"" style=""padding: 0px 2px;"">" & vbCrLf
            Response.Write "  <td width=""40"" align=""center"">" & vbCrLf
            Response.Write "    <input type=""checkbox"" value=" & ItemID & " name=""ItemID""> " & vbCrLf
            Response.Write "  </td>" & vbCrLf
            Response.Write "  <td width=""100"" align=""center"">" & ItemName & "</td> " & vbCrLf
            Response.Write "  <td width=""100"" align=""center""><a href=" & ListUrl & " target=""_bank"">" & WebName & "</a></td>  " & vbCrLf
            Response.Write "  <td width=""100"" height=""22"" align=""center"">"
            If IsNull(rs("ChannelName")) = True Then
                Response.Write "��û��ָ��Ƶ��"
            Else
                If rs("Disabled") = True Then
                    Response.Write rs("ChannelName") & "<font color=red>&nbsp;�ѽ���</font>"
                Else
                    Response.Write rs("ChannelName")
                End If
            End If
            Response.Write "</td> " & vbCrLf
            Response.Write "  <td width=""100"" align=""center"">"
            If IsNull(rs("ClassName")) = True Then
                Response.Write "��û��ָ����Ŀ"
            Else
                Response.Write rs("ClassName")
            End If
            Response.Write "</td>" & vbCrLf
            Response.Write "  <td width=""40"" align=""center"">" & vbCrLf
            If Flag = True Then
                Response.Write "<b>��</b>"
            Else
                Response.Write "<FONT color='red'><b>��</b></FONT>"
            End If
            Response.Write "  </td>" & vbCrLf
            Response.Write "  <td width=""120"" align=""center"">" & vbCrLf
            If DateDiff("d", NewsCollecDate, Now()) = 0 Then
                Response.Write "<font color=red>" & NewsCollecDate & "</font>"
            Else
                Response.Write NewsCollecDate
            End If
            Response.Write "  </td>" & vbCrLf
            Response.Write "  <td width=""120"" align=""center""> " & vbCrLf
            Response.Write "     <a href=Admin_CollectionManage.asp?Action=Step1&ItemID=" & ItemID & ">�༭</a> "
            Response.Write "     <a href=Admin_CollectionManage.asp?Action=Step4&ItemID=" & ItemID & ">����</a> "
            Response.Write "     <a href=Admin_CollectionManage.asp?Action=Step5&ItemID=" & ItemID & ">����</a> "
            Response.Write "     <a href=Admin_CollectionManage.asp?Action=ItemCopy&ItemID=" & ItemID & ">����</a>" & vbCrLf
            Response.Write "   </td> " & vbCrLf
            Response.Write "</tr> " & vbCrLf

            VisitorNum = VisitorNum + 1
            If VisitorNum >= MaxPerPage Then Exit Do
            rs.MoveNext
        Loop
        Response.Write "</table>"
    End If
    Response.Write "<table border='0' cellpadding='0' cellspacing='1' width='100%' height='5'>"
    Response.Write "  <tr><td></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<input name=""Action"" type=""hidden""  value=''>" & vbCrLf
    Response.Write "<input name=""chkAll"" type=""checkbox"" id=""chkAll"" onclick=CheckAll(this.form) value=""checkbox"" >ѡ��������Ŀ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "<input type=""submit"" value="" ����ɾ�� ""  onclick=""javascript:if (confirm('���Ƿ�Ҫɾ��ѡ���Ĳɼ���Ŀ��')){document.myform.Action.value='DelItem';}else{return false;};"" >&nbsp;&nbsp;" & vbCrLf
    Response.Write "<input type='submit' name='Submit3' value="" �������� "" onClick=""document.myform.Action.value='Batch'"">"
    Response.Write "</form>"
    rs.Close
    Set rs = Nothing

    If totalPut > 0 Then
        Response.Write "<center>" & ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "����Ŀ��¼", True) & "</center>"
    End If
    Response.Write "<br><b>ע�⣺</b><br>&nbsp;&nbsp;&nbsp;&nbsp;�ɼ���Ŀ���뾭��<font color=red>�������Գɹ�</font>��<font color=red>�����ú�����</font>���������С�</font>" & vbCrLf
    Call CloseConn
End Sub
'=================================================
'��������Step1
'��  �ã���������
'=================================================
Sub Step1()
    Dim ItemName, ItemDoem, WebName, WebUrl, ListStr
    Dim LoginType, LoginUrl, LoginPostUrl, LoginUser, LoginPass, LoginFalse
    Dim arrLoginUser, arrLoginPass, InputLoginUser, InputLoginPass
    
    If ItemID > 0 Then
        'ȡ������
        sql = "select ItemName,ItemDoem,WebName,WebUrl,ListStr,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse from PE_Item where ItemID=" & ItemID
        Set rsItem = Server.CreateObject("adodb.recordset")
        rsItem.Open sql, Conn, 1, 1
        If rsItem.EOF Then   'û���ҵ�����Ŀ
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���������û���ҵ�����Ŀ��</li>"
        Else
            ItemName = rsItem("ItemName")
            ItemDoem = rsItem("ItemDoem")
            WebName = rsItem("WebName")
            WebUrl = PE_CLng(rsItem("WebUrl"))
            ListStr = rsItem("ListStr")
            LoginType = PE_CLng(rsItem("LoginType"))
            LoginUrl = rsItem("LoginUrl")
            LoginPostUrl = rsItem("LoginPostUrl")
            LoginUser = rsItem("LoginUser")
            LoginPass = rsItem("LoginPass")
            LoginFalse = rsItem("LoginFalse")
        End If
        rsItem.Close
        Set rsItem = Nothing
    End If
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    If InStr(LoginUser, "=") > 0 Then
        arrLoginUser = Split(LoginUser, "=")
        InputLoginUser = arrLoginUser(0)
        LoginUser = arrLoginUser(1)
    End If
    If InStr(LoginPass, "=") > 0 Then
        arrLoginPass = Split(LoginPass, "=")
        InputLoginPass = arrLoginPass(0)
        LoginPass = arrLoginPass(1)
    End If

    Call ShowChekcFormVbs

    Response.Write "<FORM name=form1 action='Admin_CollectionManage.asp' method=post>" & vbCrLf
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center' height='24'>" & vbCrLf
    Response.Write "   <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>��������</td>" & vbCrLf
    Response.Write "   <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>��¼����</td>" & vbCrLf
    Response.Write "   <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td width=""120"" class='tdbg5' align=""right"" > ��Ŀ���ƣ�&nbsp;</td>" & vbCrLf
    Response.Write "          <td class=""tdbg""><input name=""ItemName"" type=""text"" id=""ItemName"" size=""30"" maxlength=""30"" value='" & ItemName & "'>&nbsp;<FONT color='red'>*</FONT></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td width=""120"" class='tdbg5' align=""right""> ��վ���ƣ�&nbsp;</td>" & vbCrLf
    Response.Write "          <td class=""tdbg""><input name=""WebName"" type=""text"" id=""WebName"" size=""30"" maxlength=""30"" value='" & WebName & "'> &nbsp;<FONT color='red'>*</FONT></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td width=""120"" class='tdbg5' align=""right""> ��ҳ���룺&nbsp;</td>" & vbCrLf
    Response.Write "          <td>" & vbCrLf
    Response.Write "            <INPUT TYPE='radio' NAME='WebUrl' value='0' " & IsRadioChecked(WebUrl, 0) & "> GB2312" & vbCrLf
    Response.Write "            <INPUT TYPE='radio' NAME='WebUrl' value='1' " & IsRadioChecked(WebUrl, 1) & "> UTF-8" & vbCrLf
    Response.Write "            <INPUT TYPE='radio' NAME='WebUrl' value='2' " & IsRadioChecked(WebUrl, 2) & "> Big5&nbsp;<FONT color='red'>*</FONT>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td width=""120"" class='tdbg5' align=""right""> �б�ҳURL��&nbsp;</td>" & vbCrLf
    Response.Write "          <td class=""tdbg""><input name=""ListStr"" type=""text"" id=""ListStr"" size=""70"" maxlength=""255"" value='" & ListStr & "'>&nbsp;<FONT color='red'>*</FONT> <br><font color=blue> ���磺http://www.powereasy.net/News/ShowClass.asp?ClassID=2</font></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "       <tr class=""tdbg""> " & vbCrLf
    Response.Write "         <td width=""120"" class='tdbg5' align=""right""> ��Ŀ��ע��&nbsp;</td>" & vbCrLf
    Response.Write "         <td>" & vbCrLf
    Response.Write "           <textarea name=""ItemDoem"" style='width:450px;height:100px' id=""ItemDoem"">" & ItemDoem & "</textarea>" & vbCrLf
    Response.Write "         </td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "       </tbody>" & vbCrLf
    Response.Write "       <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "       <tr class=""tdbg""> " & vbCrLf
    Response.Write "        <td width=""120"" class='tdbg5' align=""right"">��վ��¼��</td>" & vbCrLf
    Response.Write "        <td width=""620"">" & vbCrLf
    Response.Write "         <input type=""radio"" value=""0"" name=""LoginType"" " & IsRadioChecked(LoginType, 0) & " onClick=""Login.style.display='none'"">����Ҫ��¼<span lang=""en-us""></span>" & vbCrLf
    Response.Write "         <input type=""radio"" value=""1"" name=""LoginType"" " & IsRadioChecked(LoginType, 1) & " onClick=""Login.style.display=''"">���ò���  <FONT style='font-size:12px' color='blue'>��ֻ���ڶԷ���վû�п�����¼��֤�빦��ʱ�����ܽ��е�¼�ɼ���</FONT> </td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "       <tr class=""tdbg"" id=""Login"" " & IsStyleDisplay(LoginType, 1) & ">" & vbCrLf
    Response.Write "         <td align='left' colspan='2'>" & vbCrLf
    Response.Write "           <table border='0' cellpadding='0' cellspacing='1' width=""620"" height='100%' align='left' bgcolor='#ffffff'>" & vbCrLf
    Response.Write "             <tr>" & vbCrLf
    Response.Write "               <td width=""130"" class=""tdbg5"" align=""right""> ��¼��ַ��&nbsp;&nbsp;</td>" & vbCrLf
    Response.Write "               <td class=""tdbg""><input name=""LoginUrl"" type=""text"" size=""40"" maxlength=""150"" value=""" & LoginUrl & """></td>" & vbCrLf
    Response.Write "             </tr>" & vbCrLf
    Response.Write "             <tr>" & vbCrLf
    Response.Write "               <td width=""130"" class=""tdbg5"" align=""right""> �ύ��ַ��&nbsp;&nbsp;</td>" & vbCrLf
    Response.Write "               <td class=""tdbg""><input name=""LoginPostUrl"" type=""text"" size=""40"" maxlength=""150"" value=""" & LoginPostUrl & """></td>" & vbCrLf
    Response.Write "             </tr>" & vbCrLf
    Response.Write "             <tr>" & vbCrLf
    Response.Write "               <td width=""130"" class=""tdbg5"" align=""right""> �û�������&nbsp;&nbsp;</td>" & vbCrLf
    Response.Write "               <td class=""tdbg"">�û��ı������ƣ�<input name=""InputLoginUser"" type=""text"" size=""10"" maxlength=""150"" value=""" & InputLoginUser & """>" & vbCrLf
    Response.Write "               �û����ƣ�<input name=""LoginUser"" type=""text"" size=""10"" maxlength=""150"" value=""" & LoginUser & """></td>" & vbCrLf
    Response.Write "             </tr>" & vbCrLf
    Response.Write "             <tr>" & vbCrLf
    Response.Write "               <td width=""130"" class=""tdbg5"" align=""right""> ���������&nbsp;&nbsp;</td>" & vbCrLf
    Response.Write "               <td class=""tdbg"">�����ı������ƣ�<input name=""InputLoginPass"" type=""text"" size=""10"" maxlength=""150"" value=""" & InputLoginPass & """>" & vbCrLf
    Response.Write "                �û����룺<input name=""LoginPass"" type=""text"" size=""10"" maxlength=""150"" value=""" & LoginPass & """>" & vbCrLf
    Response.Write "               </td>" & vbCrLf
    Response.Write "             </tr>" & vbCrLf
    Response.Write "             <tr>" & vbCrLf
    Response.Write "               <td width=""130"" class=""tdbg5"" align=""right""> ʧ����Ϣ��&nbsp;&nbsp;</td>" & vbCrLf
    Response.Write "               <td class=""tdbg""><input name=""LoginFalse"" type=""text"" size=""30"" maxlength=""150"" value=""" & LoginFalse & """></td>" & vbCrLf
    Response.Write "             </tr>" & vbCrLf
    Response.Write "           </table>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "       </td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "     </table>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "   </tr>" & vbCrLf
    Response.Write " </table>" & vbCrLf
    Response.Write " <br>" & vbCrLf
    Response.Write " <center>" & vbCrLf
    Response.Write "   <INPUT id=""ItemID"" type=""hidden"" value=" & ItemID & " name=ItemID>" & vbCrLf
    Response.Write "   <INPUT id=""NeedSave"" type=""hidden"" value='True' name='NeedSave'>" & vbCrLf
    If ItemID = 0 Then
        Response.Write "   <INPUT id=""IsNew"" type=""hidden"" value='True' name='IsNew'>" & vbCrLf
    End If
    Response.Write "   <INPUT id=""Action"" type=""hidden"" value=""Step2"" name=Action>" & vbCrLf
    Response.Write "   <INPUT id=Cancel  type=button value="" ȡ  �� "" name='Cancel' onclick=""window.location.href='Admin_CollectionManage.asp'"">&nbsp;&nbsp;" & vbCrLf
    Response.Write "   <INPUT  type=submit value="" ��һ�� "" name=""Submit""></td>" & vbCrLf
    Response.Write " </center>" & vbCrLf
    Response.Write "</FORM>" & vbCrLf
    Call CloseConn
End Sub
'=================================================
'��������Step2
'��  �ã��б�����
'=================================================
Sub Step2()
    Dim ItemName, WebName, WebUrl, ItemDoem
    Dim ListStr, LsString, LoString, ListPaingType, LPsString, LPoString, ListPaingStr1, ListPaingStr2
    Dim HsString, HoString, HttpUrlType, HttpUrlStr
    Dim ListPaingID1, ListPaingID2, ListPaingStr3, IsNew
    Dim LoginType, LoginUrl, LoginPostUrl, LoginUser, LoginPass, LoginFalse, LoginData, LoginResult
    Dim InputLoginUser, InputLoginPass

    '�б�����ͼ
    Dim ThumbnailType, ThsString, ThoString

    IsNew = Trim(Request("IsNew"))          '�ж���Ŀ�Ƿ������

    If NeedSave = "True" Then
        ItemName = Trim(Request.Form("ItemName"))
        WebName = Trim(Request.Form("WebName"))
        WebUrl = Trim(Request.Form("WebUrl"))
        ItemDoem = Request.Form("ItemDoem")
        ListStr = Trim(Request.Form("ListStr"))
        LoginType = Trim(Request.Form("LoginType"))
        LoginUrl = Trim(Request.Form("LoginUrl"))
        LoginPostUrl = Trim(Request.Form("LoginPostUrl"))
        InputLoginUser = Trim(Request.Form("InputLoginUser"))
        InputLoginPass = Trim(Request.Form("InputLoginPass"))
        LoginUser = Trim(Request.Form("LoginUser"))
        LoginPass = Trim(Request.Form("LoginPass"))
        LoginFalse = Trim(Request.Form("LoginFalse"))
        '���ӵ�¼��ֵ
        LoginUser = InputLoginUser & "=" & LoginUser
        LoginPass = InputLoginPass & "=" & LoginPass

        If IsNew <> "True" And ItemID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵĲɼ���Ŀ��</li>"
        End If
        
        If ItemName = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��Ŀ���Ʋ���Ϊ��</li>"
        End If
        If WebName = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��վ���Ʋ���Ϊ��</li>"
        End If
        If WebUrl = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��վ�������Ͳ���Ϊ��</li>"
        End If
        If ListStr = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�б���ַ����Ϊ��</li>"
        End If
        If CheckUrl(ListStr) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�б���ַ����</li>"
        End If

        If LoginType = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ѡ����վ��¼����</li>"
        Else
            LoginType = CLng(LoginType)
            If LoginType = 1 Then
                If LoginUrl = "" Or LoginPostUrl = "" Or LoginUser = "" Or LoginPass = "" Or LoginFalse = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>��վ��¼��Ϣ������</li>"
                Else
                    LoginData = UrlEncoding(LoginUser & "&" & LoginPass)
                    LoginResult = PostHttpPage(LoginUrl, LoginPostUrl, LoginData, PE_CLng(WebUrl))
                    If InStr(LoginResult, LoginFalse) > 0 Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>��¼��վʱ����������ȷ�ϵ�¼��Ϣ����ȷ�ԣ�</li>"
                    End If
                End If
            End If
        End If

        If FoundErr = True Then
            Call WriteErrMsg(ErrMsg, ComeUrl)
            Exit Sub
        End If

        sql = "Select top 1 ItemID,ItemName,WebName,WebUrl,ListStr,ItemDoem,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,ChannelID from PE_Item"
        If IsNew <> "True" Then
            sql = sql & " where ItemID=" & ItemID
        End If
        Set rsItem = Server.CreateObject("adodb.recordset")
        rsItem.Open sql, Conn, 1, 3
        If IsNew = "True" Then
            rsItem.addnew
        End If
        rsItem("ItemName") = ItemName
        rsItem("WebName") = WebName
        rsItem("WebUrl") = WebUrl
        rsItem("ListStr") = ListStr
        rsItem("LoginType") = LoginType
        rsItem("LoginUrl") = LoginUrl
        rsItem("LoginPostUrl") = LoginPostUrl
        rsItem("LoginUser") = LoginUser
        rsItem("LoginPass") = LoginPass
        rsItem("LoginFalse") = LoginFalse
        rsItem("ItemDoem") = ItemDoem
        If IsNew = "True" Then
            rsItem("ChannelID") = 1
        End If
        
        rsItem.Update
        rsItem.Close
        Set rsItem = Nothing
        If IsNew = "True" Then
            Dim mrs
            Set mrs = Conn.Execute("select max(ItemID) from PE_Item")
            If IsNull(mrs(0)) Then
                ItemID = 1
            Else
                ItemID = mrs(0)
            End If
            Set mrs = Nothing
        End If
    End If

    sql = "Select top 1 WebUrl,ListStr,LsString,LoString,ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,ListStr,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,HsString,HoString,HttpUrlType,HttpUrlStr,ThumbnailType,ThsString,ThoString from PE_Item Where ItemID=" & ItemID
    Set rsItem = Server.CreateObject("adodb.recordset")
    rsItem.Open sql, Conn, 1, 1
    If rsItem.EOF And rsItem.BOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>û���ҵ�����Ŀ!</li>"
    Else
        LoginType = rsItem("LoginType")
        LoginUrl = rsItem("LoginUrl")
        LoginPostUrl = rsItem("LoginPostUrl")
        LoginUser = rsItem("LoginUser")
        LoginPass = rsItem("LoginPass")
        LoginFalse = rsItem("LoginFalse")
        ListStr = rsItem("ListStr")
        LsString = rsItem("LsString")
        LoString = rsItem("LoString")
        ListPaingType = rsItem("ListPaingType")
        LPsString = rsItem("LPsString")
        LPoString = rsItem("LPoString")
        ListPaingStr1 = rsItem("ListPaingStr1")
        ListPaingStr2 = rsItem("ListPaingStr2")
        ListPaingID1 = rsItem("ListPaingID1")
        ListPaingID2 = rsItem("ListPaingID2")
        ListPaingStr3 = rsItem("ListPaingStr3")

        ThumbnailType = PE_CLng(rsItem("ThumbnailType"))
        ThsString = rsItem("ThsString")
        ThoString = rsItem("ThoString")

        ListStr = rsItem("ListStr")
        WebUrl = rsItem("WebUrl")
        HsString = rsItem("HsString")
        HoString = rsItem("HoString")
        HttpUrlType = rsItem("HttpUrlType")
        HttpUrlStr = rsItem("HttpUrlStr")
    End If
    rsItem.Close
    Set rsItem = Nothing

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If


    Dim strPageContent
    strPageContent = GetHttpPage(ListStr, PE_CLng(WebUrl))
    If strPageContent = "$False$" Then
        FoundErr = True
        ErrMsg = ErrMsg & "�ɼ���Ŀ����վʧ�ܣ�ʧ��ԭ������ǣ�<br>"
        ErrMsg = ErrMsg & "1�����ķ������Ƿ������ MSXML2.XMLHTTP ���<br>"
        ErrMsg = ErrMsg & "2������������������Ƿ�����<br>"
        ErrMsg = ErrMsg & "3�����ķ������Ƿ�װ�˷���ǽ�����ҹر����йض˿ڡ�ϵͳ�ڲɼ�ʱ��Ҫ�������һ���˿�������Է�������ͨ�ţ�����ر�����Щ�˿ڣ���ᵼ����Ϊ�޷�ͨ�Ŷ��ɼ�ʧ�ܡ�<br>" & vbCrLf
        ErrMsg = ErrMsg & "4�����������վ�ܲɼ������ɼ�����վʱ���ֱ���ʾ��˵������վ�ķ�������װ�˷���ǽ���ر����йض˿ڣ����ߴ���վ�Ѿ����رա�" & vbCrLf
    End If
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    Call ShowChekcFormVbs
        
    Response.Write "<form method=""post"" action=""Admin_CollectionManage.asp"" name=""form1"">" & vbCrLf
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center' height='24'>" & vbCrLf
    Response.Write "   <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>��������</td>" & vbCrLf
    Response.Write "   <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>��ҳ����</td>" & vbCrLf
    Response.Write "   <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>�б�����ͼ</td>" & vbCrLf
    Response.Write "   <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>����Ԥ��</td>" & vbCrLf
    Response.Write "   <td id='TabTitle' class='title5' onclick=""ShowTabs(4):setFileFields('" & ListStr & "')"";'>��ҳԤ��</td>" & vbCrLf
    Response.Write "   <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr align='left' class='tdbg'><td width='5'></td>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='720' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right""> �б�ʼ���룺</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" width=""600"">"
    Response.Write "            <textarea name=""LsString"" style='width:450px;height:100px' id=""LsString"">"
    If Trim(LsString) <> "" Then Response.Write Server.HTMLEncode(LsString & "")
    Response.Write "</textarea>&nbsp;<FONT color='red'>*</FONT><input TYPE='button' value='���Դ���' onCLICK='ceshi(1)' ></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right""> �б�������룺</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"">"
    Response.Write "            <textarea name=""LoString"" style='width:450px;height:100px' id=""LoString"">"
    If Trim(LoString) <> "" Then Response.Write Server.HTMLEncode(LoString & "")
    Response.Write "</textarea>&nbsp;<FONT color='red'>*</FONT><input TYPE='button' value='���Դ���' onCLICK='ceshi(2)' ></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf

    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "           <td width=""120"" class=""tdbg5"" align='right'> ���ӿ�ʼ���룺</td>" & vbCrLf
    Response.Write "           <td class=""tdbg"">"
    Response.Write "             <textarea name=""HsString"" style='width:450px;height:40px' id=""HsString"">"
    If Trim(HsString) <> "" Then Response.Write Server.HTMLEncode(HsString & "")
    Response.Write "</textarea>&nbsp;<FONT color='red'>*</FONT></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "           <td width=""120"" class=""tdbg5"" align='right'> ���ӽ������룺</td>" & vbCrLf
    Response.Write "           <td class=""tdbg"">"
    Response.Write "             <textarea name=""HoString"" style='width:450px;height:40px' id=""HoString"">"
    If Trim(HoString) <> "" Then Response.Write Server.HTMLEncode(HoString & "")
    Response.Write "</textarea>&nbsp;<FONT color='red'>*</FONT></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "           <td width=""120"" class=""tdbg5"" align='right'></td>" & vbCrLf
    Response.Write "           <td class=""tdbg"">"
    Response.Write "             <FONT color='#0099FF'>���磺�б��е����Ӵ������磺&lt;a href='Article/Class1/1358.html' target='_blank'&gt;<br>�����ӿ�ʼ����Ӧ������Ϊ��</font><font color=red>&lt;a href='</font><FONT color='#0099FF'>�����ӽ�����������Ϊ��</font><font color=red>' target='_blank'&gt;</font>"
    Response.Write "           </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "           <td width=""120"" class=""tdbg5"" align='right'> ���⴦��</td>" & vbCrLf
    Response.Write "           <td class=""tdbg"" >" & vbCrLf
    Response.Write "    <input type=""radio"" value=""0"" name=""HttpUrlType"""
    If HttpUrlType = 0 Then Response.Write "checked"
    Response.Write "    onClick=""javascript:HttpUrl1.style.display='none'"">�ر�&nbsp;" & vbCrLf
    Response.Write "    <input type=""radio"" value=""1"" name=""HttpUrlType"""
    If HttpUrlType = 1 Then Response.Write "checked"
    Response.Write "    onClick=""javascript:HttpUrl1.style.display=''"">����" & vbCrLf
    Response.Write "            </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg"" id=""HttpUrl1"" style=""display:'"
    If HttpUrlType = 0 Then Response.Write "none"
    Response.Write "'"">" & vbCrLf
    Response.Write "           <td width=""120"" class=""tdbg5"" align='right' valign='top'>�ض���URL��</td>" & vbCrLf
    Response.Write "           <td class=""tdbg"" >" & vbCrLf
    Response.Write "             <input name=""HttpUrlStr"" type=""text"" size=""49"" maxlength=""200"" value=" & HttpUrlStr & ">" & vbCrLf
    Response.Write "             <br><font color='#0099FF'>�����Ӵ�����һЩ�ǳ������JS�������ô���ʱ�������ô�ѡ�<br>���磺�б��е����Ӵ������磺&lt;a href='#' onclick='opennews(137)'&gt;����Ӧ��opennews(id)�����Ĵ���Ϊ��<br>    window.open('http://www.xxxx.com/xxx/news.asp?id='+id,'','****')��<br>�����ӿ�ʼ��������Ϊ��</font><font color=red> &lt;a href='#' onclick='opennews(</font><font color='#0099FF'>�����ӽ�������Ϊ��<font color=red>)'&gt;</font><font color='#0099FF'>��<br>�˴����ض���URL������Ϊ��</font><font color=red>http://www.xxxx.com/xxx/news.asp?id={$ID}</font><font color='#0099FF'>��{$ID}��ϵͳ�涨�ı�ǩ��</font>" & vbCrLf
    Response.Write "           </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right""> ѡ���ҳ���ͣ�</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" width=""600"">" & vbCrLf
    Response.Write "            <input type=""radio"" value=""0"" name=""ListPaingType""" & IsRadioChecked(ListPaingType, 0) & " onClick=""javascript:ListPaing1.style.display='none';ListPaing2.style.display='none';ListPaing3.style.display='none'"">���ɼ��б��ҳ&nbsp;" & vbCrLf
    Response.Write "            <input type=""radio"" value=""1"" name=""ListPaingType""" & IsRadioChecked(ListPaingType, 1) & " onClick=""javascript:ListPaing1.style.display='';ListPaing2.style.display='none';ListPaing3.style.display='none'"">��Դ�����л�ȡ��һҳ��URL&nbsp;" & vbCrLf
    Response.Write "            <input type=""radio"" value=""2"" name=""ListPaingType""" & IsRadioChecked(ListPaingType, 2) & " onClick=""javascript:ListPaing1.style.display='none';ListPaing2.style.display='';ListPaing3.style.display='none'"">����ָ����ҳURL����&nbsp;" & vbCrLf
    Response.Write "            <input type=""radio"" value=""3"" name=""ListPaingType""" & IsRadioChecked(ListPaingType, 3) & " onClick=""javascript:ListPaing1.style.display='none';ListPaing2.style.display='none';ListPaing3.style.display=''"">�ֶ���ӷ�ҳURL���� " & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg"" id=""ListPaing1"" style=""display:'"
    If ListPaingType <> 1 Then Response.Write "none"
    Response.Write "'"">" & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right"">����һҳ��<br>URL��ʼ���룺<br><br><br><br><br><br>" & vbCrLf
    Response.Write "            ����һҳ��<br>URL�������룺</font>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" width=""600"">" & vbCrLf
    Response.Write "            <textarea name=""LPsString"" style='width:450px;height:100px'>"
    If Trim(LPsString) <> "" Then Response.Write Server.HTMLEncode(LPsString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(3)' ><br>" & vbCrLf
    Response.Write "            <textarea name=""LPoString"" style='width:450px;height:100px'>"
    If Trim(LPoString) <> "" Then Response.Write Server.HTMLEncode(LPoString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(4)' >" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf

    'Response.Write "        <tr class=""tdbg"" id=""ListPaing12"" style=""display:'"
    'If ListPaingType <> 1 Then Response.Write "none"
    'Response.Write "'"">" & vbCrLf
    'Response.Write "          <td width=""120"" class=""tdbg5"" align=""right"">������ҳ�ض���</td>" & vbCrLf
    'Response.Write "          <td class=""tdbg"" width=""600"">" & vbCrLf
    'Response.Write "            <input name=""ListPaingStr1"" type=""text"" size=""60"" maxlength=""200"" value=" & ListPaingStr1 & ">" & vbCrLf
    'Response.Write "            <br><font color=#0099FF>һ�㲻���õ�,����ɼ���ҳ������,������һҳ���������·����" & vbCrLf
    'Response.Write "            <br>����һ���������÷���������һҳ�б��URL��ʵ�ʲ���,Ӧ�ô˹��ܡ�" & vbCrLf
    'Response.Write "            <br>���б����ò������·��,����Ƕ�̬ҳ����ID��" & vbCrLf
    'Response.Write "            <br>������������ҳ����дʵ��·�� http://www.xxxxx.com/xxx/xx/xxx/news/{$ID}  {$ID}�����б�������·����̬ID��</font>" & vbCrLf
    'Response.Write "            <br>ϵͳ�����ܷ�����վ�����·��,������������������,�밴��������ʹ�ô˹��ܡ�"
    'Response.Write "          </td>" & vbCrLf
    'Response.Write "        </tr>" & vbCrLf

    Response.Write "        <tr class=""tdbg"" id=""ListPaing2"" style=""display:'"
    If ListPaingType <> 2 Then Response.Write "none"
    Response.Write "'"">" & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right"">URL�ַ�����<br><br><br>ID��Χ��</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" width=""600""><input name=""ListPaingStr2"" type=""text"" size=""80"" maxlength=""200"" value=" & ListPaingStr2 & "><br>" & vbCrLf
    Response.Write "            <font color=#0099FF>����http://www.xxxxx.com/news/index_{$ID}.html&nbsp;&nbsp;&nbsp;&nbsp;{$ID}�����ҳ��</font><br>" & vbCrLf
    Response.Write "            <br>" & vbCrLf
    Response.Write "            <input name=""ListPaingID1"" type=""text"" size=""8"" maxlength=""200"" value=" & ListPaingID1 & "><span lang=""en-us""> To </span><input name=""ListPaingID2"" type=""text"" size=""8"" maxlength=""200"" value=" & ListPaingID2 & ">" & vbCrLf
    Response.Write "            <font color=#0099FF>���� 1 ~ 9 �� 9 ~ 1 �������ɼ�</font><br>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg"" id=""ListPaing3"" style=""display:'"
    If ListPaingType <> 3 Then Response.Write "none"
    Response.Write "'"">" & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right"">URL�б�&nbsp;</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <textarea name=""ListPaingStr3"" style='width:500px;height:100px'>"
    If Trim(ListPaingStr3) <> "" Then Response.Write Server.HTMLEncode(ListPaingStr3 & "")
    Response.Write "</textarea>" & vbCrLf
    Response.Write "            <br><font color=#0099FF>ע��һ��дһ����ҳ��ַ</font>" & vbCrLf
    Response.Write "         </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf

    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right""> ����ͼ���ã�</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" width=""600"">" & vbCrLf
    Response.Write "            <input type=""radio"" value=""0"" name=""ThumbnailType""" & IsRadioChecked(ThumbnailType, 0) & " onClick=""javascript:ThumbnailPaing.style.display='none';"">������&nbsp;" & vbCrLf
    Response.Write "            <input type=""radio"" value=""1"" name=""ThumbnailType""" & IsRadioChecked(ThumbnailType, 1) & " onClick=""javascript:ThumbnailPaing.style.display='';"">����&nbsp; <FONT style='font-size:12px' color='blue'>ע�������ڽ�ȡһЩ�б�ҳ������ͼ����վ</FONT> " & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg"" id=""ThumbnailPaing"" style=""display:'"
    If ThumbnailType <> 1 Then Response.Write "none"
    Response.Write "'"">" & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right""><br>����ͼ��ʼ���룺<br><br><br><br><br><br>" & vbCrLf
    Response.Write "            <br>����ͼ�������룺</font>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" width=""600"">" & vbCrLf
    Response.Write "            <textarea name=""ThsString"" style='width:450px;height:100px'>"
    If Trim(ThsString) <> "" Then Response.Write Server.HTMLEncode(ThsString & "")
    Response.Write "</textarea><br>" & vbCrLf
    Response.Write "            <textarea name=""ThoString"" style='width:450px;height:100px'>"
    If Trim(ThoString) <> "" Then Response.Write Server.HTMLEncode(ThoString & "")
    Response.Write "</textarea>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf


    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "          <textarea name=""Content""  style='width:785px;height:400px'>" & Server.HTMLEncode(strPageContent & "") & "</textarea>" & vbCrLf
    Response.Write "         </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "         <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td align='center' id='objFiles'></td>" & vbCrLf
    Response.Write "         </tr>" & vbCrLf
    Response.Write "       </td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    Response.Write "   </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write " </table>" & vbCrLf
    Response.Write " <br>" & vbCrLf
    Response.Write " <center>" & vbCrLf
    Response.Write "   <input name=""ListStr"" type=""hidden"" id=""ListStr"" value=""" & ListStr & """>" & vbCrLf
    Response.Write "   <input name=""ItemID"" type=""hidden"" id=""ItemID"" value=""" & ItemID & """>" & vbCrLf
    Response.Write "   <input name=""Action"" type=""hidden"" id=""Action"" value=""Step3"">" & vbCrLf
    Response.Write "   <input name=""NeedSave"" type=""hidden"" id=""NeedSave"" value=""True"">" & vbCrLf
    Response.Write "   <input TYPE='button' value='������һ��'  onCLICK='history.back(-1)'>  &nbsp;&nbsp;" & vbCrLf
    Response.Write "   <input  type=""submit"" name=""Submit"" value=""�� һ ��""  onClick='CheckForm()'>" & vbCrLf
    Response.Write " </center>" & vbCrLf
    Response.Write "</FORM>" & vbCrLf
    Response.Write "<b>ע�⣺</b><br>&nbsp;&nbsp;&nbsp;&nbsp;��ʼ������������<font color=red>������һ������ҳ����Ψһ��</font>�����ܱ�֤������ȷ�ɼ���������ݡ���Ϊÿ���б�ҳ�Ĵ��붼���ܲ�ͬ��������Ҫ����������б�ҳ���ҵ���ͬ�Ŀ�ʼ����ͽ������룬���ܱ�֤���Դ������б�ҳ��׼ȷ�ɼ����������ݡ�" & vbCrLf

    Call CloseConn
End Sub

'=================================================
'��������Step3
'��  �ã���������
'=================================================
Sub Step3()

    Dim LoginResult, LoginData
    Dim LoginType, LoginUrl, LoginPostUrl, LoginUser, LoginPass, LoginFalse
    Dim ListStr, LsString, LoString, ListPaingType, LPsString, LPoString, ListPaingStr1, ListPaingStr2, ListPaingID1, ListPaingID2, ListPaingStr3, HsString, HoString, HttpUrlType, HttpUrlStr
    Dim TsString, ToString, CsString, CoString, AuthorType, AsString, AoString, AuthorStr, CopyFromType, FsString, FoString, CopyFromStr, KeyType, KsString, KoString, KeyStr, KeyScatterNum, NewsPaingType, NPsString, NPoString, NewsPaingStr1, NewsPaingStr2
    Dim PsString, PoString, PhsString, PhoString
    Dim IsString, IoString, IntroType, IntroStr, IntroNum
    Dim WebUrl, ListUrl, ListCode, NewsArrayCode, NewsArray, UrlTest, Testi, testUrl
    Dim DateType, DsString, DoString
    Dim IsField, Field, i, iField, iFieldNum
    Dim arrField, arrField2, FieldID, FieldName, FieldType, FisSting, FioSting, FieldStr
    '�б�����ͼ
    Dim ThumbnailType, ThsString, ThoString
    Dim ThumbnailArrayCode, ThumbnailArray, ThumbnailUrl

    testUrl = Trim(Request("testUrl"))
    FoundErr = False

    If ItemID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ���õĲɼ���Ŀ</li>"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If


    '����
    If NeedSave = "True" Then
        '�б����
        ListStr = Trim(Request.Form("ListStr"))
        LsString = Request.Form("LsString")
        LoString = Request.Form("LoString")
        ListPaingType = Request.Form("ListPaingType")
        LPsString = Request.Form("LPsString")
        LPoString = Request.Form("LPoString")
        ListPaingStr1 = Trim(Request.Form("ListPaingStr1"))
        ListPaingStr2 = Trim(Request.Form("ListPaingStr2"))
        ListPaingID1 = Request.Form("ListPaingID1")
        ListPaingID2 = Request.Form("ListPaingID2")
        ListPaingStr3 = Request.Form("ListPaingStr3")
        '���ӱ���
        HsString = Request.Form("HsString")
        HoString = Request.Form("HoString")
        HttpUrlType = Trim(Request.Form("HttpUrlType"))
        HttpUrlStr = Trim(Request.Form("HttpUrlStr"))
        '�б�����ͼ����
        ThumbnailType = PE_CLng(Trim(Request.Form("ThumbnailType")))
        ThsString = Trim(Request.Form("ThsString"))
        ThoString = Trim(Request.Form("ThoString"))

        If LsString = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�б�ʼ��ǲ���Ϊ��</li>"
        End If
        If LoString = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�б������ǲ���Ϊ��</li>"
        End If
        If ListPaingType = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ѡ���б�������ҳ����</li>"
        Else
            ListPaingType = CLng(ListPaingType)
            Select Case ListPaingType '�����б�,�ж��б�����
                Case 0, 1 '0 �޷�ҳ,1 �����ҳ
                    If ListStr = "" Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>�б�����ҳ����Ϊ��</li>"
                    Else
                        ListStr = Trim(ListStr)
                    End If
                    If ListPaingType = 1 Then
                        If LPsString = "" Or LPoString = "" Then
                            FoundErr = True
                            ErrMsg = ErrMsg & "<li>������ҳ��ʼ/������ǲ���Ϊ��</li>"
                        End If
                        'If ListPaingStr1 <> "" And Len(ListPaingStr1) < 15 Then
                        '    FoundErr = True
                        '    ErrMsg = ErrMsg & "<li>������ҳ�ض������ò���ȷ(����15���ַ�)</li>"
                        'End If
                    End If
                Case 2 '�������ַ�ҳ
                    If ListPaingStr2 = "" Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>���������ַ�����Ϊ��</li>"
                    End If
                    If IsNumeric(ListPaingID1) = False Or IsNumeric(ListPaingID2) = False Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>�������ɵķ�Χֻ��������</li>"
                    Else
                        ListPaingID1 = PE_CLng(ListPaingID1)
                        ListPaingID2 = PE_CLng(ListPaingID2)
                        ListPaingID1 = PE_CLng(ListPaingID1)
                        ListPaingID2 = PE_CLng(ListPaingID2)
                        If ListPaingID1 = 0 And ListPaingID2 = 0 Then
                            FoundErr = True
                            ErrMsg = ErrMsg & "<li>�������ɷ�Χ���ò���ȷ</li>"
                        End If
                    End If
                Case 3 '�ֹ���ҳ
                    If ListPaingStr3 = "" Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>�б�������ҳ����Ϊ��,���ֶ����</li>"
                    Else
                        ListPaingStr3 = ListPaingStr3 'Replace(ListPaingStr3, Chr(13), vbCrLf)
                    End If
                Case Else
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>��ѡ���б�������ҳ����</li>"
            End Select
        End If

        If HsString = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���ӿ�ʼ��ǲ���Ϊ��</li>"
        End If

        If HoString = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���ӽ�����ǲ���Ϊ��</li>"
        End If

        If HttpUrlType = 1 Then
            If HttpUrlStr = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�������ض���URL</li>"
            Else
                If Len(HttpUrlStr) < 15 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�������ӵ�ַ���ò���ȷ(����15���ַ�)</li>"
                End If
            End If
        End If

        If ThumbnailType = 1 Then
            If ThsString = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�б�����ͼ��ʼ��ǲ���Ϊ��</li>"
            End If

            If ThoString = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�б�����ͼ������ǲ���Ϊ��</li>"
            End If
        End If

        If FoundErr <> True Then
            sql = "Select * from PE_Item Where ItemID=" & ItemID
            Set rsItem = Server.CreateObject("adodb.recordset")
            rsItem.Open sql, Conn, 1, 3

            '�����б�
            rsItem("LsString") = LsString
            rsItem("LoString") = LoString
            rsItem("ListPaingType") = ListPaingType
            Select Case ListPaingType
                Case 0, 1
                    rsItem("ListStr") = ListStr
                    If ListPaingType = 1 Then
                        rsItem("LPsString") = LPsString
                        rsItem("LPoString") = LPoString
                        rsItem("ListPaingStr1") = Trim(ListPaingStr1)
                    End If
                Case 2
                    rsItem("ListPaingStr2") = ListPaingStr2
                    rsItem("ListPaingID1") = ListPaingID1
                    rsItem("ListPaingID2") = ListPaingID2
                Case 3
                    rsItem("ListPaingStr3") = ListPaingStr3
            End Select
            '��������
            rsItem("HsString") = HsString
            rsItem("HoString") = HoString
            rsItem("HttpUrlType") = HttpUrlType
            '�����б�����ͼ
            rsItem("ThumbnailType") = ThumbnailType
            rsItem("ThsString") = ThsString
            rsItem("ThoString") = ThoString
            If HttpUrlType = 1 Then
                rsItem("HttpUrlStr") = HttpUrlStr
            End If
            rsItem.Update
            rsItem.Close
            Set rsItem = Nothing
        End If
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    sql = "Select * from PE_Item Where ItemID=" & ItemID
    Set rsItem = Server.CreateObject("adodb.recordset")
    rsItem.Open sql, Conn, 1, 1

    If rsItem.EOF And rsItem.BOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���Ĳɼ���Ŀ</li>"
    Else
        WebUrl = rsItem("WebUrl")
        
        LoginType = rsItem("LoginType")
        LoginUrl = rsItem("LoginUrl")
        LoginPostUrl = rsItem("LoginPostUrl")
        LoginUser = rsItem("LoginUser")
        LoginPass = rsItem("LoginPass")
        LoginFalse = rsItem("LoginFalse")
      
        ListStr = rsItem("ListStr")
        LsString = rsItem("LsString")
        LoString = rsItem("LoString")
        ListPaingType = rsItem("ListPaingType")
        LPsString = rsItem("LPsString")
        LPoString = rsItem("LPoString")
        ListPaingStr1 = rsItem("ListPaingStr1")
        ListPaingStr2 = rsItem("ListPaingStr2")
        ListPaingID1 = rsItem("ListPaingID1")
        ListPaingID2 = rsItem("ListPaingID2")
        ListPaingStr3 = rsItem("ListPaingStr3")
        HsString = rsItem("HsString")
        HoString = rsItem("HoString")
        HttpUrlType = rsItem("HttpUrlType")
        HttpUrlStr = rsItem("HttpUrlStr")

        ThumbnailType = rsItem("ThumbnailType")
        ThsString = rsItem("ThsString")
        ThoString = rsItem("ThoString")

        TsString = rsItem("TsString")
        ToString = rsItem("ToString")
        CsString = rsItem("CsString")
        CoString = rsItem("CoString")
        AuthorType = rsItem("AuthorType")
        AsString = rsItem("AsString")
        AoString = rsItem("AoString")
        AuthorStr = rsItem("AuthorStr")
        DateType = PE_CLng(rsItem("DateType"))
        DsString = rsItem("DsString")

        DoString = rsItem("DoString")
    
        CopyFromType = rsItem("CopyFromType")
        FsString = rsItem("FsString")
        FoString = rsItem("FoString")
        CopyFromStr = rsItem("CopyFromStr")
        
        KeyType = rsItem("KeyType")
        KsString = rsItem("KsString")
        KoString = rsItem("KoString")
        KeyStr = rsItem("KeyStr")
        KeyScatterNum = rsItem("KeyScatterNum")
        
        NewsPaingType = rsItem("NewsPaingType")
        NPsString = rsItem("NPsString")
        NPoString = rsItem("NPoString")
        NewsPaingStr1 = rsItem("NewsPaingStr1")
        NewsPaingStr2 = rsItem("NewsPaingStr2")

        PsString = rsItem("PsString")
        PoString = rsItem("PoString")
        PhsString = rsItem("PhsString")
        PhoString = rsItem("PhoString")

        IsString = rsItem("IsString")
        IoString = rsItem("IoString")
        IntroType = PE_CLng(rsItem("IntroType"))
        IntroStr = rsItem("IntroStr")
        IntroNum = rsItem("IntroNum")
        IsField = rsItem("IsField")
        Field = rsItem("Field")
    End If

    rsItem.Close
    Set rsItem = Nothing
        
    If LsString = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�б�ʼ��ǲ���Ϊ�գ��޷�����,�뷵����һ���������ã�</li>"
    End If

    If LoString = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�б������ǲ���Ϊ�գ��޷�����,�뷵����һ���������ã�</li>"
    End If

    If ListPaingType = 0 Or ListPaingType = 1 Then
        If ListStr = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�б�����ҳ����Ϊ�գ��޷�����,�뷵����һ���������ã�</li>"
        End If

        If ListPaingType = 1 Then
            If LPsString = "" Or LPoString = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>������ҳ��ʼ��������ǲ���Ϊ�գ��޷�����,�뷵����һ���������ã�</li>"
            End If
        End If

        If ListPaingStr1 <> "" And Len(ListPaingStr1) < 15 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>������ҳ�ض������ò���ȷ���޷�����,�뷵����һ���������ã�</li>"
        End If

    ElseIf ListPaingType = 2 Then

        If ListPaingStr2 = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��������ԭ�ַ�������Ϊ�գ��޷�����,�뷵����һ����������</li>"
        End If

        If IsNumeric(ListPaingID1) = False Or IsNumeric(ListPaingID2) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�������ɵķ�Χֻ�������֣��޷�����,�뷵����һ����������</li>"
        Else

            If ListPaingID1 = 0 And ListPaingID2 = 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�������ɵķ�Χ����ȷ���޷�����,�뷵����һ����������</li>"
            End If
        End If

    ElseIf ListPaingType = 3 Then

        If ListPaingStr3 = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>������ҳ����Ϊ�գ��޷�����,�뷵����һ����������</li>"
        End If

    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ѡ�񷵻���һ������������ҳ����</li>"
    End If

    If ThumbnailType = 1 Then
        If ThsString = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�б�����ͼ��ʼ��ǲ���Ϊ�գ��޷�����,�뷵����һ����������</li>"
        End If

        If ThoString = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�б�����ͼ������ǲ���Ϊ�գ��޷�����,�뷵����һ����������</li>"
        End If
    End If

    If LoginType = 1 Then
        If LoginUrl = "" Or LoginPostUrl = "" Or LoginUser = "" Or LoginPass = "" Or LoginFalse = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�뽫��¼��Ϣ��д����</li>"
        Else
            LoginData = UrlEncoding(LoginUser & "&" & LoginPass)
            LoginResult = PostHttpPage(LoginUrl, LoginPostUrl, LoginData, PE_CLng(WebUrl))
            If InStr(LoginResult, LoginFalse) > 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>��¼��վʱ����������ȷ�ϵ�¼��Ϣ����ȷ�ԣ�</li>"
            End If
        End If
    End If
 
    If FoundErr <> True Then

        Select Case ListPaingType

            Case 0, 1
                ListUrl = ListStr

            Case 2
                ListUrl = Replace(ListPaingStr2, "{$ID}", CStr(ListPaingID1))

            Case 3

                If InStr(ListPaingStr3, vbCrLf) > 0 Then
                    ListUrl = Left(ListPaingStr3, InStr(ListPaingStr3, vbCrLf) - 1)
                Else
                    ListUrl = ListPaingStr3
                End If

        End Select

    End If
              
    If FoundErr <> True Then
        ListCode = GetHttpPage(ListUrl, PE_CLng(WebUrl)) '��Դ����

        If ListCode <> "$False$" Then
            ListCode = GetBody(ListCode, LsString, LoString, False, False)

            If ListCode = "$False$" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�ڽ�ȡ�б�ʱ��������</li>"
            End If

        Else
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�ڻ�ȡ��" & ListUrl & "��ҳԴ��ʱ��������</li>"
        End If
    End If

    If FoundErr <> True Then
        NewsArrayCode = GetArray(ListCode, HsString, HoString, False, False) '�������

        If NewsArrayCode = "$False$" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�ڷ�����" & ListUrl & "�����б�ʱ��������</li>"
        Else
            NewsArray = Split(NewsArrayCode, "$Array$")
            If IsArray(NewsArray) = True Then
                If HttpUrlType = 1 Then
                    If testUrl <> "" Then
                        UrlTest = Replace(HttpUrlStr, "{$ID}", NewsArray(PE_CLng(testUrl)))
                    Else
                        UrlTest = Replace(HttpUrlStr, "{$ID}", NewsArray(0))
                    End If
                Else
                    If testUrl <> "" Then
                        UrlTest = DefiniteUrl(NewsArray(PE_CLng(testUrl)), ListUrl)
                    Else
                        UrlTest = DefiniteUrl(NewsArray(0), ListUrl)
                    End If
                End If

                If InStr(UrlTest, "/?") > 0 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�б�������/?������������·���볢�����⴦��</li>"
                End If
            Else
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�ڷ�����" & ListUrl & "�����б�ʱ��������</li>"
            End If
        End If
        If UrlTest = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���ܵõ���ȷ������ҳURL</li>"
        End If
    End If

    If FoundErr <> True And ThumbnailType = 1 Then
        ThumbnailArrayCode = GetArray(ListCode, ThsString, ThoString, False, False) '����б�����ͼ

        If ThumbnailArrayCode = "$False$" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�ڷ�����" & ListUrl & "�б�����ͼʱ��������</li>"
        Else
            ThumbnailArray = Split(ThumbnailArrayCode, "$Array$")

            If IsArray(NewsArray) = True Then
                If testUrl <> "" Then
                    ThumbnailUrl = DefiniteUrl(ThumbnailArray(PE_CLng(testUrl)), ListUrl)
                Else
                    ThumbnailUrl = DefiniteUrl(ThumbnailArray(0), ListUrl)
                End If
            Else
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�ڷ�����" & ListUrl & "�б�����ͼʱ��������</li>"
            End If
        End If
        If ThumbnailUrl = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���ܵõ���ȷ���б�����ͼ</li>"
        End If
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If
    
    Call ShowChekcFormVbs

    Response.Write "<script language=""JavaScript"">" & vbCrLf
    Response.Write "<!--" & vbCrLf
    Response.Write "function IsDigit(){" & vbCrLf
    Response.Write "  return ((event.keyCode >= 48) && (event.keyCode <= 57));" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<br><br>" & vbCrLf
    Response.Write "<font color='red'>���Ƿ��������õ�������ҳURL�б�</font>" & vbCrLf
    Response.Write "<form method=""post"" action=""Admin_CollectionManage.asp"" name=""form1"">" & vbCrLf
    Response.Write "<select name=""testUrl"" onchange=""javascript:window.location='Admin_CollectionManage.asp?Action=Step3&ItemID=" & ItemID & "&CollectionModify=true&testUrl='+this.options[this.selectedIndex].value;"">" & vbCrLf
    For Testi = 0 To UBound(NewsArray)
        Response.Write "<option value='" & Testi & "'"
        If NewsArray(Testi) = UrlTest Then Response.Write " selected"
        Response.Write ">" & DefiniteUrl(NewsArray(Testi), ListUrl) & "</option>" & vbCrLf
    Next
    Response.Write "</select>" & vbCrLf
    Response.Write "&nbsp;&nbsp; <a href='" & DefiniteUrl(UrlTest, ListUrl) & "' target='_blank'><font color=red>�����Ƿ���ȷ</a>��</font><br><br>"
    If ThumbnailType = 1 Then
        Response.Write "<font color='red'>���Ƿ��������õ�������ͼ��</font><IMG SRC=" & ThumbnailUrl & "  align='�����Ƿ���ȷ' width='130' height='90'  BORDER='0'><br>" & vbCrLf
    End If

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr align='center' height='24'>"
    Response.Write "<td id='TabTitle' class='title6' onclick='ShowTabs(0)'>��������</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(1)'>ѡ������</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(2)'>�Զ�������</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(3)'>��ҳ����</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(4)'>����Ԥ��</td>" & vbCrLf
    Response.Write "   <td id='TabTitle' class='title5' onclick=""ShowTabs(5):setFileFields('" & UrlTest & "')"";'>��ҳԤ��</td>" & vbCrLf
    Response.Write "<td>&nbsp;</td></tr></table>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr align='left' class='tdbg'><td width='5'></td>"
    Response.Write "    <td class='tdbg' valign='top' >"
    Response.Write "      <table width='720' border='0' align='left' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>���⿪ʼ���룺<br><br><br><br><br><br>" & vbCrLf
    Response.Write "            ����������룺</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" width=620>" & vbCrLf
    Response.Write "            <textarea name=""TsString"" style='width:450px;height:100px'>"
    If Trim(TsString) <> "" Then Response.Write Server.HTMLEncode(TsString & "")
    Response.Write "</textarea>&nbsp;<FONT color='red'>*</FONT><input TYPE='button' value='���Դ���' onCLICK='ceshi(5)' ><br>" & vbCrLf
    Response.Write "            <textarea name=""ToString"" style='width:450px;height:100px'>"
    If Trim(ToString) <> "" Then Response.Write Server.HTMLEncode(ToString & "")
    Response.Write "</textarea>&nbsp;<FONT color='red'>*</FONT><input TYPE='button' value='���Դ���' onCLICK='ceshi(6)' ></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>���Ŀ�ʼ���룺<p>��</p><p>��</p>" & vbCrLf
    Response.Write "            ���Ľ������룺</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <textarea name=""CsString"" style='width:450px;height:100px'>"
    If Trim(CsString) <> "" Then Response.Write Server.HTMLEncode(CsString & "")
    Response.Write "</textarea>&nbsp;<FONT color='red'>*</FONT><input TYPE='button' value='���Դ���' onCLICK='ceshi(7)' ><br>" & vbCrLf
    Response.Write "            <textarea name=""CoString"" style='width:450px;height:100px'>"
    If Trim(CoString) <> "" Then Response.Write Server.HTMLEncode(CoString & "")
    Response.Write "</textarea>&nbsp;<FONT color='red'>*</FONT><input TYPE='button' value='���Դ���' onCLICK='ceshi(8)' ></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>&nbsp;&nbsp;&nbsp;����ʱ�䣺</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <input type=""radio"" value=""0"" name=""DateType"""
    If DateType = 0 Then Response.Write "checked"
    Response.Write " onClick=""javascript:Date1.style.display='none'"">�ɼ�ʱ��ϵͳʱ��&nbsp;" & vbCrLf
    Response.Write "      <input type=""radio"" value=""1"" name=""DateType"""
    If DateType = 1 Then Response.Write "checked"
    Response.Write " onClick=""javascript:Date1.style.display=''"">��Դ�����л�ȡʱ��&nbsp;</td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg"" id=""Date1"" style=""display:'"
    If DateType <> 1 Then Response.Write "none"
    Response.Write "'""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>ʱ�俪ʼ���룺</font><br><br><br><br><br><br>" & vbCrLf
    Response.Write "            ʱ��������룺</font></td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <textarea name=""DsString"" style='width:450px;height:100px'>"
    If Trim(DsString) <> "" Then Response.Write Server.HTMLEncode(DsString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(17)' ><br>" & vbCrLf
    Response.Write "            <textarea name=""DoString"" style='width:450px;height:100px'>"
    If Trim(DoString) <> "" Then Response.Write Server.HTMLEncode(DoString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(18)' ></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>&nbsp;&nbsp;&nbsp;�������ߣ�</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <input type=""radio"" value=""0"" name=""AuthorType"""
    If AuthorType = 0 Then Response.Write "checked"
    Response.Write " onClick=""javascript:Author1.style.display='none'"">ָ��Ϊ��������&nbsp;" & vbCrLf
    Response.Write "            <input type=""radio"" value=""1"" name=""AuthorType"""
    If AuthorType = 1 Then Response.Write "checked"
    Response.Write " onClick=""javascript:Author1.style.display=''"">��Դ�����л�ȡ����&nbsp;" & vbCrLf
    Response.Write "            <input type=""radio"" value=""2"" name=""AuthorType"""
    If AuthorType = 2 Then Response.Write "checked"
    Response.Write " onClick=""javascript:Author1.style.display='none'"">ָ��Ϊ<input name=""AuthorStr"" type=""text"" id=""AuthorStr"" value=""" & AuthorStr & """></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg"" id=""Author1"" style=""display:'"
    If AuthorType <> 1 Then Response.Write "none"
    Response.Write "'""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>���߿�ʼ���룺</font><br><br><br><br><br><br>" & vbCrLf
    Response.Write "            ���߽������룺</font></td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <textarea name=""AsString"" style='width:450px;height:100px'>"
    If Trim(AsString) <> "" Then Response.Write Server.HTMLEncode(AsString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(9)' ><br>" & vbCrLf
    Response.Write "            <textarea name=""AoString"" style='width:450px;height:100px'>"
    If Trim(AoString) <> "" Then Response.Write Server.HTMLEncode(AoString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(10)' ></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>&nbsp;&nbsp;&nbsp;������Դ��</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <input type=""radio"" value=""0"" name=""CopyFromType"""
    If CopyFromType = 0 Then Response.Write "checked"
    Response.Write " onClick=""javascript:CopyFrom1.style.display='none'"">ָ��Ϊ�����ꡱ&nbsp;" & vbCrLf
    Response.Write "            <input type=""radio"" value=""1"" name=""CopyFromType"""
    If CopyFromType = 1 Then Response.Write "checked"
    Response.Write " onClick=""javascript:CopyFrom1.style.display=''"">��Դ�����л�ȡ��Դ&nbsp;" & vbCrLf
    Response.Write "            <input type=""radio"" value=""2"" name=""CopyFromType"""
    If CopyFromType = 2 Then Response.Write "checked"
    Response.Write " onClick=""javascript:CopyFrom1.style.display='none'"">ָ��Ϊ<input name=""CopyFromStr"" type=""text"" id=""CopyFromStr"" value=""" & CopyFromStr & """></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg"" id=""CopyFrom1"" style=""display:'"
    If CopyFromType <> 1 Then Response.Write "none"
    Response.Write "'""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>��Դ��ʼ���룺</font><br><br><br><br><br><br>" & vbCrLf
    Response.Write "            ��Դ�������룺</font></td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <textarea name=""FsString"" style='width:450px;height:100px'>"
    If Trim(FsString) <> "" Then Response.Write Server.HTMLEncode(FsString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(11)' ><br>" & vbCrLf
    Response.Write "            <textarea name=""FoString"" style='width:450px;height:100px'>"
    If Trim(FoString) <> "" Then Response.Write Server.HTMLEncode(FoString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(12)' ></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>���¹ؼ��֣�</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <input type=""radio"" value=""0"" name=""KeyType"""
    If KeyType = 0 Then Response.Write "checked"
    Response.Write " onClick=""javascript:Key1.style.display='none';Key2.style.display='none'"">�������ɢ��Ϊ�ؼ���&nbsp;�ؼ��ʳ��ȣ�<input type=""text"" name=""KeyScatterNum"" value=""" & KeyScatterNum & """ maxlength=""1"" size=""1"" ONKEYPRESS=""javascript:event.returnValue=IsDigit()""> �ַ�" & vbCrLf
    Response.Write "            <input type=""radio"" value=""1"" name=""KeyType"""
    If KeyType = 1 Then Response.Write "checked"
    Response.Write " onClick=""javascript:Key1.style.display='';Key2.style.display='none'"">��Դ�����л�ȡ�ؼ���&nbsp;" & vbCrLf
    Response.Write "            <input type=""radio"" value=""2"" name=""KeyType"""
    If KeyType = 2 Then Response.Write "checked"
    Response.Write " onClick=""javascript:Key1.style.display='none';Key2.style.display=''"">ָ���ؼ���&nbsp;" & vbCrLf
    Response.Write " ������</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg"" id=""Key1"" style=""display:'"
    If KeyType <> 1 Then Response.Write "none"
    Response.Write "'""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>�ؼ��ֿ�ʼ���룺</font><br><br><br><br><br><br>" & vbCrLf
    Response.Write "            �ؼ��ֽ������룺</font></td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <textarea name=""KsString"" style='width:450px;height:100px'>"
    If Trim(KsString) <> "" Then Response.Write Server.HTMLEncode(KsString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(13)' ><br>" & vbCrLf
    Response.Write "            <textarea name=""KoString"" style='width:450px;height:100px'>"
    If Trim(KoString) <> "" Then Response.Write Server.HTMLEncode(KoString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(14)' ></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg"" id=""Key2"" style=""display:'"
    If KeyType <> 2 Then Response.Write "none"
    Response.Write "'""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>ָ��Ϊ��</font></td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <input name=""KeyStr"" type=""text"" id=""KeyStr"" value=""" & KeyStr & """>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>���¼�飺</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <input type=""radio"" value=""0"" name=""IntroType"""
    If IntroType = 0 Then Response.Write "checked"
    Response.Write " onClick=""javascript:Intro1.style.display='none';Intro2.style.display='none';Intro3.style.display='none'"">��¼��&nbsp;" & vbCrLf
    Response.Write "            <input type=""radio"" value=""1"" name=""IntroType"""
    If IntroType = 1 Then Response.Write "checked"
    Response.Write " onClick=""javascript:Intro1.style.display='';Intro2.style.display='none';Intro3.style.display='none'"">��Դ�����л�ȡ���&nbsp;" & vbCrLf
    Response.Write "            <input type=""radio"" value=""3"" name=""IntroType"""
    If IntroType = 3 Then Response.Write "checked"
    Response.Write " onClick=""javascript:Intro1.style.display='none';Intro2.style.display='none';Intro3.style.display=''"">ָ������ǰ�����ַ�Ϊ���" & vbCrLf
    Response.Write "            <input type=""radio"" value=""2"" name=""IntroType"""
    If IntroType = 2 Then Response.Write "checked"
    Response.Write " onClick=""javascript:Intro1.style.display='none';Intro2.style.display='';Intro3.style.display='none'"">ָ���������</td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg"" id=""Intro1"" style=""display:'"
    If IntroType <> 1 Then Response.Write "none"
    Response.Write "'""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>��鿪ʼ���룺</font><br><br><br><br><br><br>" & vbCrLf
    Response.Write "            ���������룺</font></td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <textarea name=""IsString"" style='width:450px;height:100px'>"
    If Trim(IsString) <> "" Then Response.Write Server.HTMLEncode(IsString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(21)' ><br>" & vbCrLf
    Response.Write "            <textarea name=""IoString"" style='width:450px;height:100px'>"
    If Trim(IoString) <> "" Then Response.Write Server.HTMLEncode(IoString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(22)' ></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg"" id=""Intro2"" style=""display:'"
    If IntroType <> 2 Then Response.Write "none"
    Response.Write "'""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>��ָ����飺</font></td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <TEXTAREA NAME='IntroStr' ROWS='' COLS='' style='width:450px;height:100px'>" & IntroStr & "</TEXTAREA>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg"" id=""Intro3"" style=""display:'"
    If IntroType <> 3 Then Response.Write "none"
    Response.Write "'""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>�ַ�����</font></td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <input name=""IntroNum"" type=""text"" id=""IntroNum"" value=""" & IntroNum & """>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>&nbsp;&nbsp;&nbsp;�Զ����ֶ����ã�</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" width='620'>" & vbCrLf
    Response.Write "            <input type=""radio"" value=""0"" name=""IsField"""
    If IsField = 0 Then Response.Write "checked"
    Response.Write " onClick=""javascript:Field.style.display='none'"">������&nbsp;" & vbCrLf
    Response.Write "            <input type=""radio"" value=""1"" name=""IsField"""
    If IsField = 1 Then Response.Write "checked"
    Response.Write " onClick=""javascript:Field.style.display=''"">����&nbsp;" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "       <tr class=""tdbg"" id=""Field"" style=""display:'"
    If IsField <> 1 Then Response.Write "none"
    Response.Write "'"">" & vbCrLf
    Response.Write "          <td class=""tdbg5"" colspan='2'>" & vbCrLf
    Response.Write "            <table border='0' cellSpacing=1 cellPadding=1 width='100%' height='100%' align='right' bgcolor='#FFFFFF'>" & vbCrLf
        
    sql = "select FieldID,FieldName,Title from PE_Field where (FieldType=1 or FieldType=2) and (ChannelID=-1 or ChannelID in (select ChannelID from PE_Channel where ModuleType=1 and Disabled=" & PE_False & ")) Order by FieldID desc"
    
    Set rs = Conn.Execute(sql)

    If rs.BOF And rs.EOF Then
        Response.Write "û���κ��Զ����ֶΣ�"
    Else
        Do While Not rs.EOF
            i = i + 1

            If InStr(Field, "|||") > 0 Then
                arrField = Split(Field, "|||")

                For iField = 0 To UBound(arrField)
                    arrField2 = Split(arrField(iField), "{#F}")

                    If PE_CLng(arrField2(0)) = rs("FieldID") Then
                        FieldType = PE_CLng(arrField2(2))
                        FisSting = arrField2(3)
                        FioSting = arrField2(4)
                        FieldStr = arrField2(5)
                        Exit For
                    End If

                Next

            Else

                If InStr(Field, "{#F}") > 0 Then
                    arrField2 = Split(Field, "{#F}")

                    If PE_CLng(arrField2(0)) = rs("FieldID") Then
                        FieldType = PE_CLng(arrField2(2))
                        FisSting = arrField2(3)
                        FioSting = arrField2(4)
                        FieldStr = arrField2(5)
                    End If
                End If
            End If

            Response.Write "    <tr class=""tdbg""> " & vbCrLf
            Response.Write "      <td width=""120"" class=""tdbg5"" align='right'>&nbsp;&nbsp;&nbsp;" & rs("Title") & "���ã�</td>" & vbCrLf
            Response.Write "      <td class=""tdbg"" >" & vbCrLf
            Response.Write "        <input type=""radio"" value=""0"" name=""FieldType" & i & """ "

            If FieldType = 0 Then Response.Write "checked"
            Response.Write " onClick=""javascript:FieldA" & i & ".style.display='none';FieldB" & i & ".style.display='none'"">��¼��&nbsp;" & vbCrLf
            Response.Write "        <input type=""radio"" value=""1"" name=""FieldType" & i & """ "

            If FieldType = 1 Then Response.Write "checked"
            Response.Write " onClick=""javascript:FieldA" & i & ".style.display='';FieldB" & i & ".style.display='none'"">��Դ�����л�ȡ" & rs("Title") & "����&nbsp;" & vbCrLf
            Response.Write "        <input type=""radio"" value=""2"" name=""FieldType" & i & """ "

            If FieldType = 2 Then Response.Write "checked"
            Response.Write " onClick=""javascript:FieldA" & i & ".style.display='none';FieldB" & i & ".style.display=''"">ָ��" & rs("Title") & "</td>" & vbCrLf
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class=""tdbg"" id=""FieldA" & i & """ style=""display:'"

            If FieldType <> 1 Then Response.Write "none"
            Response.Write "'""> " & vbCrLf
            Response.Write "      <td width=""120"" class=""tdbg5"" align='right'>" & rs("Title") & "��ʼ���룺</font><br><br><br><br><br><br>" & vbCrLf
            Response.Write "      " & rs("Title") & "�������룺</font></td>" & vbCrLf
            Response.Write "      <td class=""tdbg"" >" & vbCrLf
            Response.Write "        <textarea name=""FisSting" & i & """ style='width:450px;height:100px'>"

            If Trim(FisSting) <> "" Then Response.Write Server.HTMLEncode(FisSting & "")
            Response.Write "</textarea><br>" & vbCrLf
            Response.Write "        <textarea name=""FioSting" & i & """ style='width:450px;height:100px'>"

            If Trim(FioSting) <> "" Then Response.Write Server.HTMLEncode(FioSting & "")
            Response.Write "</textarea></td>" & vbCrLf
            Response.Write "    </tr>" & vbCrLf
            Response.Write "    <tr class=""tdbg"" id=""FieldB" & i & """ style=""display:'"

            If FieldType <> 2 Then Response.Write "none"
            Response.Write "'""> " & vbCrLf
            Response.Write "      <td width=""120"" class=""tdbg5"" align='right'>��ָ��" & rs("FieldName") & "��</font></td>" & vbCrLf
            Response.Write "      <td class=""tdbg"" >" & vbCrLf
            Response.Write "        <input name=""FieldStr" & i & """ type=""text""  value=""" & FieldStr & """>" & vbCrLf
            Response.Write "      </td>" & vbCrLf
            Response.Write "      <INPUT TYPE='hidden' name='FieldID" & i & "' value='" & rs("FieldID") & "'>" & vbCrLf
            Response.Write "      <INPUT TYPE='hidden' name='FieldName" & i & "' value='" & rs("FieldName") & "'>" & vbCrLf
            Response.Write "    </tr>" & vbCrLf

            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing
    Response.Write "              <INPUT TYPE='hidden' name='iFieldNum' value='" & i & "'>" & vbCrLf
    Response.Write "            </table>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>���ķ�ҳ���ã�</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <input type=""radio"" value=""0"" name=""NewsPaingType"""
    If NewsPaingType = 0 Then Response.Write "checked"
    Response.Write " onClick=""javascript:NewsPaing1.style.display='none';NewsPaing2.style.display='none'"">���ɼ����ķ�ҳ&nbsp;" & vbCrLf
    Response.Write "            <input type=""radio"" value=""1"" name=""NewsPaingType"""
    If NewsPaingType = 1 Then Response.Write "checked"
    Response.Write " onClick=""javascript:NewsPaing1.style.display='';NewsPaing2.style.display='none'"">��Դ�����л�ȡ����һҳ��URL&nbsp;" & vbCrLf
    Response.Write "            <input type=""radio"" value=""2"" name=""NewsPaingType"""
    If NewsPaingType = 2 Then Response.Write "checked"
    Response.Write " onClick=""javascript:NewsPaing1.style.display='none';NewsPaing2.style.display=''"">��Դ�����л�ȡ��ҳURL" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg"" id=""NewsPaing1"" style=""display:'"
    If NewsPaingType <> 1 Then Response.Write "none"
    Response.Write "'"">" & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>���ķ�ҳ����һҳ��<br>URL��ʼ���룺</font><br><br><br><br><br><br>" & vbCrLf
    Response.Write "            ���ķ�ҳ����һҳ��<br>URL�������룺</font></td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <textarea name=""NPsString"" style='width:450px;height:100px'>"
    If Trim(NPsString) <> "" Then Response.Write Server.HTMLEncode(NPsString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(15)' ><br>" & vbCrLf
    Response.Write "            <textarea name=""NPoString"" style='width:450px;height:100px'>"
    If Trim(NPoString) <> "" Then Response.Write Server.HTMLEncode(NPoString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(16)' ></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "       <tr class=""tdbg"" id=""NewsPaing2"" style=""display:'"
    If NewsPaingType <> 2 Then Response.Write "none"
    Response.Write "'"">" & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align='right'>��ҳ���뿪ʼ��<br><br><br><br><br><br>" & vbCrLf
    Response.Write "            ��ҳ���������<br><br><br><br><br><br><br>��ҳURL��ʼ���룺<br>��ҳURL�������룺</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"" >" & vbCrLf
    Response.Write "            <textarea name=""PsString"" style='width:450px;height:100px'>"
    If Trim(PsString) <> "" Then Response.Write Server.HTMLEncode(PsString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(19)' ><br>" & vbCrLf
    Response.Write "            <textarea name=""PoString"" style='width:450px;height:100px'>"
    If Trim(PoString) <> "" Then Response.Write Server.HTMLEncode(PoString & "")
    Response.Write "</textarea>&nbsp;<input TYPE='button' value='���Դ���' onCLICK='ceshi(20)' ><br>" & vbCrLf
    Response.Write "           <input type=""text"" name=""PhsString"" size=""50"" maxlength=""200"" value=""" & Server.HTMLEncode(PhsString & "") & """><br>" & vbCrLf
    Response.Write "           <input type=""text"" name=""PhoString"" size=""50"" maxlength=""200"" value=""" & Server.HTMLEncode(PhoString & "") & """>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "         <td class=""tdbg"" >" & vbCrLf
    Response.Write "<TEXTAREA NAME='Content' style='width:785px;height:400px'>"
    Response.Write Server.HTMLEncode(GetHttpPage(UrlTest, PE_CLng(WebUrl)))
    Response.Write "</TEXTAREA>" & vbCrLf
    Response.Write "        </td>" & vbCrLf
    Response.Write "       </tr>" & vbCrLf
    Response.Write "       </tbody>" & vbCrLf
    Response.Write "       <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "         <td align='center' id='objFiles'></td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "     </tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    Response.Write "   </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write " </table>" & vbCrLf
    Response.Write " <br>" & vbCrLf
    Response.Write " <center>" & vbCrLf
    Response.Write "  <input name=""Action"" type=""hidden"" id=""Action"" value=""Step4"">" & vbCrLf
    Response.Write "  <input name=""NeedSave"" type=""hidden"" id=""NeedSave"" value=""True"">" & vbCrLf
    Response.Write "  <INPUT id=ItemID type=hidden value=" & ItemID & " name=ItemID>" & vbCrLf
    Response.Write "  <INPUT id=Cancel  onclick=""window.location.href='javascript:history.go(-1)'"" type=button value=""������һ��"" name=Cancel>&nbsp;&nbsp;" & vbCrLf
    Response.Write "  <INPUT type=submit value="" ��һ�� "" name=Submit onClick='CheckForm()'>" & vbCrLf
    Response.Write "  <input type=""hidden"" name=""UrlTest"" id=""UrlTest"" value=" & UrlTest & ">" & vbCrLf
    Response.Write " </center>" & vbCrLf
    Response.Write "</FORM>" & vbCrLf
    Response.Write "<br><b>ע�⣺</b><br>&nbsp;&nbsp;&nbsp;&nbsp;��ʼ������������<font color=red>������һ������ҳ����Ψһ��</font>�����ܱ�֤������ȷ�ɼ���������ݡ���Ϊÿ������ҳ�Ĵ��붼���ܲ�ͬ��������Ҫ�������������ҳ���ҵ���ͬ�Ŀ�ʼ����ͽ������룬���ܱ�֤���Դ���������ҳ��׼ȷ�ɼ����������ݡ�" & vbCrLf

    Call CloseConn
End Sub

'=================================================
'��������Step4
'��  �ã���������
'=================================================
Sub Step4()
    Dim LoginType, LoginUrl, LoginPostUrl, LoginUser, LoginPass, LoginFalse, LoginData, LoginResult
    Dim ListStr, LsString, LoString, ListPaingType, LPsString, LPoString, ListPaingStr1, ListPaingStr2, ListPaingID1, ListPaingID2, ListPaingStr3, HsString, HoString, HttpUrlType, HttpUrlStr
    Dim DateType, DsString, DoString, UpdateTime, UpDateType
    Dim TsString, ToString, CsString, CoString, AuthorType, AsString, AoString, AuthorStr, CopyFromType, FsString, FoString, CopyFromStr, KeyType, KsString, KoString, KeyStr, KeyScatterNum, NewsPaingType, NPsString, NPoString, NewsPaingStr1, NewsPaingStr2
    Dim PsString, PoString, PhsString, PhoString
    Dim IsString, IoString, IntroType, IntroStr, IntroNum, Intro
    Dim IsField, Field, i, iField, iFieldNum
    Dim arrField, arrField2, FieldID, FieldName, FieldType, FisSting, FioSting, FieldStr
    Dim UrlTest, ListUrl, ListCode
    Dim NewsUrl, NewsCode, NewsArrayCode, NewsArray
    Dim Title, Content, Author, CopyFrom, Key, ListPaingNext, NewsPaingNextCode, Testi
    Dim rsFilters, WebUrl
    Dim Arr_Filters, Filteri, FilterStr
    Dim ConversionTrails 'ת��·��
    '�ɼ����ķ�ҳ����
    Dim PageListCode, PageArrayCode, PageArray

    If ItemID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ���ԵĲɼ���Ŀ</li>"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    '����
    If NeedSave = "True" Then
        TsString = Request.Form("TsString")
        ToString = Request.Form("ToString")
        CsString = Request.Form("CsString")
        CoString = Request.Form("CoString")
        
        DateType = Trim(Request.Form("DateType"))
        DsString = Request.Form("DsString")

        DoString = Request.Form("DoString")

        AuthorType = Request.Form("AuthorType")
        AsString = Request.Form("AsString")
        AoString = Request.Form("AoString")
        AuthorStr = Request.Form("AuthorStr")

        CopyFromType = Request.Form("CopyFromType")
        FsString = Request.Form("FsString")
        FoString = Request.Form("FoString")
        CopyFromStr = Request.Form("CopyFromStr")

        KeyType = Request.Form("KeyType")
        KsString = Request.Form("KsString")
        KoString = Request.Form("KoString")
        KeyStr = Request.Form("KeyStr")
        KeyScatterNum = Request.Form("KeyScatterNum")

        NewsPaingType = Request.Form("NewsPaingType")
        NPsString = Request.Form("NpsString")
        NPoString = Request.Form("NpoString")
        NewsPaingStr1 = Request.Form("NewsPaingStr1")
        NewsPaingStr2 = Request.Form("NewsPaingStr2")

        PsString = Request.Form("PsString")
        PoString = Request.Form("PoString")
        PhsString = Request.Form("PhsString")
        PhoString = Request.Form("PhoString")

        IsString = Request.Form("IsString")
        IoString = Request.Form("IoString")
        IntroType = Request.Form("IntroType")
        IntroStr = Request.Form("IntroStr")
        IntroNum = Request.Form("IntroNum")

        UrlTest = Trim(Request.Form("UrlTest"))
        Testi = Trim(Request.Form("testUrl"))


        IsField = PE_CLng(Trim(Request.Form("IsField")))

        iFieldNum = PE_CLng(Trim(Request.Form("iFieldNum")))



        If iFieldNum >= 1 Then

            For i = 1 To iFieldNum
                FieldID = PE_CLng(Request.Form("FieldID" & i & ""))
                FieldName = Request.Form("FieldName" & i & "")
                FieldType = PE_CLng(Request.Form("FieldType" & i & ""))
                FisSting = Request.Form("FisSting" & i & "")
                FioSting = Request.Form("FioSting" & i & "")
                FieldStr = Request.Form("FieldStr" & i & "")

                If Field = "" Then
                    Field = FieldID & "{#F}" & FieldName & "{#F}" & FieldType & "{#F}" & FisSting & "{#F}" & FioSting & "{#F}" & FieldStr
                Else
                    Field = Field & "|||" & FieldID & "{#F}" & FieldName & "{#F}" & FieldType & "{#F}" & FisSting & "{#F}" & FioSting & "{#F}" & FieldStr
                End If

            Next

        End If
                
        If UrlTest = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�����������ݴ���ʱ��������</li>"
        Else
            NewsUrl = UrlTest
        End If

        If TsString = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���⿪ʼ��ǲ���Ϊ��</li>"
        End If

        If ToString = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���������ǲ���Ϊ��</li>"
        End If

        If CsString = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���Ŀ�ʼ��ǲ���Ϊ��</li>"
        End If

        If CoString = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���Ľ�����ǲ���Ϊ��</li>"
        End If
        
        If DateType = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�����ø���ʱ��Ĳɼ�ѡ��</li>"
        Else
            DateType = CLng(DateType)

            If DateType = 0 Then
            ElseIf DateType = 1 Then

                If DsString = "" Or DoString = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�뽫����ʱ��Ŀ�ʼ/���������д������</li>"
                End If

            Else
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�������������Ч���ӽ���</li>"
            End If
        End If

        If AuthorType = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�������������ߵĲɼ�ѡ��</li>"
        Else
            AuthorType = CLng(AuthorType)

            If AuthorType = 0 Then
            ElseIf AuthorType = 1 Then

                If AsString = "" Or AoString = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�뽫�������ߵĿ�ʼ/���������д������</li>"
                End If

            ElseIf AuthorType = 2 Then

                If AuthorStr = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>��ָ����������</li>"
                End If

            Else
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�������������Ч���ӽ���</li>"
            End If
        End If

        If CopyFromType = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>������������Դ�Ĳɼ�ѡ��</li>"
        Else
            CopyFromType = CLng(CopyFromType)

            If CopyFromType = 0 Then
            ElseIf CopyFromType = 1 Then

                If FsString = "" Or FoString = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�뽫������Դ�Ŀ�ʼ/���������д������</li>"
                End If

            ElseIf CopyFromType = 2 Then

                If CopyFromStr = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>��ָ��������Դ</li>"
                End If

            Else
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�������������Ч���ӽ���</li>"
            End If
        End If

        If KeyType = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�����ùؼ��ֵĲɼ�ѡ��</li>"
        Else
            KeyType = CLng(KeyType)

            If KeyType = 0 Then
                If PE_CLng(KeyScatterNum) = 0 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>��������Ч�Ĺؼ��ʳ���</li>"
                End If
            ElseIf KeyType = 1 Then
                If KsString = "" Or KoString = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�뽫�ؼ��ֵĿ�ʼ/���������д����</li>"
                End If

            ElseIf KeyType = 2 Then

                If KeyStr = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>��ָ���ؼ���</li>"
                End If
            Else
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�������������Ч���ӽ���</li>"
            End If
        End If


        If IntroType = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�����ü��Ĳɼ�ѡ��</li>"
        Else
            IntroType = PE_CLng(IntroType)

            If IntroType = 0 Then
            ElseIf IntroType = 1 Then

                If IsString = "" Or IoString = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�뽫���Ŀ�ʼ/���������д����</li>"
                End If

            ElseIf IntroType = 2 Then

                If IntroStr = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>��ָ�����</li>"
                End If

            ElseIf IntroType = 3 Then

                If IntroNum = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>��ָ��������ĵ�����</li>"
                End If

            Else
                FoundErr = True
                ErrMsg = ErrMsg & "<li>��������,�����Ч���ӽ���</li>"
            End If
        End If

        If IsField = 1 Then
            If Field = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>���趨�Զ����ֶεĲɼ�ѡ��</li>"
            End If
        End If

        If NewsPaingType = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���������ķ�ҳ�Ĳɼ�ѡ��</li>"
        Else
            NewsPaingType = CLng(NewsPaingType)

            If NewsPaingType = 0 Then
            ElseIf NewsPaingType = 1 Then

                If NPsString = "" Or NPoString = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�뽫���ķ�ҳ��һҳURL�Ŀ�ʼ/���������д����</li>"
                End If

                'If NewsPaingStr1 <> "" And Len(NewsPaingStr1) < 10 Then
                '    FoundErr = True
                '    ErrMsg = ErrMsg & "<li>���ķ�ҳ�����������ò���ȷ(����10���ַ�)</li>"
                'End If

            ElseIf NewsPaingType = 2 Then

                If PsString = "" Or PoString = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�뽫���ķ�ҳ����Ŀ�ʼ/���������д����</li>"
                End If

                If PhsString = "" Or PhoString = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�뽫���ķ�ҳURL�Ŀ�ʼ/���������д����</li>"
                End If

            Else
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�������������Ч���ӽ���</li>"
            End If
        End If

        If FoundErr <> True Then
            sql = "Select * from PE_Item Where ItemID=" & ItemID
            Set rsItem = Server.CreateObject("adodb.recordset")
            rsItem.Open sql, Conn, 1, 3
            
            LoginType = rsItem("LoginType")
            LoginUrl = rsItem("LoginUrl")
            LoginPostUrl = rsItem("LoginPostUrl")
            LoginUser = rsItem("LoginUser")
            LoginPass = rsItem("LoginPass")
            LoginFalse = rsItem("LoginFalse")
      
            rsItem("TsString") = TsString
            rsItem("ToString") = ToString
            rsItem("CsString") = CsString
            rsItem("CoString") = CoString
            
            rsItem("DateType") = DateType

            If DateType = 1 Then
                rsItem("DsString") = DsString
                rsItem("DoString") = DoString
            End If

            rsItem("AuthorType") = AuthorType

            If AuthorType = 1 Then
                rsItem("AsString") = AsString
                rsItem("AoString") = AoString
            ElseIf AuthorType = 2 Then
                rsItem("AuthorStr") = AuthorStr
            End If

            rsItem("CopyFromType") = CopyFromType

            If CopyFromType = 1 Then
                rsItem("FsString") = FsString
                rsItem("FoString") = FoString
            ElseIf CopyFromType = 2 Then
                rsItem("CopyFromStr") = CopyFromStr
            End If

            rsItem("KeyType") = KeyType

            If KeyType = 1 Then
                rsItem("KsString") = KsString
                rsItem("KoString") = KoString
            ElseIf KeyType = 2 Then
                rsItem("KeyStr") = KeyStr
            End If
            rsItem("KeyScatterNum") = KeyScatterNum

            rsItem("IntroType") = IntroType

            If IntroType = 1 Then
                rsItem("IsString") = IsString
                rsItem("IoString") = IoString
            ElseIf IntroType = 2 Then
                rsItem("IntroStr") = IntroStr
            ElseIf IntroType = 3 Then
                rsItem("IntroNum") = IntroNum
            End If

            rsItem("IsField") = IsField

            If IsField = 1 Then
                rsItem("Field") = Field
            End If

            rsItem("NewsPaingType") = NewsPaingType

            If NewsPaingType = 1 Then
                rsItem("NPsString") = NPsString
                rsItem("NPoString") = NPoString
                rsItem("NewsPaingStr1") = NewsPaingStr1
            ElseIf NewsPaingType = 2 Then
                rsItem("PsString") = PsString
                rsItem("PoString") = PoString
                rsItem("PhsString") = PhsString
                rsItem("PhoString") = PhoString
            ElseIf NewsPaingType = 3 Then
                rsItem("NewsPaingStr2") = NewsPaingStr2
            End If

            rsItem.Update
            rsItem.Close
            Set rsItem = Nothing
        End If
    End If

    If FoundErr <> True Then
        sql = "Select * from PE_Item Where ItemID=" & ItemID
        Set rsItem = Server.CreateObject("adodb.recordset")
        rsItem.Open sql, Conn, 1, 1

        If rsItem.EOF And rsItem.BOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���������Ҳ�������Ŀ</li>"
        Else
            WebUrl = rsItem("WebUrl")
            ListStr = rsItem("ListStr")
            LsString = rsItem("LsString")
            LoString = rsItem("LoString")
            ListPaingType = rsItem("ListPaingType")
            LPsString = rsItem("LPsString")
            LPoString = rsItem("LPoString")
            ListPaingStr1 = rsItem("ListPaingStr1")
            ListPaingStr2 = rsItem("ListPaingStr2")
            ListPaingID1 = rsItem("ListPaingID1")
            ListPaingID2 = rsItem("ListPaingID2")
            ListPaingStr3 = rsItem("ListPaingStr3")

            HsString = rsItem("HsString")
            HoString = rsItem("HoString")
            HttpUrlType = rsItem("HttpUrlType")
            HttpUrlStr = rsItem("HttpUrlStr")

            TsString = rsItem("TsString")
            ToString = rsItem("ToString")
            CsString = rsItem("CsString")
            CoString = rsItem("CoString")
            
            DateType = rsItem("DateType")
            DsString = rsItem("DsString")

            DoString = rsItem("DoString")
            UpDateType = rsItem("UpDateType")

            AuthorType = rsItem("AuthorType")
            AsString = rsItem("AsString")
            AoString = rsItem("AoString")
            AuthorStr = rsItem("AuthorStr")

            CopyFromType = rsItem("CopyFromType")
            FsString = rsItem("FsString")
            FoString = rsItem("FoString")
            CopyFromStr = rsItem("CopyFromStr")

            KeyType = rsItem("KeyType")
            KsString = rsItem("KsString")
            KoString = rsItem("KoString")
            KeyStr = rsItem("KeyStr")
            KeyScatterNum = rsItem("KeyScatterNum")

            NewsPaingType = rsItem("NewsPaingType")
            NPsString = rsItem("NPsString")
            NPoString = rsItem("NPoString")
            NewsPaingStr1 = rsItem("NewsPaingStr1")
            NewsPaingStr2 = rsItem("NewsPaingStr2")

            PsString = rsItem("PsString")
            PoString = rsItem("PoString")
            PhsString = rsItem("PhsString")
            PhoString = rsItem("PhoString")

            PsString = rsItem("PsString")
            PoString = rsItem("PoString")
            PhsString = rsItem("PhsString")
            PhoString = rsItem("PhoString")

            IsString = rsItem("IsString")
            IoString = rsItem("IoString")
            IntroType = rsItem("IntroType")
            IntroStr = rsItem("IntroStr")
            IntroNum = rsItem("IntroNum")

            IsField = rsItem("IsField")
            Field = rsItem("Field")
        End If

        rsItem.Close
        Set rsItem = Nothing

        If LsString = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�б�ʼ��ǲ���Ϊ�գ��޷��������뷵����һ���������ã�</li>"
        End If

        If LoString = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�б������ǲ���Ϊ�գ��޷��������뷵����һ���������ã�</li>"
        End If

        If ListPaingType = 0 Or ListPaingType = 1 Then
            If ListStr = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�б�����ҳ����Ϊ�գ��޷��������뷵����һ���������ã�</li>"
            End If

            If ListPaingType = 1 Then
                If LPsString = "" Or LPoString = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>������ҳ��ʼ��������ǲ���Ϊ�գ��޷��������뷵����һ���������ã�</li>"
                End If
            End If

            If ListPaingStr1 <> "" And Len(ListPaingStr1) < 15 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>������ҳ�ض������ò���ȷ���޷��������뷵����һ���������ã�</li>"
            End If

        ElseIf ListPaingType = 2 Then

            If ListPaingStr2 = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>��������ԭ�ַ�������Ϊ�գ��޷��������뷵����һ����������</li>"
            End If

            If IsNumeric(ListPaingID1) = False Or IsNumeric(ListPaingID2) = False Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�������ɵķ�Χֻ�������֣��޷��������뷵����һ����������</li>"
            Else

                If ListPaingID1 = 0 And ListPaingID2 = 0 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�������ɵķ�Χ����ȷ���޷��������뷵����һ����������</li>"
                End If
            End If

        ElseIf ListPaingType = 3 Then

            If ListPaingStr3 = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>������ҳ����Ϊ�գ��޷��������뷵����һ����������</li>"
            End If

        Else
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ѡ�񷵻���һ������������ҳ����</li>"
        End If

        If HsString = "" Or HoString = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���ӿ�ʼ/������ǲ���Ϊ�գ��޷�����,�뷵����һ����������</li>"
        End If

        If LoginType = 1 Then
            If LoginUrl = "" Or LoginPostUrl = "" Or LoginUser = "" Or LoginPass = "" Or LoginFalse = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�뽫��¼��Ϣ��д����</li>"
            Else
                LoginData = UrlEncoding(LoginUser & "&" & LoginPass)
                LoginResult = PostHttpPage(LoginUrl, LoginPostUrl, LoginData, PE_CLng(WebUrl))
                If InStr(LoginResult, LoginFalse) > 0 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>��¼��վʱ����������ȷ�ϵ�¼��Ϣ����ȷ�ԣ�</li>"
                End If
            End If
        End If
        
        If FoundErr <> True And LoginType = 1 Then
            LoginData = UrlEncoding(LoginUser & "&" & LoginPass)
            LoginResult = PostHttpPage(LoginUrl, LoginPostUrl, LoginData, PE_CLng(WebUrl))

            If InStr(LoginResult, LoginFalse) > 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>��¼��վʱ����������ȷ�ϵ�¼��Ϣ����ȷ�ԣ�</li>"
            End If
        End If
    
        If FoundErr <> True Then

            Select Case ListPaingType

                Case 0, 1
                    ListUrl = ListStr

                Case 2
                    ListUrl = Replace(ListPaingStr2, "{$ID}", CStr(ListPaingID1))

                Case 3

                    If InStr(ListPaingStr3, vbCrLf) > 0 Then
                        ListUrl = Left(ListPaingStr3, InStr(ListPaingStr3, vbCrLf))
                    Else
                        ListUrl = ListPaingStr3
                    End If

            End Select

        End If
            
        If FoundErr <> True Then
            ListCode = GetHttpPage(ListUrl, PE_CLng(WebUrl)) '��ȡ��ҳԴ����

            If ListCode <> "$False$" Then
                ListCode = GetBody(ListCode, LsString, LoString, False, False) '��ȡ�б�ҳ

                If ListCode <> "$False$" Then
                    NewsArrayCode = GetArray(ListCode, HsString, HoString, False, False) '��ȡ���ӵ�ַ

                    If NewsArrayCode <> "$False$" Then
                        If InStr(NewsArrayCode, "$Array$") > 0 Then
                            NewsArray = Split(NewsArrayCode, "$Array$") '�ָ�õ���ַ

                            If HttpUrlType = 1 Then
                                NewsUrl = Trim(Replace(HttpUrlStr, "{$ID}", NewsArray(PE_CLng(Testi))))
                            Else
                                NewsUrl = Trim(DefiniteUrl(NewsArray(PE_CLng(Testi)), ListUrl)) 'תΪ����·��
                            End If
                            
                            NewsPaingNextCode = GetHttpPage(NewsUrl, PE_CLng(WebUrl)) '��ȡ��ҳԴ����

                            If NewsPaingType = 1 Then '�������ô����ҳʱ
                                If NewsPaingStr1 <> "" And Len(NewsPaingStr1) > 15 Then
                                    '��ȡ��ҳ��ַ
                                    ListPaingNext = Replace(NewsPaingStr1, "{$ID}", GetPaing(NewsPaingNextCode, NPsString, NPoString, False, False))
                                Else

                                    ListPaingNext = GetPaing(NewsPaingNextCode, NPsString, NPoString, False, False) '��ȡ��ҳ��ַ
                                    If ListPaingNext <> "$False$" Then
                                        'Ӱ���˲�������ҳ��ҳ��ʱ��ֹ
                                        'If Left(ListPaingNext,1) = "/" then
                                            ConversionTrails = NewsUrl
                                        'Else
                                         '   ConversionTrails = ListUrl
                                        'End If

                                        ListPaingNext = DefiniteUrl(ListPaingNext, ConversionTrails) '�����·��ת����·��
                                    End If
                                End If

                            ElseIf NewsPaingType = 2 Then
                                
                                PageListCode = GetBody(NewsPaingNextCode, PsString, PoString, False, False) '��ȡ�б�ҳ

                                If PageListCode <> "$False$" Then

                                    PageArrayCode = GetArray(PageListCode, PhsString, PhoString, False, False) '��ȡ���ӵ�ַ
                                    
                                    If PageArrayCode <> "$False$" Then
                                        If InStr(PageArrayCode, "$Array$") > 0 Then
                                            PageArray = Split(PageArrayCode, "$Array$") '�ָ�õ���ַ

                                            For i = 0 To UBound(PageArray)
                                                'Ӱ���˲�������ҳ��ҳ��ʱ��ֹ
                                                'If Left(PageArray(i),1) = "/" then
                                                    ConversionTrails = NewsUrl
                                                'Else
                                                '    ConversionTrails = ListUrl
                                                'End If

                                                If ListPaingNext = "" Then
                                                    ListPaingNext = DefiniteUrl(PageArray(i), ConversionTrails) '�����·��ת����·��
                                                Else
                                                    ListPaingNext = ListPaingNext & "$Array$" & DefiniteUrl(PageArray(i), ConversionTrails) '�����·��ת����·��
                                                End If

                                            Next

                                            'ȥ����ַ��ʼ
                                            Dim TempPaingArray, tempj
                                            TempPaingArray = Split(ListPaingNext, "$Array$")
                                            ListPaingNext = ""
                                            For tempj = 0 To UBound(TempPaingArray)
                                                If InStr(LCase(ListPaingNext), LCase(TempPaingArray(tempj))) < 1 Then
                                                    ListPaingNext = ListPaingNext & "$Array$" & TempPaingArray(tempj)
                                                End If
                                            Next
                                            ListPaingNext = Right(ListPaingNext, Len(ListPaingNext) - 7)
                                            'ȥ����ַ����

                                        Else
                                            ListPaingNext = DefiniteUrl(PageArrayCode, NewsUrl) '�����·��ת����·��
                                        End If

                                    Else
                                        FoundErr = True
                                        ErrMsg = ErrMsg & "<li>�ڻ�ȡ��ҳ�����б�ʱ����</li>"
                                    End If

                                Else
                                    FoundErr = True
                                    ErrMsg = ErrMsg & "<li>�ڽ�ȡ��ҳ���뷢������</li>"
                                End If
                            End If

                        Else
                            FoundErr = True
                            ErrMsg = ErrMsg & "<li>ֻ����һ����Ч���ӣ���" & NewsArrayCode & "</li>"
                        End If

                    Else
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>�ڻ�ȡ�����б�ʱ����</li>"
                    End If

                Else
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�ڽ�ȡ�б�ʱ��������</li>"
                End If

            Else
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�ڻ�ȡ:" & ListUrl & "��ҳԴ��ʱ��������</li>"
            End If
        End If

        If FoundErr <> True Then
            NewsCode = GetHttpPage(NewsUrl, PE_CLng(WebUrl))

            If NewsCode <> "$False$" Then
                Title = GetBody(NewsCode, TsString, ToString, False, False)
                Content = GetBody(NewsCode, CsString, CoString, False, False)

                If Title = "$False$" Or Content = "$False$" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>�ڽ�ȡ����/���ĵ�ʱ��������" & NewsUrl & "</li>"
                Else
                    Title = FpHtmlEnCode(Title)
                    
                    If UpDateType = 0 Then
                        UpdateTime = Now()
                    ElseIf UpDateType = 1 Then

                        If DateType = 0 Then
                            UpdateTime = Now()
                        Else
                            UpdateTime = GetBody(NewsCode, DsString, DoString, False, False)
                            UpdateTime = PE_CDate(FpHtmlEnCode(UpdateTime))
                        End If

                    ElseIf UpDateType = 2 Then
                    Else
                        UpdateTime = Now()
                    End If

                    '�������
                    If AuthorType = 1 Then
                        Author = GetBody(NewsCode, AsString, AoString, False, False)
                    ElseIf AuthorType = 2 Then
                        Author = AuthorStr
                    Else
                        Author = "����"
                    End If

                    If Author = "$False$" Then
                        Author = "����"
                    Else
                        Author = FpHtmlEnCode(Trim(Author))
                    End If

                    If CopyFromType = 1 Then
                        CopyFrom = GetBody(NewsCode, FsString, FoString, False, False)
                    ElseIf CopyFromType = 2 Then
                        CopyFrom = CopyFromStr
                    Else
                        CopyFrom = "����"
                    End If

                    If CopyFrom = "$False$" Then
                        CopyFrom = "����"
                    Else
                        CopyFrom = FpHtmlEnCode(Trim(CopyFrom))
                    End If
                    
                    If KeyType = 0 Then
                        Key = Title
                        Key = CreateKeyWord(Key, KeyScatterNum)
                    ElseIf KeyType = 1 Then
                        Key = GetBody(NewsCode, KsString, KoString, False, False)
                        Key = FpHtmlEnCode(Key)
                        Key = Replace(Key, ",", "|")
                        Key = Replace(Key, "&nbsp;", "|")
                        Key = Replace(Key, " ", "|")

                        Dim arrKey, KeyString, j
                        arrKey = Split(Key, "|")
                        For j = 0 To UBound(arrKey)
                            If arrKey(j) <> "" Then
                                If KeyString = "" Then
                                    KeyString = arrKey(j)
                                Else
                                    KeyString = KeyString & "|" & arrKey(j)
                                End If
                            End If
                        Next
                        Key = KeyString
                        'Key = CreateKeyWord(Key, KeyScatterNum)
                        If Len(Key) > 253 Then
                            Key = "|" & Left(Key, 253) & "|"
                        Else
                            Key = "|" & Key & "|"
                        End If
                    ElseIf KeyType = 2 Then
                        Key = FpHtmlEnCode(KeyStr)
                    End If

                    If Key = "$False$" Or Trim(Key) = "" Then
                        Key = Title
                    End If

                    '���˷Ƿ��ַ�
                    Key = ReplaceBadChar(Key)

                    If IntroType = 0 Then
                    ElseIf IntroType = 1 Then
                        Intro = GetBody(NewsCode, IsString, IoString, False, False)
                        Intro = Trim(nohtml(Intro))
                    ElseIf IntroType = 2 Then
                        Intro = nohtml(IntroStr)
                    ElseIf IntroType = 3 Then
                        Intro = Left(Replace(Replace(Replace(Replace(Trim(nohtml(Content)), vbCrLf, ""), " ", ""), "&nbsp;", ""), "��", ""), IntroNum)
                    End If

                    If Intro = "$False$" Then
                        Intro = ""
                    End If
                End If

            Else
                FoundErr = True
                ErrMsg = ErrMsg & "<li>�ڻ�ȡԴ��ʱ��������" & NewsUrl & "</li>"
            End If
        End If

        If FoundErr <> True Then
            sql = "Select * from PE_Filters Where ItemID=" & ItemID & "  or ItemID = -1 order by FilterID ASC"
            Set rsFilters = Conn.Execute(sql)

            If rsFilters.EOF And rsFilters.BOF Then
            Else
                Arr_Filters = rsFilters.GetRows()
            End If

            rsFilters.Close
            Set rsFilters = Nothing

            If IsNull(Arr_Filters) = True Or IsArray(Arr_Filters) = False Then
            Else
                For Filteri = 0 To UBound(Arr_Filters, 2)
                    FilterStr = ""
                    If Arr_Filters(1, Filteri) = ItemID Or Arr_Filters(1, Filteri) = -1 And Arr_Filters(9, Filteri) = True Then
                        If Arr_Filters(3, Filteri) = 1 Then '�������
                            If Arr_Filters(4, Filteri) = 1 Then
                                Title = Replace(Title, Arr_Filters(5, Filteri), Arr_Filters(8, Filteri))
                            ElseIf Arr_Filters(4, Filteri) = 2 Then
                                FilterStr = GetBody(Title, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
                                Do While FilterStr <> "$False$"
                                    Title = Replace(Title, FilterStr, Arr_Filters(8, Filteri))
                                    FilterStr = GetBody(Title, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
                                Loop
                            End If
                        ElseIf Arr_Filters(3, Filteri) = 2 Then '���Ĺ���
                            If Arr_Filters(4, Filteri) = 1 Then
                                Content = Replace(Content, Arr_Filters(5, Filteri), Arr_Filters(8, Filteri))
                            ElseIf Arr_Filters(4, Filteri) = 2 Then
                                FilterStr = GetBody(Content, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
                                Do While FilterStr <> "$False$"
                                    Content = Replace(Content, FilterStr, Arr_Filters(8, Filteri))
                                    FilterStr = GetBody(Content, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
                                Loop
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    Response.Write "<br>" & vbCrLf
    Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"" >" & vbCrLf
    Response.Write "  <tr> " & vbCrLf
    Response.Write "    <td height=""22"" colspan=""2"" class=""title""> <div align=""center""><strong>�� �� �� ��</strong></div></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr> " & vbCrLf
    Response.Write "    <td height=""22"" colspan=""2"" class=""tdbg"" align=""center""><br><a href=""" & NewsUrl & """ target=""_blank""><font color=red>" & Title & "</font></a>" & vbCrLf
    Response.Write "      <br><br><center>��Դ��" & CopyFrom & "&nbsp;&nbsp;���ߣ�" & Author & "&nbsp;&nbsp;����ʱ�䣺" & UpdateTime & "</center>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf

    If IntroType <> 0 Then
        Response.Write "  <tr> " & vbCrLf
        Response.Write "    <td height=""22"" colspan=""2"" class=""tdbg"" align=""center"">��飺" & Intro & "</td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
    End If

    Response.Write "  <tr> " & vbCrLf
    Response.Write "    <td height=""400"" width=""100%"" colspan=""2"">" & vbCrLf
    Response.Write "       <textarea name='ListContent' style='display:none'>" & Content & " </textarea>"
    Response.Write "       <iframe src='Admin_CollectionPreview.asp?tContentid=ListContent' frameborder='1' scrolling='yes' width='100%' height='400' ></iframe>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'> " & vbCrLf
    Response.Write "    <td height=""22"" width=""1"" ></td>" & vbCrLf
    Response.Write "    <td height=""22""  class=""tdbg"" align=""left"">" & vbCrLf
    
    If UBound(Split(LCase(Content), "</table>")) > UBound(Split(LCase(Content), "<table>")) Or UBound(Split(LCase(Content), "<table>")) > UBound(Split(LCase(Content), "</table>")) Then
        Response.Write "<font color='red'>ע�⣺<br>&nbsp;&nbsp;&nbsp;&nbsp;�����õ��Ĵ����к���δ�رյ�HTML��ǩ�����磺���ڶ����&lt;/table&gt;������ܻᵼ������ҳ�İ�ʽ���Ρ������Է�����һ�������������ݽ�ȡ��ǣ���������һ���Ĳɼ����������ù�����Ŀ�����޸�������⡣</font>"
    End If
    
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr> " & vbCrLf
    Response.Write "    <td height=""22"" colspan=""2"" class=""tdbg"" align=""center"">�ؼ��֣�" & Key & "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf

    If IsField > 0 Then
        Response.Write "  <tr> " & vbCrLf
        Response.Write "    <td height=""22"" colspan=""2"" class=""tdbg"" align=""center"">�Զ����ֶΣ���</td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        If InStr(Field, "|||") > 0 Then
            arrField = Split(Field, "|||")

            For iField = 0 To UBound(arrField)
                arrField2 = Split(arrField(iField), "{#F}")
                FieldType = arrField2(2)
                FisSting = arrField2(3)
                FioSting = arrField2(4)
                FieldStr = arrField2(5)

                If FieldType = 0 Then
                    Field = "��������"
                ElseIf FieldType = 1 Then
                    Field = GetBody(NewsCode, FisSting, FioSting, False, False)
                    Field = Trim(Field)
                ElseIf FieldType = 2 Then
                    Field = FieldStr
                End If

                If Field = "$False$" Then
                    Field = "<font color='red'>��ȡ����!��ע��!ʧ��¼��Ĭ��Ϊ�ա�</font>"
                End If

                Response.Write "  <tr> " & vbCrLf
                Response.Write "    <td height=""22"" colspan=""2"" class=""tdbg"" align=""center"">" & vbCrLf
                Response.Write arrField2(1) & ":" & Field & "<br>"
                Response.Write "    </td> " & vbCrLf
                Response.Write "  </tr> " & vbCrLf
            Next

        Else

            If InStr(Field, "{#F}") > 0 Then
                arrField2 = Split(Field, "{#F}")
                FieldType = PE_CLng(arrField2(2))
                FisSting = arrField2(3)
                FioSting = arrField2(4)
                FieldStr = arrField2(5)
            End If

            If FieldType = 0 Then
                Field = "��������"
            ElseIf FieldType = 1 Then
                Field = GetBody(NewsCode, FisSting, FioSting, False, False)
                Field = Trim(Field)
            ElseIf FieldType = 2 Then
                Field = FieldStr
            End If

            If Field = "$False$" Then
                Field = "<font color='red'>��ȡ����!��ע��!ʧ��¼��Ĭ��Ϊ�ա�</font>"
            End If

            Response.Write "  <tr> " & vbCrLf
            Response.Write "    <td height=""22"" colspan=""2"" class=""tdbg"" align=""center"">" & vbCrLf
            Response.Write arrField2(1) & ":" & Field & "<br>"
            Response.Write "    </td> " & vbCrLf
            Response.Write "  </tr> " & vbCrLf
        End If
    End If

    Response.Write "  <tr> " & vbCrLf
    Response.Write "    <td height=""22"" colspan=""2"" class=""tdbg"" align=""center"">" & vbCrLf
    
    If NewsPaingType = 1 And ListPaingNext <> "$False$" And ListPaingNext <> "" Then
        Response.Write "&nbsp;&nbsp;���ķ�ҳ����һҳ��<a href=" & ListPaingNext & " target=_blank><font color=Red>" & ListPaingNext & "</font></a>&nbsp;&nbsp;&nbsp;&nbsp;�����Ƿ���ȷ"
    ElseIf NewsPaingType = 2 And ListPaingNext <> "$False$" And ListPaingNext <> "" Then
        Response.Write "&nbsp;&nbsp;�����������ķ�ҳ��ַ�Ƿ���ȷ��<br>"

        If InStr(ListPaingNext, "$Array$") = 0 Then
            ListPaingNext = "<a href=" & ListPaingNext & " target=_blank><font color=Red>" & ListPaingNext & "</font></a>"
        Else
            PageArray = Split(ListPaingNext, "$Array$")
            ListPaingNext = ""

            For i = 0 To UBound(PageArray)

                If ListPaingNext = "" Then
                    ListPaingNext = "<a href=" & PageArray(i) & " target=_blank><font color=Red>" & PageArray(i) & "</font></a>"
                Else
                    ListPaingNext = ListPaingNext & "<br>" & "<a href=" & PageArray(i) & " target=_blank><font color=Red>" & PageArray(i) & "</font></a>"
                End If

            Next

        End If

        Response.Write ListPaingNext
    ElseIf Trim(ListPaingNext) = "$False$" Then
        Response.Write "<font color=Red>���ķ�ҳ����ʧ��/���ߵ�ǰ�����޷�ҳ</font>"
    End If

    Response.Write "</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<form method=""post"" action=""Admin_CollectionManage.asp"" name=""form1"">" & vbCrLf
    Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"" >" & vbCrLf
    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td colspan=""2"" align=""center"" class=""tdbg"" height=""50"">" & vbCrLf
    Response.Write "      <input name=""ItemID"" type=""hidden"" id=""ItemID"" value=""" & ItemID & """>" & vbCrLf
    Response.Write "      <input name=""Action"" type=""hidden"" id=""Action"" value=""Step5"">" & vbCrLf
    Response.Write "      <input name=""Flag"" type=""hidden"" id=""Flag"" value=""True"">" & vbCrLf
    Response.Write "      <input name=""Cancel"" type=""button"" id=""Cancel"" value=""������һ��"" onClick=""window.location.href='javascript:history.go(-1)'"" >&nbsp;&nbsp;" & vbCrLf
    Response.Write "      <input  type=""submit"" name=""Submit"" value="" ��һ�� ""  >" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "   </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</form>" & vbCrLf

    Call CloseConn
End Sub
'=================================================
'��������Step5
'��  �ã���������
'=================================================
Sub Step5()
    '�������Ա���
    Dim ClassID, arrSpecialID
    Dim PaginationType, MaxCharPerPage, InfoPoint, Stars, OnTop, Hot, Elite, Hits
    Dim UploadDir, UpFileType, IncludePicYn, DefaultPicYn, SkinID, TemplateID, Flag
    Dim Script_Table, Script_Tr, Script_Td, ShowCommentLink
    
    Dim UpDateType, UpdateTime
    '���˱���
    Dim Script_Iframe, Script_Object, Script_Script, Script_Class, Script_Font, Script_A
    Dim Script_Img, Script_Div, Script_Span, Script_Html
    '�ɼ�����
    Dim CollecOrder, Status, CreateImmediate, CollectionNum, CollectionType
    Dim SaveFiles, AddWatermark, AddThumb, SaveFlashUrlToFile
    Dim InfoPurview, arrGroupID, ChargeType, PitchTime, ReadTimes, DividePercent
    Dim ChannelShortName, strDisabled
    ChannelShortName = "����"
    strDisabled = ""
    ChannelID = PE_CLng(Trim(Request("ChannelID")))

    If ItemID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ĿID����Ϊ��</li>"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If
    
    sql = "select * from PE_Item where ItemID=" & ItemID
    Set rsItem = Server.CreateObject("adodb.recordset")
    rsItem.Open sql, Conn, 1, 1
    If rsItem.EOF Then   'û���ҵ�����Ŀ
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���������û���ҵ�����Ŀ��</li>"
    Else
        ItemName = rsItem("ItemName")
        If ChannelID = 0 Then
            ChannelID = PE_CLng(rsItem("ChannelID"))
        End If
        ClassID = rsItem("ClassID")
        arrSpecialID = rsItem("SpecialID")
        PaginationType = rsItem("PaginationType")
        If rsItem("MaxCharPerPage") = 0 Then
            MaxCharPerPage = 10000
        Else
            MaxCharPerPage = rsItem("MaxCharPerPage")
        End If
        OnTop = rsItem("OnTop")
        Hot = rsItem("Hot")
        Hits = rsItem("Hits")
        Elite = rsItem("Elite")
        UpdateTime = rsItem("UpdateTime")
        SkinID = rsItem("SkinID")
        Stars = rsItem("Stars")
        TemplateID = rsItem("TemplateID")
        IncludePicYn = rsItem("IncludePicYn")
        DefaultPicYn = rsItem("DefaultPicYn")
        AddWatermark = rsItem("AddWatermark")
        AddThumb = rsItem("AddThumb")
        SaveFlashUrlToFile = rsItem("SaveFlashUrlToFile")
        CollecOrder = rsItem("CollecOrder")
        Script_Iframe = rsItem("Script_Iframe")
        Script_Object = rsItem("Script_Object")
        Script_Script = rsItem("Script_Script")
        Script_Class = rsItem("Script_Class")
        Script_Div = rsItem("Script_Div")
        Script_Span = rsItem("Script_Span")
        Script_Img = rsItem("Script_Img")
        Script_Html = rsItem("Script_Html")
        Script_Font = rsItem("Script_Font")
        Script_A = rsItem("Script_A")
        Script_Table = rsItem("Script_Table")
        Script_Tr = rsItem("Script_Tr")
        Script_Td = rsItem("Script_Td")
        ShowCommentLink = rsItem("ShowCommentLink")
        SaveFiles = rsItem("SaveFiles")
        Status = PE_CLng(rsItem("Status"))
        IncludePicYn = rsItem("IncludePicYn")
        DefaultPicYn = rsItem("DefaultPicYn")
        CreateImmediate = rsItem("CreateImmediate")
        CollectionNum = PE_CLng(rsItem("CollectionNum"))
        CollectionType = PE_CLng(rsItem("CollectionType"))
        Flag = rsItem("Flag")
        UpDateType = PE_CLng(rsItem("UpdateType"))
        UpdateTime = rsItem("UpdateTime")

        InfoPurview = rsItem("InfoPurview")
        arrGroupID = rsItem("arrGroupID")
        InfoPoint = rsItem("InfoPoint")
        ChargeType = rsItem("ChargeType")
        PitchTime = rsItem("PitchTime")
        ReadTimes = rsItem("ReadTimes")
        DividePercent = rsItem("DividePercent")
    End If

    rsItem.Close
    Set rsItem = Nothing
        
    '��������ȵ���
    If ChannelID > 0 Then
        Dim sqlChannel, rsChannel, HitsOfHot
        sqlChannel = "select * from PE_Channel where ChannelID=" & ChannelID & " and ModuleType = 1 order by OrderID"
        Set rsChannel = Conn.Execute(sqlChannel)
        If rsChannel.BOF And rsChannel.EOF Then
            HitsOfHot = 1000
        Else
            HitsOfHot = rsChannel("HitsOfHot")
        End If
        rsChannel.Close
        Set rsChannel = Nothing
    Else
        HitsOfHot = 1000
        ChannelID = 1
    End If
    
    Call PopCalendarInit
    Response.Write "<script language=""JavaScript"">" & vbCrLf
    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "<!--" & vbCrLf
    Response.Write "// ֻ������������" & vbCrLf
    Response.Write "function IsDigit(){" & vbCrLf
    Response.Write "  return ((event.keyCode >= 48) && (event.keyCode <= 57));" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "    ShowTabs(0);" & vbCrLf
    Response.Write "    if (document.form1.ClassID.value==''){" & vbCrLf
    Response.Write "        alert('����������Ŀ����Ϊ��,���ƶ�Ϊ�ⲿ��Ŀ��');" & vbCrLf
    Response.Write "        document.form1.ClassID.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.form1.ClassID.value=='0'){" & vbCrLf
    Response.Write "        alert('ָ������Ŀ������������£�ֻ������������Ŀ��������¡�');" & vbCrLf
    Response.Write "        document.form1.ClassID.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  return true;  " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

    Response.Write "<br>" & vbCrLf
    Response.Write "<form method=""post"" action=""Admin_CollectionManage.asp"" name=""form1"" onSubmit=""return CheckForm();"">" & vbCrLf
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center' height='24'>" & vbCrLf
    Response.Write "   <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>��������</td>" & vbCrLf
    Response.Write "   <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>��������</td>" & vbCrLf
    Response.Write "   <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>�շ�����</td>" & vbCrLf
    Response.Write "   <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>�ɼ�����</td>" & vbCrLf
    Response.Write "   <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right""> ����Ƶ����&nbsp;</td>" & vbCrLf
    Response.Write "          <td class=""tdbg"">" & vbCrLf
    Response.Write "            <select name='ChannelID' onChange=""javascript:location.href='Admin_CollectionManage.asp?Action=Step5&Flag=" & Trim(Request("Flag")) & "&ItemID=" & ItemID & "&ChannelID=' + (this.options[this.selectedIndex].value)"">" & vbCrLf
    Call GetChannel_Option(ChannelID)
    Response.Write "            </select> " & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "          <tr class='tdbg'>" & vbCrLf
    Response.Write "            <td width=""120"" class=""tdbg5"" align=""right""> ��Ŀ/ר�⣺&nbsp;</td>" & vbCrLf
    Response.Write "            <td  class=""tdbg"">" & vbCrLf
    Response.Write "              <table width='98%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "                <tr align='left'><td><b>������Ŀ��</b></td><td><b>����ר�⣺</b></td></tr>"
    Response.Write "                <tr align='left'><td><select name='ClassID' size='2'" & strDisabled & " style='height:300px;width:260px;'>" & GetClass_Option(3, ClassID) & "</select></td>"
    Response.Write "                  <td><select name='SpecialID' size='2'" & strDisabled & " multiple style='height:300px;width:260px;'>" & GetSpecial_Option(arrSpecialID) & "</select></td></tr>"
    Response.Write "                <tr align='left'><td>����ѡ���ⲿ��Ŀ"
    If AdminPurview = 2 And AdminPurview_Channel = 3 Then
        Response.Write "<br>��ֻ����<font color='#FF0000'>��ɫ��Ŀ</font>��������Ŀ�з���" & ChannelShortName & ""
    End If
    Response.Write "</td><td>��סCtrl������ͬʱѡ����ר��</td></tr></table>" & vbCrLf
    Response.Write "            </td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'>" & vbCrLf
    Response.Write "        <td width='120' align='right' class='tdbg5' align=""right"">����״̬��&nbsp;</td>"
    Response.Write "         <td>" & vbCrLf
    Response.Write "           <input name='Status' type='radio' id='Status' value='-1' " & RadioValue(Status, -1) & ">�ݸ�&nbsp;&nbsp;" & vbCrLf
    Response.Write "           <input Name='Status' Type='Radio' Id='Status' Value='0' " & RadioValue(Status, 0) & ">�����&nbsp;&nbsp;" & vbCrLf
    Response.Write "           <input Name='Status' Type='Radio' Id='Status' Value='3' " & RadioValue(Status, 3) & "> ����ͨ��" & vbCrLf
  '  If UseCreateHTML > 0 And AutoCreateType > 0 Then
    Response.Write "            &nbsp;&nbsp;<input name=""CreateImmediate"" type=""checkbox"" id=""CreateImmediate"" value=""yes"" "
    If CreateImmediate = True Then Response.Write " checked"
    Response.Write ">�������� " & vbCrLf
  '  End If
    Response.Write "         </td>"
    Response.Write "        </tr>"
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "        <tr class='tdbg' id='ArticleContent2' style=""display:''""> "
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right"">���ݷ�ҳ��ʽ��&nbsp;</td>"
    Response.Write "          <td>"
    Response.Write "            <select name='PaginationType' id='PaginationType' onchange=""javascript:if(document.form1.PaginationType.selectedIndex==1){document.form1.Script_Table.checked=true;document.form1.Script_Tr.checked=true;document.form1.Script_td.checked=true;document.form1.Script_Script.checked=true;document.form1.Script_Object.checked=true;document.form1.Script_Iframe.checked=true;document.form1.Script_A.checked=true;document.form1.Script_Div.checked=true;document.form1.Script_Span.checked=true;}"">" & vbCrLf
    Response.Write "               <option value='2'"
    If PaginationType = 2 Then Response.Write " selected"
    Response.Write ">�ֶ���ҳ</option>"
    Response.Write "           <option value='1'"
    If PaginationType = 1 Then Response.Write " selected"
    Response.Write ">�Զ���ҳ</option>"
    Response.Write "           <option value='0'"
    If PaginationType = 0 Then Response.Write " selected"
    Response.Write ">����ҳ</option>"
    Response.Write "            </select>"
    Response.Write "        �Զ���ҳʱ��ÿҳ��Լ�ַ���������HTML��ǣ��� <input name='MaxCharPerPage' type='text' id='MaxCharPerPage' value=" & MaxCharPerPage & " size='8' maxlength='8'><br><font color='blue'>ע�⣺������òɼ���������Զ����ɼ���ģʽ��ҳ��������ѡ���ֶ���ҳ��</font>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right"">�������ԣ�&nbsp;</td>"
    Response.Write "          <td>"
    Response.Write "            <input name=""IncludePicYn"" type=""checkbox"" id=""SaveFiles"" value=""yes"" "
    If IncludePicYn = True Then Response.Write " checked"
    Response.Write ">����ͼƬ" & vbCrLf
    Response.Write "            <input name=""DefaultPicYn"" type=""checkbox"" value=""yes"" "
    If DefaultPicYn = True Then Response.Write " checked"
    Response.Write ">��ҳͼƬ" & vbCrLf
    Response.Write "            <input name='OnTop' type='checkbox' id='OnTop' value='yes'"
    If OnTop = True Then Response.Write " checked"
    Response.Write " > �̶�����&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "            <input name='Hot' type='checkbox' id='Hot' value='yes' ONKEYPRESS=""event.returnValue=IsDigit();"" onclick=""javascript:document.form1.Hits.value='" & HitsOfHot & "'""> ��������&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "            <input name='Elite' type='checkbox' id='Elite' value='yes'"
    If Elite = True Then Response.Write " checked"
    Response.Write "> �Ƽ�����&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "     �������ֵȼ��� <select name='Stars' id='Stars'>" & GetStars(Stars) & "</select>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right"">�������ʼֵ��&nbsp;</td>"
    Response.Write "          <td>"
    Response.Write "            <input name='Hits' type='text' id='Hits' value='" & Hits & "' size='10' maxlength='10' ONKEYPRESS=""event.returnValue=IsDigit();"">&nbsp;&nbsp; <font color='#0000FF'>�⹦�����ṩ������Ա�����õġ�����������Ҫ��ѽ��^_^</font>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class=""tdbg"">" & vbCrLf
    Response.Write "          <td height=""30"" width=""120"" class=""tdbg5"" align=""right"">����¼��ʱ�䣺&nbsp;</td>" & vbCrLf
    Response.Write "          <td>" & vbCrLf
    Response.Write "            <Input name=""UpdateType"" type=""radio"" value=""0"" " & vbCrLf
    If UpDateType = 0 Then Response.Write "checked"
    Response.Write " >��ǰʱ��" & vbCrLf
    Response.Write "            <Input name=""UpdateType"" type=""radio"" value=""1"" " & vbCrLf
    If UpDateType = 1 Then Response.Write "checked"
    Response.Write " >��ǩ�е�ʱ��" & vbCrLf
    Response.Write "            <Input name=""UpdateType"" type=""radio"" value=""2"" " & vbCrLf
    If UpDateType = 2 Then Response.Write "checked"
    Response.Write " >�Զ��壺" & vbCrLf
    Response.Write "            <Input name='UpdateTime' type='text' size='20' maxlength='20' value='" & UpdateTime & "' maxlength='50' onClick='PopCalendar.show(document.form1.UpdateTime, ""yyyy-mm-dd"", null, null, null, ""11"");'><a style='cursor:hand;' onClick='PopCalendar.show(document.form1.UpdateTime, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>"
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg"">" & vbCrLf
    Response.Write "          <td height=""30"" width=""120"" class=""tdbg5"" align=""right"">�������ӣ�&nbsp;</td>" & vbCrLf
    Response.Write "          <td>" & vbCrLf
    Response.Write "            <Input name=""ShowCommentLink"" type=""radio"" id=""ShowCommentLink"" value=""yes"" "
    If ShowCommentLink = True Then Response.Write "Checked"
    Response.Write " >��ʾ��������  " & vbCrLf
    Response.Write "            <Input name=""ShowCommentLink"" type=""radio"" id=""ShowCommentLink"" value=""no"" "
    If ShowCommentLink = False Then Response.Write "Checked"
    Response.Write " >����ʾ��������" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width=""120""  class=""tdbg5"" align=""right"">��ɫ���&nbsp;</td>"
    Response.Write "          <td><select Name='SkinID'>" & GetSkin_Option(SkinID) & "</select>&nbsp;���ģ���а���CSS����ɫ��ͼƬ����Ϣ</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "            <td width=""120"" class=""tdbg5"" align=""right"">�������ģ�壺&nbsp;</td>"
    Response.Write "            <td><select Name='TemplateID'>" & GetTemplate_Option(ChannelID, 3, TemplateID) & "</select>&nbsp;���ģ���а����˰�����Ƶİ�ʽ����Ϣ</td>"
    Response.Write "        </tr>"
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�Ķ�Ȩ�ޣ�&nbsp;</td>"
    Response.Write "            <td><input name='InfoPurview' type='radio' value='0'" & strDisabled
    If InfoPurview = 0 Then Response.Write " checked"
    Response.Write ">�̳���ĿȨ�ޣ���������ĿΪ��֤��Ŀʱ������ѡ����<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='1'" & strDisabled
    If InfoPurview = 1 Then Response.Write " checked"
    Response.Write ">���л�Ա����������ĿΪ������Ŀ���뵥����ĳЩ���½��в鿴Ȩ�����ã�����ѡ����<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='2'" & strDisabled
    If InfoPurview = 2 Then Response.Write " checked"
    Response.Write ">ָ����Ա�飨��������ĿΪ������Ŀ���뵥����ĳЩ���½��в鿴Ȩ�����ã�����ѡ����<br>"
    Response.Write GetUserGroup(arrGroupID & "", strDisabled)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'> " & ChannelShortName & "�Ķ�������&nbsp; </td>"
    Response.Write "            <td><input name='InfoPoint' type='text' id='InfoPoint' value='" & InfoPoint & "' size='5' maxlength='4' style='text-align:center'" & strDisabled & ">&nbsp;&nbsp;&nbsp;&nbsp; <font color='#0000FF'>�������0�����Ա�Ķ���" & ChannelShortName & "ʱ��������Ӧ��������Ϊ9999ʱ���⣩���οͽ��޷��鿴��" & ChannelShortName & "��</font></td>"
    Response.Write "          </tr>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ظ��շѣ�&nbsp; </td>"
    Response.Write "            <td><input name='ChargeType' type='radio' value='0'" & strDisabled
    If ChargeType = 0 Then Response.Write " checked"
    Response.Write ">���ظ��շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='1'" & strDisabled
    If ChargeType = 1 Then Response.Write " checked"
    Response.Write ">�����ϴ��շ�ʱ�� <input name='PitchTime' type='text' value='" & PitchTime & "' size='8' maxlength='8' style='text-align:center'" & strDisabled & "> Сʱ�������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='2'" & strDisabled
    If ChargeType = 2 Then Response.Write " checked"
    Response.Write ">��Ա�ظ��鿴������ <input name='ReadTimes' type='text' value='" & ReadTimes & "' size='8' maxlength='8' style='text-align:center'" & strDisabled & "> �κ������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='3'" & strDisabled
    If ChargeType = 3 Then Response.Write " checked"
    Response.Write ">�������߶�����ʱ�����շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='4'" & strDisabled
    If ChargeType = 4 Then Response.Write " checked"
    Response.Write ">����������һ������ʱ�������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='5'" & strDisabled
    If ChargeType = 5 Then Response.Write " checked"
    Response.Write ">ÿ�Ķ�һ�ξ��ظ��շ�һ�Σ����鲻Ҫʹ�ã�"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ֳɱ�����&nbsp; </td>"
    Response.Write "            <td><input name='DividePercent' type='text' id='DividePercent' value='" & DividePercent & "' size='5' maxlength='4' style='text-align:center'" & strDisabled & "> %"
    Response.Write "              &nbsp;&nbsp;<font color='#0000FF'>�����������0���򽫰����������Ķ�����ȡ�ĵ���֧����¼����</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right"">����ѡ�&nbsp;</td>"
    Response.Write "          <td height=""22"">"
    Response.Write "      <input name=""Script_Iframe"" type=""checkbox"" id=""Script_Iframe""  value=""yes"" "
    If Script_Iframe = True Then Response.Write " checked"
    Response.Write ">Iframe��  &nbsp;��������ҳ��<br>" & vbCrLf
    Response.Write "      <input name=""Script_Object"" type=""checkbox"" id=""Script_Object""  value=""yes"" "
    If Script_Object = True Then Response.Write " checked"
    Response.Write ">Object�� &nbsp;����Falsh���,�ؼ��ȡ�<br>" & vbCrLf
    Response.Write "      <input name=""Script_Script"" type=""checkbox"" id=""Script_Script""  value=""yes"" "
    If Script_Script = True Then Response.Write " checked"
    Response.Write ">Script�� &nbsp;����js��vbs�Ƚű���<br>" & vbCrLf
    Response.Write "      <input name=""Script_Class"" type=""checkbox"" id=""Script_Class""  value=""yes"" "
    If Script_Class = True Then Response.Write " checked"
    Response.Write ">Style�� &nbsp;����Css �ࡣ<br>" & vbCrLf
    Response.Write "      <input name=""Script_Div"" type=""checkbox"" id=""Script_Div""  value=""yes"" "
    If Script_Div = True Then Response.Write " checked"
    Response.Write ">Div�� &nbsp;���˲㡣<br>" & vbCrLf
    Response.Write "      <input name=""Script_Span"" type=""checkbox"" id=""Script_Span""  value=""yes"" "
    If Script_Span = True Then Response.Write " checked"
    Response.Write ">Span�� ��������Ԫ��Span������<br>" & vbCrLf
    Response.Write "      <input name=""Script_Table"" type=""checkbox"" id=""Script_Table""  value=""yes"" "
    If Script_Table = True Then Response.Write " checked"
    Response.Write ">Table" & vbCrLf
    Response.Write "      <input name=""Script_Tr"" type=""checkbox"" id=""Script_tr""  value=""yes"" "
    If Script_Tr = True Then Response.Write " checked"
    Response.Write ">Tr" & vbCrLf
    Response.Write "      <input name=""Script_td"" type=""checkbox"" id=""Script_td""  value=""yes"" "
    If Script_Td = True Then Response.Write " checked"
    Response.Write ">Td �����˱�����ԡ�<br>" & vbCrLf
    Response.Write "      <input name=""Script_Img"" type=""checkbox"" id=""Script_Img""  value=""yes"" "
    If Script_Img = True Then Response.Write " checked"
    Response.Write ">Img��&nbsp;����ͼƬ��<Font color=blue >ע�����ѡ�����ͼƬ�ɼ������������н�������ͼƬ</Font><br>" & vbCrLf
    Response.Write "      <input name=""Script_Font"" type=""checkbox"" id=""Script_Font""  value=""yes"" "
    If Script_Font = True Then Response.Write " checked"
    Response.Write ">FONT��&nbsp;�������嶨�塣 (��������ʽȥ��)<br>" & vbCrLf
    Response.Write "      <input name=""Script_A"" type=""checkbox"" id=""Script_A""  value=""yes"" "
    If Script_A = True Then Response.Write " checked"
    Response.Write ">A��&nbsp;�������� (����������ȥ��)<br>" & vbCrLf
    Response.Write "      <input name=""Script_Html"" type=""checkbox"" id=""Script_Html""  value=""yes"" "
    If Script_Html = True Then Response.Write " checked"
    Response.Write ">Html�� &nbsp;���˲ɼ�����ҳ�е�html�ַ���<Font color=blue >ע�����ѡ�����HTML�ɼ����������ݽ��Դ��ı���ʽ����</Font><br>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right"">�ɼ�������&nbsp;</td>"
    Response.Write "          <td  height=""30"" class=""tdbg"">" & vbCrLf
    Response.Write "            <Input type='radio' Name='iType' value='0' "
    If CollectionNum = 0 Then
        Response.Write " checked" & vbCrLf
    End If
    Response.Write "> �ɼ��б��е���������  <br>" & vbCrLf
    Response.Write "            <Input type='radio' Name='iType' value='1' "
    If CollectionType = 0 And CollectionNum <> 0 Then
        Response.Write " checked" & vbCrLf
    End If
    Response.Write "> �ɼ��б��е� <Input TYPE='text' Name='AritcleNum' value='" & CollectionNum & "' size='3' maxlength='5' ONKEYPRESS=""event.returnValue=IsDigit();"">ƪ���º�ֹͣ�ɼ� <br>" & vbCrLf
    Response.Write "            <Input type='radio' Name='iType' value='2' " & RadioValue(CollectionType, 1) & "> �ɼ��б��е� <Input TYPE='text' Name='PageNum' value='" & CollectionNum & "' size='3' maxlength='5' ONKEYPRESS=""event.returnValue=IsDigit();"">����ҳ��ֹͣ�ɼ� <br>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right"">�ɼ�ͼƬ���ã�&nbsp;</td>"
    Response.Write "          <td  height=""30"" class=""tdbg"">" & vbCrLf
    Response.Write "           <input name=""SaveFiles"" type=""checkbox"" id=""SaveFiles"" value=""yes"" " & RadioValue(SaveFiles, True) & ">����Զ��ͼƬ" & vbCrLf
    Response.Write "           <input name=""AddWatermark"" type=""checkbox"" value=""yes"" " & RadioValue(AddWatermark, True) & ">�Զ���ͼƬ����ˮӡ" & vbCrLf
    Response.Write "           <input name=""AddThumb"" type=""checkbox"" value=""yes"" " & RadioValue(AddThumb, True) & ">�Զ�Ϊ��һ��ͼƬ��������ͼ<br>" & vbCrLf
    Response.Write "           <input name=""SaveFlashUrlToFile"" type=""checkbox"" value=""yes"" " & RadioValue(SaveFlashUrlToFile, True) & ">�����������е�Flash��ͼƬ�ĵ�ַ���浽��Ŀ¼�е�CollectionFilePath.txt�ļ��У��Է������ʿ쳵�������������" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right"">���²ɼ�˳��&nbsp;</td>"
    Response.Write "          <td  height=""30"" class=""tdbg"">" & vbCrLf
    Response.Write "            <Input type='radio' Name='CollecOrder' value='0' " & RadioValue(PE_CLng(CollecOrder), 0) & ">����ɼ�"
    Response.Write "            <Input type='radio' Name='CollecOrder' value='1' " & RadioValue(PE_CLng(CollecOrder), 1) & ">����ɼ� <FONT color='blue'>���Ƽ���</FONT>"
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "       </td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "     </table>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "   </tr>" & vbCrLf
    Response.Write " </table>" & vbCrLf
    Response.Write " <br>" & vbCrLf
    Response.Write " <center>" & vbCrLf
    Response.Write "      <input name=""ItemID"" type=""hidden"" id=""ItemID"" value=""" & ItemID & """>" & vbCrLf
    Response.Write "      <input name=""Action"" type=""hidden"" id=""Action"" value=""Step6"">" & vbCrLf
    Response.Write "      <input name=""Flag"" type=""hidden"" id=""Flag"" value=" & Trim(Request("Flag")) & ">" & vbCrLf
    Response.Write "      <input name=""Cancel"" type=""button"" id=""Cancel"" value=""������һ��"" onClick=""window.location.href='javascript:history.go(-1)'"" >&nbsp;&nbsp;" & vbCrLf
    Response.Write "      <input  type=""submit"" name=""Submit"" value="" ��  �� "" >" & vbCrLf
    Response.Write " </center>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    
    Call CloseConn
End Sub
'=================================================
'��������Step6
'��  �ã���˱�����Ŀ
'=================================================
Sub Step6()
    '�������Ա���
    Dim ItemName, ChannelID, ClassID, SpecialID, PaginationType, MaxCharPerPage
    Dim OnTop, Hot, Elite, Hits, Stars, UpDateType, UpdateTime, IncludePicYn, DefaultPicYn, SkinID, TemplateID
    Dim UploadDir, UpFileType
    '�����շѱ���
    Dim InfoPurview, arrGroupID, InfoPoint, ChargeType, PitchTime, ReadTimes, DividePercent

    '���˱���
    Dim Script_Iframe, Script_Object, Script_Script, Script_Class
    Dim Script_Div, Script_Span, Script_Img, Script_Font, Script_A, Script_Html
    Dim Script_Table, Script_Tr, Script_Td, ShowCommentLink
    '�ɼ�����
    Dim SaveFiles, AddWatermark, AddThumb, SaveFlashUrlToFile, iType, CollecOrder, Status, CreateImmediate, CollectionNum, CollectionType

    
    If ItemID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ĿID����Ϊ��</li>"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    If NeedSave = "" Then
        ItemName = Trim(Request.Form("ItemName"))
        ChannelID = Trim(Request.Form("ChannelID"))
        ClassID = Trim(Request.Form("ClassID"))
        SpecialID = Replace(ReplaceBadChar(Trim(Request.Form("SpecialID"))), " ", "")
        PaginationType = Trim(Request.Form("PaginationType"))
        MaxCharPerPage = Trim(Request.Form("MaxCharPerPage"))
        OnTop = Trim(Request.Form("OnTop"))
        Hot = Trim(Request.Form("Hot"))
        Elite = Trim(Request.Form("Elite"))
        Hits = Trim(Request.Form("Hits"))
        Stars = Trim(Request.Form("Stars"))
        UpDateType = Trim(Request.Form("UpdateType"))
        UpdateTime = Trim(Request.Form("UpdateTime"))
        SkinID = Trim(Request.Form("SkinID"))
        TemplateID = Trim(Request.Form("TemplateID"))
        IncludePicYn = Trim(Request.Form("IncludePicYn"))
        DefaultPicYn = Trim(Request.Form("DefaultPicYn"))

        InfoPurview = PE_CLng(Trim(Request.Form("InfoPurview")))
        arrGroupID = Trim(Request.Form("GroupID"))
        InfoPoint = PE_CLng(Trim(Request.Form("InfoPoint")))
        ChargeType = PE_CLng(Trim(Request.Form("ChargeType")))
        PitchTime = PE_CLng(Trim(Request.Form("PitchTime")))
        ReadTimes = PE_CLng(Trim(Request.Form("ReadTimes")))
        DividePercent = PE_CLng(Trim(Request.Form("DividePercent")))

        Script_Iframe = Trim(Request.Form("Script_Iframe"))
        Script_Object = Trim(Request.Form("Script_Object"))
        Script_Script = Trim(Request.Form("Script_Script"))
        Script_Class = Trim(Request.Form("Script_Class"))
        Script_Div = Trim(Request.Form("Script_Div"))
        Script_Span = Trim(Request.Form("Script_Span"))
        Script_Img = Trim(Request.Form("Script_Img"))
        Script_Font = Trim(Request.Form("Script_Font"))
        Script_A = Trim(Request.Form("Script_A"))
        Script_Html = Trim(Request.Form("Script_Html"))
        Script_Table = Trim(Request.Form("Script_Table"))
        Script_Tr = Trim(Request.Form("Script_Tr"))
        Script_Td = Trim(Request.Form("Script_Td"))
        
        ShowCommentLink = Trim(Request("ShowCommentLink"))
        SaveFiles = Trim(Request.Form("SaveFiles"))
        SaveFlashUrlToFile = Trim(Request.Form("SaveFlashUrlToFile"))
        AddWatermark = Trim(Request.Form("AddWatermark"))
        AddThumb = Trim(Request.Form("AddThumb"))
        CollecOrder = PE_CLng(Trim(Request.Form("CollecOrder")))
        Status = PE_CLng(Trim(Request.Form("Status")))
        CreateImmediate = Trim(Request.Form("CreateImmediate"))
        iType = PE_CLng(Trim(Request.Form("iType")))

        Select Case iType
        Case 0  '�ɼ���������
          '����д����
            CollectionType = 0
        Case 1
            CollectionType = 0
            CollectionNum = PE_CLng(Trim(Request("AritcleNum")))
        Case 2
            CollectionType = 1
            CollectionNum = PE_CLng(Trim(Request("PageNum")))
        End Select


        If FoundErr = True Then Exit Sub

        If ChannelID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���������Ƶ��</li>"
        Else
            ChannelID = PE_CLng(ChannelID)
        End If

        If ClassID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>δָ����Ŀ������Ŀ����ָ������Ŀ����������Ŀ</li>"
        Else
            ClassID = PE_CLng(ClassID)
            If ClassID = 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>ָ���˷Ƿ�����Ŀ���ⲿ��Ŀ�򲻴��ڵ���Ŀ��</li>"
            End If
        End If
        If SpecialID = "" Then
            SpecialID = 0
        End If
        If SkinID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ����ɫ���</li>"
        Else
            SkinID = PE_CLng(SkinID)
        End If
        If TemplateID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ���������ģ��</li>"
        Else
            TemplateID = PE_CLng(TemplateID)
        End If
        If InfoPoint = "" Then
            InfoPoint = 0
        Else
            InfoPoint = PE_CLng(InfoPoint)
        End If
        If PaginationType = "" Then
            PaginationType = 0
        Else
            PaginationType = PE_CLng(PaginationType)
        End If
        If MaxCharPerPage = "" Then
            MaxCharPerPage = 0
        Else
            MaxCharPerPage = PE_CLng(MaxCharPerPage)
        End If
        If PaginationType = 1 And MaxCharPerPage = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ���Զ���ҳʱ��ÿҳ��Լ�ַ���,�������0</li>"
        End If
        If UpDateType = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ѡ������¼��ʱ�����ͣ�</li>"
        Else
            UpDateType = CLng(UpDateType)
            If UpDateType = 2 Then
                If IsDate(UpdateTime) = False Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>����¼��ʱ���ʽ����ȷ��</li>"
                Else
                    UpdateTime = CDate(UpdateTime)
                End If
            End If
        End If
        If FoundErr = True Then
            Call WriteErrMsg(ErrMsg, ComeUrl)
            Exit Sub
        End If

        sql = "Select * from PE_Item Where ItemID=" & ItemID
        Set rsItem = Server.CreateObject("adodb.recordset")
        rsItem.Open sql, Conn, 1, 3
        rsItem("PaginationType") = PaginationType
        rsItem("MaxCharPerPage") = MaxCharPerPage
        rsItem("CollectionNum") = CollectionNum
        rsItem("CollectionType") = CollectionType
        rsItem("InfoPoint") = InfoPoint
        If IncludePicYn = "yes" Then
            rsItem("IncludePicYn") = True
        Else
            rsItem("IncludePicYn") = False
        End If
        If DefaultPicYn = "yes" Then
            rsItem("DefaultPicYn") = True
        Else
            rsItem("DefaultPicYn") = False
        End If
        If OnTop = "yes" Then
            rsItem("OnTop") = True
        Else
            rsItem("OnTop") = False
        End If
        If Hot = "yes" Then
            rsItem("Hot") = True
        Else
            rsItem("Hot") = False
        End If
        If Elite = "yes" Then
            rsItem("Elite") = True
        Else
            rsItem("Elite") = False
        End If
        If Hits <> "" Then
            Hits = PE_CLng(Hits)
        Else
            Hits = 0
        End If
        If Stars = "" Then
            Stars = 0
        Else
            Stars = PE_CLng(Stars)
        End If

        rsItem("ChannelID") = ChannelID
        rsItem("ClassID") = ClassID
        rsItem("SpecialID") = SpecialID
        rsItem("Hits") = Hits
        rsItem("Stars") = Stars
        rsItem("UpdateType") = UpDateType
        If UpDateType = 2 Then
            rsItem("UpDateTime") = UpdateTime
        End If
        rsItem("SkinID") = SkinID
        rsItem("TemplateID") = TemplateID

        rsItem("InfoPurview") = InfoPurview
        rsItem("arrGroupID") = arrGroupID
        rsItem("InfoPoint") = InfoPoint
        rsItem("ChargeType") = ChargeType
        rsItem("PitchTime") = PitchTime
        rsItem("ReadTimes") = ReadTimes
        rsItem("DividePercent") = DividePercent

        If Script_Iframe = "yes" Then
            rsItem("Script_Iframe") = True
        Else
            rsItem("Script_Iframe") = False
        End If
        If Script_Object = "yes" Then
            rsItem("Script_Object") = True
        Else
            rsItem("Script_Object") = False
        End If
        If Script_Script = "yes" Then
            rsItem("Script_Script") = True
        Else
            rsItem("Script_Script") = False
        End If
        If Script_Class = "yes" Then
            rsItem("Script_Class") = True
        Else
            rsItem("Script_Class") = False
        End If
        If Script_Div = "yes" Then
            rsItem("Script_Div") = True
        Else
            rsItem("Script_Div") = False
        End If
        If Script_Span = "yes" Then
            rsItem("Script_Span") = True
        Else
            rsItem("Script_Span") = False
        End If
        If Script_Img = "yes" Then
            rsItem("Script_Img") = True
        Else
            rsItem("Script_Img") = False
        End If

        If Script_Font = "yes" Then
            rsItem("Script_Font") = True
        Else
            rsItem("Script_Font") = False
        End If
        If Script_A = "yes" Then
            rsItem("Script_A") = True
        Else
            rsItem("Script_A") = False
        End If
        If Script_Html = "yes" Then
            rsItem("Script_Html") = True
        Else
            rsItem("Script_Html") = False
        End If
           
        If Script_Table = "yes" Then
            rsItem("Script_Table") = True
        Else
            rsItem("Script_Table") = False
        End If
        
        If Script_Tr = "yes" Then
            rsItem("Script_Tr") = True
        Else
            rsItem("Script_Tr") = False
        End If
        
        If Script_Td = "yes" Then
            rsItem("Script_Td") = True
        Else
            rsItem("Script_Td") = False
        End If
                    
        If ShowCommentLink = "yes" Then
            rsItem("ShowCommentLink") = True
        Else
            rsItem("ShowCommentLink") = False
        End If
        
        If SaveFiles = "yes" Then
            rsItem("SaveFiles") = True
        Else
            rsItem("SaveFiles") = False
        End If

        If SaveFlashUrlToFile = "yes" Then
            rsItem("SaveFlashUrlToFile") = True
        Else
            rsItem("SaveFlashUrlToFile") = False
        End If
                
        If AddWatermark = "yes" Then
            rsItem("AddWatermark") = True
        Else
            rsItem("AddWatermark") = False
        End If
        If AddThumb = "yes" Then
            rsItem("AddThumb") = True
        Else
            rsItem("AddThumb") = False
        End If
        rsItem("CollecOrder") = CollecOrder
        rsItem("Status") = Status
        If CreateImmediate = "yes" Then
            rsItem("CreateImmediate") = True
        Else
            rsItem("CreateImmediate") = False
        End If
        If Trim(Request("Flag")) = "True" Then
            rsItem("Flag") = True
        End If
        rsItem.Update
        rsItem.Close
        Set rsItem = Nothing
    End If
    
    Call CloseConn
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    Else
        Call WriteSuccessMsg("<font color=red>" & ItemName & "</font>  �ɼ���Ŀ�ѳɹ����� ��", ComeUrl)
        Call Refresh("Admin_CollectionManage.asp?Action=Main",3)		
        'Response.Write "<meta http-equiv=""refresh"" content=3;url=""Admin_CollectionManage.asp?Action=Main"" >"
    End If
End Sub
'=================================================
'��������ItemCopy
'��  �ã�������Ŀ
'=================================================
Sub ItemCopy()
    Dim sql, rs, ItemID, FoundErr, ErrMsg, ItemName, ListStr
    Dim ClassID, SpecialID
    Dim CountItemNum, ItemNum, arrSpecialID
    FoundErr = False
    ItemID = PE_CLng(Trim(Request("ItemID")))
    If ItemID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����������Ŀ��ID���ԣ�</li>"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If
    CountItemNum = PE_CLng(Trim(Request("CountItemNum")))
    If CountItemNum <= 0 Then
        CountItemNum = 1
    End If
    Set rs = Conn.Execute("Select ItemID,ItemName,ChannelID,ClassID,SpecialID,ListStr from PE_Item Where ItemID=" & ItemID)
    If Not rs.EOF And Not rs.BOF Then
        ItemName = rs("ItemName")
        ListStr = rs("ListStr")
        ChannelID = rs("ChannelID")
        ClassID = rs("ClassID")
        SpecialID = rs("SpecialID")
    Else
        ErrMsg = "û�������Ŀ��ID��"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If
    Set rs = Nothing
    Response.Write "<script language=JavaScript>" & vbCrLf
    Response.Write "<!--" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "    for(var i=1,str="""";i<=" & CountItemNum & ";i++){" & vbCrLf
    Response.Write "        if(eval(""document.form1.ItemName""+i+"".value==''"")){" & vbCrLf
    Response.Write "            alert(""��Ŀ"" + i + ""����Ϊ�գ�"");" & vbCrLf
    Response.Write "            eval(""document.form1.ItemName""+i+"".focus()"");" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }   " & vbCrLf
    Response.Write "        if (eval(""document.form1.ListStr""+i+"".value==''"")){" & vbCrLf
    Response.Write "            alert(""�����б�"" + i + ""����Ϊ�գ�"");" & vbCrLf
    Response.Write "            eval(""document.form1.ListStr""+i+"".focus()"");" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }   " & vbCrLf
    Response.Write "        if (eval(""document.form1.ClassID""+i+"".value==''"")){" & vbCrLf
    Response.Write "            alert(""��Ŀ"" + i + ""������Ŀ����Ϊ��,���ƶ�Ϊ�ⲿ��Ŀ��"");" & vbCrLf
    Response.Write "            eval(""document.form1.ClassID""+i+"".focus()"");" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        if (eval(""document.form1.ClassID""+i+"".value=='0'"")){" & vbCrLf
    Response.Write "            alert(""��Ŀ"" + i + ""ָ������Ŀ������������£�ֻ������������Ŀ��������¡�"");" & vbCrLf
    Response.Write "            eval(""document.form1.ClassID""+i+"".focus()"");" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "  <form name=form1 action=Admin_CollectionManage.asp method=post >" & vbCrLf
    Response.Write "  <table class=border cellSpacing=1 cellPadding=0 width=""100%"" align=center border=0>" & vbCrLf
    Response.Write "    <tr>" & vbCrLf
    Response.Write "      <td class=title colSpan=2 height=22>" & vbCrLf
    Response.Write "      <div align=center><STRONG>�� Ŀ �� ��</STRONG>" & vbCrLf
    Response.Write "      </div></TD>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class=tdbg>" & vbCrLf
    Response.Write "      <td class=tdbg align=left>&nbsp;&nbsp;<STRONG>��ѡ��Ҫ���Ƶ���Ŀ���� </STRONG>" & vbCrLf
    Response.Write "       <Select  name='CountItemNum' onchange=""javascript:window.location='Admin_CollectionManage.asp?Action=ItemCopy&ItemID=" & ItemID & "&CountItemNum='+this.options[this.selectedIndex].value;"">" & vbCrLf
    Response.Write "         <option value=1 "
    If CountItemNum = 1 Then Response.Write "selected"
    Response.Write ">1</OPTION>" & vbCrLf
    Response.Write "         <option value=2 "
    If CountItemNum = 2 Then Response.Write "selected"
    Response.Write ">2</OPTION>" & vbCrLf
    Response.Write "         <option value=3 "
    If CountItemNum = 3 Then Response.Write "selected"
    Response.Write ">3</OPTION>" & vbCrLf
    Response.Write "         <option value=4 "
    If CountItemNum = 4 Then Response.Write "selected"
    Response.Write ">4</OPTION>" & vbCrLf
    Response.Write "         <option value=5 "
    If CountItemNum = 5 Then Response.Write "selected"
    Response.Write ">5</OPTION>" & vbCrLf
    Response.Write "         <option value=6 "
    If CountItemNum = 6 Then Response.Write "selected"
    Response.Write ">6</OPTION>" & vbCrLf
    Response.Write "         <option value=7 "
    If CountItemNum = 7 Then Response.Write "selected"
    Response.Write ">7</OPTION>" & vbCrLf
    Response.Write "         <option value=8 "
    If CountItemNum = 8 Then Response.Write "selected"
    Response.Write ">8</OPTION>" & vbCrLf
    Response.Write "         <option value=9 "
    If CountItemNum = 9 Then Response.Write "selected"
    Response.Write ">9</OPTION>" & vbCrLf
    Response.Write "         <option value=10 "
    If CountItemNum = 10 Then Response.Write "selected"
    Response.Write ">10</OPTION>" & vbCrLf
    Response.Write "       </Select></TD>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class=tdbg>" & vbCrLf
    Response.Write "      <td class=tdbg align=middle colSpan=2 height=50>" & vbCrLf
    For ItemNum = 1 To CountItemNum
        Response.Write "  <table class=border cellSpacing=2 cellPadding=1 width='95%' align=center border=0>" & vbCrLf
        Response.Write "    <tr class='title'>" & vbCrLf
        Response.Write "      <td align='left' colspan='2'><b>������Ŀ" & ItemNum & "</b></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='150' class='tdbg5' align='right'>��Ŀ���ƣ�</td>" & vbCrLf
        Response.Write "      <td class='tdbg'>" & vbCrLf
        Response.Write "       <Input name='ItemName" & ItemNum & "' type='text' id='ItemName' size='30' maxlength='30' value='" & ItemName & " ����'><font color=red> * </font></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='150' class='tdbg5' align='right'>�б�ҳURL��</td>" & vbCrLf
        Response.Write "      <td class='tdbg'>" & vbCrLf
        Response.Write "       <Input name='ListStr" & ItemNum & "' type='text' id='ListStr' size='49' maxlength='150' value='" & ListStr & "'> <font color=red> * </font><font color=blue>'ע����Ҫ�ɼ���վ���б�ҳ</font></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class=""tdbg""> " & vbCrLf
        Response.Write "      <td width=""150"" class=""tdbg5"" align=""right"">������Ŀ��</td>" & vbCrLf
        Response.Write "      <td class=""tdbg"">" & vbCrLf
        Response.Write "        <select name='ClassID" & ItemNum & "'>" & vbCrLf
        Response.Write GetClass_Option(3, ClassID)
        Response.Write "        </select> " & vbCrLf
        Response.Write "        &nbsp;</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class=""tdbg""> " & vbCrLf
        Response.Write "      <td width=""150"" class=""tdbg5"" align=""right"">����ר�⣺</td>" & vbCrLf
        Response.Write "      <td class=""tdbg"">" & vbCrLf
        Response.Write "      <select name='SpecialID" & ItemNum & "'>" & vbCrLf
        Response.Write GetSpecial_Option(arrSpecialID)
        Response.Write "      </select> " & vbCrLf
        Response.Write "        &nbsp;</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </table>" & vbCrLf
    Next
    Response.Write "      </TD>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class=tdbg>" & vbCrLf
    Response.Write "      <td class=tdbg align=middle colSpan=2 height=50>" & vbCrLf
    Response.Write "       <input name='ItemID' value='" & ItemID & "' type='hidden'>" & vbCrLf
    Response.Write "       <input name='Action' value='DoItemCopy' type='hidden'>" & vbCrLf
    Response.Write "       <Input style=""CURSOR: hand; BACKGROUND-COLOR: #ffffff""  type=submit value="" ��ʼ���� "" name=Submit onClick='return CheckForm()'></TD>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "  </form>" & vbCrLf
End Sub
'=================================================
'��������DoItemCopy
'��  �ã�������Ŀ����
'=================================================
Sub DoItemCopy()
    Dim sql, rsItem, rsFilters, ItemID, FoundErr, ErrMsg, Arr_Item, Arr_i, ItemName, Arr_Filter, CountItemNum, ItemNum
    FoundErr = False
    ItemID = Trim(Request("ItemID"))
    CountItemNum = Trim(Request("CountItemNum"))
    If ItemID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��������,��Ŀ��ID���ԣ�</li>"
    Else
        ItemID = CLng(ItemID)
    End If
    If CountItemNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��Ŀ�ĸ������ԣ�</li>"
    Else
        CountItemNum = CLng(CountItemNum)
    End If
    If FoundErr <> True Then
        Set rsItem = Conn.Execute("Select * from PE_Item Where ItemID=" & ItemID)
        If Not rsItem.EOF And Not rsItem.BOF Then
            Arr_Item = rsItem.GetRows()
        Else
            Arr_Item = ""
        End If
        Set rsItem = Nothing
        If IsArray(Arr_Item) = True Then
            Set rsFilters = Conn.Execute("Select * from PE_Filters Where ItemID=" & ItemID)
            If Not rsFilters.EOF And Not rsFilters.BOF Then
                Arr_Filter = rsFilters.GetRows()
            Else
                Arr_Filter = ""
            End If
            Set rsFilters = Nothing
        End If
        If IsArray(Arr_Item) = True Then
            Set rsItem = Server.CreateObject("adodb.recordset")
            sql = "select top 1 * from PE_Item"
            rsItem.Open sql, Conn, 1, 3
            Set rsFilters = Server.CreateObject("adodb.recordset")
            sql = "select top 1 * from PE_Filters"
            rsFilters.Open sql, Conn, 1, 3
            For ItemNum = 1 To CountItemNum
                If Trim(Request("ItemName" & ItemNum & "")) <> "" And Trim(Request("ListStr" & ItemNum & "")) <> "" Then
                    rsItem.addnew
                    rsItem(1) = Trim(Request("ItemName" & ItemNum & ""))
                    ItemName = Arr_Item(1, 0)
                    For Arr_i = 2 To UBound(Arr_Item, 1)
                        If Arr_i = 14 Then
                            rsItem(14) = Trim(Request("ListStr" & ItemNum & ""))
                        ElseIf Arr_i = 3 Then
                            rsItem(3) = Trim(Request("ClassID" & ItemNum & ""))
                        ElseIf Arr_i = 4 Then
                            rsItem(4) = Trim(Request("SpecialID" & ItemNum & ""))
                        Else
                            rsItem(Arr_i) = Arr_Item(Arr_i, 0)
                        End If
                    Next
                    If SystemDatabaseType <> "SQL" Then
                        ItemID = rsItem("ItemID")
                    End If
                    rsItem.Update
                    ErrMsg = ErrMsg & "<br>�µ���Ŀ����Ϊ��<font color=red>" & Trim(Request("ItemName" & ItemNum & "")) & "</font>"
                    If SystemDatabaseType = "SQL" Then
                        Dim mrs
                        Set mrs = Conn.Execute("select max(ItemID) from PE_Item")
                        If IsNull(mrs(0)) Then
                            ItemID = 1
                        Else
                            ItemID = mrs(0)
                        End If
                        Set mrs = Nothing
                    End If
                    If IsArray(Arr_Filter) = True Then
                        rsFilters.addnew
                        rsFilters(1) = ItemID
                        For Arr_i = 2 To UBound(Arr_Filter, 1)
                            rsFilters(Arr_i) = Arr_Filter(Arr_i, 0)
                        Next
                        rsFilters.Update
                    End If
                End If
            Next
            rsItem.Close
            Set rsItem = Nothing
            rsFilters.Close
            Set rsFilters = Nothing
        Else
            FoundErr = True
            ErrMsg = ErrMsg & "��������,û���ҵ�����Ŀ"
        End If
    End If
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    Else
        Response.Write "<br>"
        Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
        Response.Write "  <tr align='center' class='title'><td height='22'><strong>��ϲ����</strong></td></tr>" & vbCrLf
        Response.Write "  <tr class='tdbg'><td height='100' valign='top'><br>" & ItemName & " ��Ŀ�������." & ErrMsg & "<br></td></tr>" & vbCrLf
        Response.Write "  <tr align='center' class='tdbg'><td>"
        Response.Write "</td></tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
        Call Refresh("Admin_CollectionManage.asp?Action=ItemManage",5)				
        'Response.Write "<meta http-equiv=""refresh"" content=5;url=""Admin_CollectionManage.asp?Action=ItemManage"" >"
    End If
    Call CloseConn
End Sub
'=================================================
'��������Batch
'��  �ã�����������Ŀ����
'=================================================
Sub Batch()

    Dim ChannelShortName, arrGroupID, strDisabled
    ChannelShortName = "����"

    ItemID = Replace(Trim(Request("ItemID")), " ", "")
    Call PopCalendarInit
    Response.Write "<script language=""JavaScript"">" & vbCrLf
    Response.Write "<!--" & vbCrLf
    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "// ֻ������������" & vbCrLf
    Response.Write "function IsDigit(){" & vbCrLf
    Response.Write "    return ((event.keyCode >= 48) && (event.keyCode <= 57));" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "    for(var i=0;i<document.form1.BatchItemID.length;i++){" & vbCrLf
    Response.Write "    document.form1.BatchItemID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "    for(var i=0;i<document.form1.BatchItemID.length;i++){" & vbCrLf
    Response.Write "    document.form1.BatchItemID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<form method=""post"" action=""Admin_CollectionManage.asp"" name=""form1"" >" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0' class='border'>" & vbCrLf
    Response.Write "    <tr class='title'>" & vbCrLf
    Response.Write "      <td height='22' colspan='2' align='center'><b>�����޸���Ŀ����</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr align='center'>" & vbCrLf
    Response.Write "      <td class='tdbg' valign='top'>" & vbCrLf
    Response.Write "        <table width='100%' border='0' cellpadding='2' cellspacing='1'>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td align='center'><b>���÷�Χ��</b></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td>" & vbCrLf
    Dim SqlI, RsI
    SqlI = "select ItemID,ItemName from PE_Item order by ItemID desc"
    Set RsI = Server.CreateObject("adodb.recordset")
    RsI.Open SqlI, Conn, 1, 1
    Response.Write "<select name='BatchItemID' size='2' multiple style='height:290px;width:180px;'>"
    If RsI.EOF And RsI.BOF Then
        Response.Write "<option value=""0"">�������Ŀ</option>"
    Else
        Do While Not RsI.EOF
            Response.Write "<option value=" & RsI("ItemID") & " " & vbCrLf
            If ItemID <> "" Then
                If InStr(ItemID, ",") > 0 Then
                    ItemID = ReplaceBadChar(ItemID)
                    If FoundInArr(ItemID, RsI("ItemID"), ",") = True Then Response.Write "selected"
                Else
                    ItemID = CLng(Trim(ItemID))
                    If RsI("ItemID") = ItemID Then Response.Write "selected"
                End If
            End If
            Response.Write " >" & RsI("ItemName") & "</option>" & vbCrLf
            RsI.MoveNext
        Loop
    End If
    Response.Write "</select>"
    RsI.Close
    Set RsI = Nothing
    Response.Write "<br><b>&nbsp;�� Ctrl �� Shift ���ɶ�ѡ</b>" & vbCrLf
    Response.Write "      <br><div align='center'>" & vbCrLf
    Response.Write "      <input type='button' name='Submit' value='  ѡ��������Ŀ  ' onclick='SelectAll()'><br>" & vbCrLf
    Response.Write "      <input type='button' name='Submit' value='ȡ��ѡ��������Ŀ' onclick='UnSelectAll()'></div>" & vbCrLf
    Response.Write "            </td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td valign='top'>" & vbCrLf
    Response.Write "     <br>" & vbCrLf
    Response.Write "    <table width='100%' border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "     <tr align='center' height='24'>" & vbCrLf
    Response.Write "      <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>��������</td>" & vbCrLf
    Response.Write "      <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>�շ�����</td>" & vbCrLf
    Response.Write "      <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>�ɼ�����</td>" & vbCrLf
    Response.Write "      <td>&nbsp;</td>" & vbCrLf
    Response.Write "     </tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    Response.Write "    <table width='100%' border='0' cellpadding='2' cellspacing='1'  class='border'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "         <tr class='tdbg' id='ArticleContent2' style=""display:''""> "
    Response.Write "           <td width='30' align='center' class=""tdbg5""><input type='checkbox' name='ModifyPageFrom' value='yes'></td>"
    Response.Write "           <td width='120'  class='tdbg5' align=""right"">���ݷ�ҳ��ʽ��&nbsp;</td>"
    Response.Write "           <td>"
    Response.Write "             <select name='PaginationType' id='PaginationType'>"
    Response.Write "               <option value='2'>�ֶ���ҳ</option>"
    Response.Write "               <option value='1'>�Զ���ҳ</option>"
    Response.Write "               <option value='0'>����ҳ</option>"
    Response.Write "            </select>"
    Response.Write "            �Զ���ҳʱ��ÿҳ��Լ�ַ���������HTML��ǣ��� <input name='MaxCharPerPage' type='text' id='MaxCharPerPage' value='' size='8' maxlength='8'>"
    Response.Write "           </td>"
    Response.Write "         </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "           <td width='30' align='center' class=""tdbg5""><input type='checkbox' name='ModifyArticleAttributeFrom' value='yes'></td>"
    Response.Write "           <td width='120'  class='tdbg5' align=""right"">�������ԣ�&nbsp;</td>"
    Response.Write "           <td>"
    Response.Write "             <input name=""IncludePicYn"" type=""checkbox"" value=""yes"" >����ͼƬ" & vbCrLf
    Response.Write "             <input name=""DefaultPicYn"" type=""checkbox"" value=""yes"" >��ҳͼƬ" & vbCrLf
    Response.Write "             <input name='OnTop' type='checkbox' id='OnTop' value='yes'> �̶�����&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "             <input name='Elite' type='checkbox' id='Elite' value='yes'> �Ƽ�����&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "             <br>�������ֵȼ�&nbsp;<select name='Stars' id='Stars'>" & GetStars(5) & "</select>"
    Response.Write "           </td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>"
    Response.Write "           <td width='30' align='center' class=""tdbg5""><input type='checkbox' name='ModifyHistFrom' value='yes'></td>"
    Response.Write "           <td width='120'  class='tdbg5' align=""right"">�������ʼֵ��&nbsp;</td>"
    Response.Write "           <td><input name='Hits' type='text' id='Hits' value='' size='10' maxlength='10' ONKEYPRESS=""event.returnValue=IsDigit();"">&nbsp;&nbsp; <font color='#0000FF'>�⹦�����ṩ������Ա�����õġ�����������Ҫ��ѽ��^_^</font></td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class=""tdbg"">" & vbCrLf
    Response.Write "           <td width='30' align='center' class=""tdbg5""><input type='checkbox' name='ModifyTimeFrom' value='yes'></td>"
    Response.Write "           <td width='120'  class='tdbg5' align=""right"">����¼��ʱ�䣺&nbsp;</td>" & vbCrLf
    Response.Write "           <td><Input name=""UpdateType"" type=""radio"" value=""0"" >��ǰʱ��" & vbCrLf
    Response.Write "               <Input name=""UpdateType"" type=""radio"" value=""1"" >��ǩ�е�ʱ��" & vbCrLf
    Response.Write "               <Input name=""UpdateType"" type=""radio"" value=""2"" >�Զ��壺" & vbCrLf
    Response.Write "               <Input name='UpdateTime' type='text' size='20' maxlength='20' value='' maxlength='50' onClick='PopCalendar.show(document.form1.UpdateTime, ""yyyy-mm-dd"", null, null, null, ""11"");'><a style='cursor:hand;' onClick='PopCalendar.show(document.form1.UpdateTime, ""yyyy-mm-dd"", null, null, null, ""11"");'><img src='Images/Calendar.gif' border='0' Style='Padding-Top:10px' align='absmiddle'></a>"
    Response.Write "           </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class=""tdbg"">" & vbCrLf
    Response.Write "           <td width='30' align='center' class=""tdbg5""><input type='checkbox' name='ModifyCommentFrom' value='yes'></td>"
    Response.Write "           <td width='120'  class='tdbg5' align=""right"">�������ӣ�&nbsp;</td>" & vbCrLf
    Response.Write "           <td>" & vbCrLf
    Response.Write "               <Input name=""ShowCommentLink"" type=""radio"" id=""ShowCommentLink"" value=""yes"" >��ʾ��������  " & vbCrLf
    Response.Write "               <Input name=""ShowCommentLink"" type=""radio"" id=""ShowCommentLink"" value=""no"" >����ʾ��������" & vbCrLf
    Response.Write "           </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "           <td width='30' align='center' class=""tdbg5""><input type='checkbox' name='ModifySkinIDFrom' value='yes'></td>"
    Response.Write "           <td width='120'  class='tdbg5' align=""right"">��ɫ���&nbsp;</td>"
    Response.Write "           <td><select Name='SkinID'>" & GetSkin_Option(0) & "</select>&nbsp;���ģ���а���CSS����ɫ��ͼƬ����Ϣ</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'>" & vbCrLf
    Response.Write "          <td width='30' align='center' class=""tdbg5""><input type='checkbox' name='ModifyItemEstate' value='yes'></td>"
    Response.Write "          <td width='120' align='right' class='tdbg5' align=""right"">����״̬��&nbsp;</td>"
    Response.Write "          <td>" & vbCrLf
    Response.Write "            <input name='Status' type='radio' id='Status' value='-1' >�ݸ�&nbsp;&nbsp;" & vbCrLf
    Response.Write "            <input Name='Status' Type='Radio' Id='Status' Value='0' >�����&nbsp;&nbsp;" & vbCrLf
    Response.Write "            <input Name='Status' Type='Radio' Id='Status' Value='3' checked> ����ͨ��" & vbCrLf
    Response.Write "            &nbsp;&nbsp;<input name=""CreateImmediate"" type=""checkbox"" id=""CreateImmediate"" value=""yes"" >�������� <font color=blue>ע�� ������Ҫ�ǵ�������Ӧ��JS�ļ�</font>" & vbCrLf
    Response.Write "           </td>"
    Response.Write "        </tr>"
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "           <td width='30' align='center' class=""tdbg5""><input type='checkbox' name='ModifyInfoPurview' value='yes'></td>"
    Response.Write "            <td width='120'  class='tdbg5' align=""right"">�Ķ�Ȩ�ޣ�&nbsp;</td>"
    Response.Write "            <td><input name='InfoPurview' type='radio' value='0'>�̳���ĿȨ�ޣ���������ĿΪ��֤��Ŀʱ������ѡ����<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='1'>���л�Ա����������ĿΪ������Ŀ���뵥����ĳЩ���½��в鿴Ȩ�����ã�����ѡ����<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='2'>ָ����Ա�飨��������ĿΪ������Ŀ���뵥����ĳЩ���½��в鿴Ȩ�����ã�����ѡ����<br>"
    Response.Write GetUserGroup(arrGroupID & "", strDisabled)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class=""tdbg5""><input type='checkbox' name='ModifyInfoPoint' value='yes'></td>"
    Response.Write "            <td width='120' align='right' class='tdbg5'> " & ChannelShortName & "�Ķ�������&nbsp; </td>"
    Response.Write "            <td><input name='InfoPoint' type='text' id='InfoPoint' value='' size='5' maxlength='4' style='text-align:center'>&nbsp;&nbsp;&nbsp;&nbsp; <font color='#0000FF'>�������0�����Ա�Ķ���" & ChannelShortName & "ʱ��������Ӧ��������Ϊ9999ʱ���⣩���οͽ��޷��鿴��" & ChannelShortName & "��</font></td>"
    Response.Write "          </tr>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "           <td width='30' align='center' class=""tdbg5""><input type='checkbox' name='ModifyChargeType' value='yes'></td>"
    Response.Write "            <td width='120'  class='tdbg5' align=""right"">�ظ��շѣ�&nbsp; </td>"
    Response.Write "            <td><input name='ChargeType' type='radio' value='0'>���ظ��շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='1'>�����ϴ��շ�ʱ�� <input name='PitchTime' type='text' value='' size='8' maxlength='8' style='text-align:center'" & strDisabled & "> Сʱ�������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='2'>��Ա�ظ��鿴������ <input name='ReadTimes' type='text' value='' size='8' maxlength='8' style='text-align:center'> �κ������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='3'>�������߶�����ʱ�����շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='4'>����������һ������ʱ�������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='5'>ÿ�Ķ�һ�ξ��ظ��շ�һ�Σ����鲻Ҫʹ�ã�"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "           <td width='30' align='center' class=""tdbg5""><input type='checkbox' name='ModifyDividePercent' value='yes'></td>"
    Response.Write "           <td width='120' class='tdbg5' align=""right"">�ֳɱ�����&nbsp; </td>"
    Response.Write "            <td><input name='DividePercent' type='text' id='DividePercent' value='' size='5' maxlength='4' style='text-align:center'> %"
    Response.Write "              &nbsp;&nbsp;<font color='#0000FF'>�����������0���򽫰����������Ķ�����ȡ�ĵ���֧����¼����</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "           <td width='30' align='center' class=""tdbg5""><input type='checkbox' name='ModifyScriptFrom' value='yes'></td>"
    Response.Write "           <td width='120'  class='tdbg5' align=""right"">����ѡ�&nbsp;</td>"
    Response.Write "           <td height=""22""><input name=""Script_Iframe"" type=""checkbox"" id=""Script_Iframe""  value=""yes"" >IFRAME" & vbCrLf
    Response.Write "               <input name=""Script_Object"" type=""checkbox"" id=""Script_Object""  value=""yes"" >Object" & vbCrLf
    Response.Write "               <input name=""Script_Script"" type=""checkbox"" id=""Script_Script""  value=""yes"" >Script" & vbCrLf
    Response.Write "               <input name=""Script_Class"" type=""checkbox"" id=""Script_Class""  value=""yes"" >Style" & vbCrLf
    Response.Write "               <input name=""Script_Div"" type=""checkbox"" id=""Script_Div""  value=""yes"" >Div" & vbCrLf
    Response.Write "               <input name=""Script_Table"" type=""checkbox"" id=""Script_Table""  value=""yes"" >Table" & vbCrLf
    Response.Write "               <input name=""Script_Tr"" type=""checkbox"" id=""Script_tr""  value=""yes"" >Tr" & vbCrLf
    Response.Write "               <input name=""Script_td"" type=""checkbox"" id=""Script_td""  value=""yes"" >Td" & vbCrLf
    Response.Write "               <br>" & vbCrLf
    Response.Write "               <input name=""Script_Span"" type=""checkbox"" id=""Script_Span""  value=""yes"" >Span" & vbCrLf
    Response.Write "               <input name=""Script_Img"" type=""checkbox"" id=""Script_Img""  value=""yes"" >Img&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "               <input name=""Script_Font"" type=""checkbox"" id=""Script_Font""  value=""yes"" >FONT&nbsp;&nbsp;" & vbCrLf
    Response.Write "               <input name=""Script_A"" type=""checkbox"" id=""Script_A""  value=""yes"" >A&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "               <input name=""Script_Html"" type=""checkbox"" id=""Script_Html""  value=""yes"">Html" & vbCrLf
    Response.Write "           </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf

    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='30' align='center' class=""tdbg5""><input type='checkbox' name='ModifyCollectionNumFrom' value='yes'></td>"
    Response.Write "          <td width='120' class='tdbg5' align=""right"">�ɼ�������&nbsp;</td>"
    Response.Write "          <td  height=""30"" class=""tdbg"">" & vbCrLf
    Response.Write "            <Input type='radio' Name='iType' value='0' checked> �ɼ��б��е���������  <br>" & vbCrLf
    Response.Write "            <Input type='radio' Name='iType' value='1'> �ɼ��б��е� <Input TYPE='text' Name='AritcleNum' value='30' size='3' maxlength='5' ONKEYPRESS=""event.returnValue=IsDigit();"">ƪ���º�ֹͣ�ɼ� <br>" & vbCrLf
    Response.Write "            <Input type='radio' Name='iType' value='2'> �ɼ��б��е� <Input TYPE='text' Name='PageNum' value='5' size='3' maxlength='5' ONKEYPRESS=""event.returnValue=IsDigit();"">����ҳ��ֹͣ�ɼ� <br>" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='30' align='center' class=""tdbg5""><input type='checkbox' name='ModifyCollectionImageFrom' value='yes'></td>"
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right"">�ɼ�ͼƬ���ã�&nbsp;</td>"
    Response.Write "          <td  height=""30"" class=""tdbg"">" & vbCrLf
    Response.Write "           <input name=""SaveFiles"" type=""checkbox"" id=""SaveFiles"" value=""yes"" checked>����Զ��ͼƬ" & vbCrLf
    Response.Write "           <input name=""AddWatermark"" type=""checkbox"" value=""yes"" >�Զ���ͼƬ����ˮӡ" & vbCrLf
    Response.Write "           <input name=""AddThumb"" type=""checkbox"" value=""yes"" >�Զ�Ϊ��һ��ͼƬ��������ͼ<br>" & vbCrLf
    Response.Write "           <input name=""SaveFlashUrlToFile"" type=""checkbox"" value=""yes"" checked>�����������е�Flash��ͼƬ�ĵ�ַ���浽��Ŀ¼�е�CollectionFilePath.txt�ļ��У��Է������ʿ쳵�������������" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf

    Response.Write "        <tr class='tdbg'>"
    Response.Write "          <td width='30' align='center' class=""tdbg5""><input type='checkbox' name='ModifyCollectionCompositorFrom' value='yes'></td>"
    Response.Write "          <td width=""120"" class=""tdbg5"" align=""right"">���²ɼ�˳��&nbsp;</td>"
    Response.Write "          <td  height=""30"" class=""tdbg"">" & vbCrLf
    Response.Write "            <Input type='radio' Name='CollecOrder' value='0' >����ɼ�"
    Response.Write "            <Input type='radio' Name='CollecOrder' value='1' >����ɼ� <FONT color='blue'>���Ƽ���</FONT>"
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        <tr class=""tdbg""> " & vbCrLf
    Response.Write "          <td colspan=""4"" align=""center"" class=""tdbg"" height=""50"">" & vbCrLf
    Response.Write "            <input name=""Action"" type=""hidden"" id=""Action"" value=""DoBatch"">" & vbCrLf
    Response.Write "            <input name=""Cancel"" type=""button"" id=""Cancel"" value=""������һ��"" onClick=""window.location.href='javascript:history.go(-1)'"" >&nbsp;&nbsp;" & vbCrLf
    Response.Write "            <input  type=""submit"" name=""Submit"" value="" ��  �� "" >" & vbCrLf
    Response.Write "          </td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    Response.Write "  </td>" & vbCrLf
    Response.Write " </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    
    Call CloseConn
End Sub
'=================================================
'��������DoBatch
'��  �ã���������������Ŀ����
'=================================================
Sub DoBatch()
    
    Dim rs, sql, BatchType, BatchItemID
    '�������Ա���
    Dim ClassID, SpecialID, PaginationType, MaxCharPerPage, InfoPoint
    Dim OnTop, Hot, Elite, Hits, Stars, UpDateType, UpdateTime, IncludePicYn, DefaultPicYn, SkinID, TemplateID
    Dim UploadDir, UpFileType
    '���˱���
    Dim Script_Iframe, Script_Object, Script_Script, Script_Class
    Dim Script_Div, Script_Span, Script_Img, Script_Font, Script_A, Script_Html
    Dim Script_Table, Script_Tr, Script_Td, ShowCommentLink
    '�ɼ�����
    Dim SaveFiles, AddWatermark, AddThumb, CollecOrder, SaveFlashUrlToFile, Status, iType, CreateImmediate, CollectionNum, CollectionType
    '�շѱ���
    Dim InfoPurview, arrGroupID, ChargeType, DividePercent, PitchTime, ReadTimes

    FoundErr = False
    ComeUrl = "Admin_CollectionManage.asp?Action=Batch"
    
    BatchType = PE_CLng(Trim(Request("BatchType")))
    BatchItemID = Trim(Request.Form("BatchItemID"))
          
    ChannelID = Trim(Request.Form("ChannelID"))
    ClassID = Trim(Request.Form("ClassID"))
    SpecialID = Trim(Request.Form("SpecialID"))
    PaginationType = Trim(Request.Form("PaginationType"))
    MaxCharPerPage = Trim(Request.Form("MaxCharPerPage"))
    InfoPoint = Trim(Request.Form("InfoPoint"))
    OnTop = Trim(Request.Form("OnTop"))
    Hot = Trim(Request.Form("Hot"))
    Elite = Trim(Request.Form("Elite"))
    Hits = Trim(Request.Form("Hits"))
    Stars = Trim(Request.Form("Stars"))
    UpDateType = Trim(Request.Form("UpdateType"))
    UpdateTime = Trim(Request.Form("UpdateTime"))
    SkinID = Trim(Request.Form("SkinID"))
    TemplateID = Trim(Request.Form("TemplateID"))
    IncludePicYn = Trim(Request.Form("IncludePicYn"))
    DefaultPicYn = Trim(Request.Form("DefaultPicYn"))

    Script_Iframe = Trim(Request.Form("Script_Iframe"))
    Script_Object = Trim(Request.Form("Script_Object"))
    Script_Script = Trim(Request.Form("Script_Script"))
    Script_Class = Trim(Request.Form("Script_Class"))
    Script_Div = Trim(Request.Form("Script_Div"))
    Script_Span = Trim(Request.Form("Script_Span"))
    Script_Img = Trim(Request.Form("Script_Img"))
    Script_Font = Trim(Request.Form("Script_Font"))
    Script_A = Trim(Request.Form("Script_A"))
    Script_Html = Trim(Request.Form("Script_Html"))
    Script_Table = Trim(Request.Form("Script_Table"))
    Script_Tr = Trim(Request.Form("Script_Tr"))
    Script_Td = Trim(Request.Form("Script_Td"))
    
    ShowCommentLink = Trim(Request("ShowCommentLink"))
    SaveFiles = Trim(Request.Form("SaveFiles"))
    AddWatermark = Trim(Request.Form("AddWatermark"))
    AddThumb = Trim(Request.Form("AddThumb"))
    CollecOrder = PE_CLng(Trim(Request.Form("CollecOrder")))
    SaveFlashUrlToFile = Trim(Request.Form("SaveFlashUrlToFile"))
    Status = PE_CLng(Trim(Request.Form("Status")))
    CreateImmediate = Trim(Request.Form("CreateImmediate"))
    CollectionNum = Trim(Request.Form("CollectionNum"))

    InfoPurview = PE_CLng(Trim(Request.Form("InfoPurview")))
    arrGroupID = Trim(Request.Form("GroupID"))
    ChargeType = PE_CLng(Trim(Request.Form("ChargeType")))
    DividePercent = PE_CLng(Trim(Request.Form("DividePercent")))
    PitchTime = PE_CLng(Trim(Request.Form("PitchTime")))
    ReadTimes = PE_CLng(Trim(Request.Form("ReadTimes")))
    iType = PE_CLng(Trim(Request.Form("iType")))
    Select Case iType
    Case 0  '�ɼ���������
      '����д����
        CollectionType = 0
    Case 1
        CollectionType = 0
        CollectionNum = PE_CLng(Trim(Request("AritcleNum")))
    Case 2
        CollectionType = 1
        CollectionNum = PE_CLng(Trim(Request("PageNum")))
    End Select
    
    If IsValidID(BatchItemID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����޸ĵ���Ŀ��ID</li>"
    End If

    If Trim(Request("ModifyChannelIDFrom")) = "yes" Then
        If ChannelID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>���������Ƶ��</li>"
        Else
            ChannelID = PE_CLng(ChannelID)
        End If
    End If
    If Trim(Request("ModifyClassIDFrom")) = "yes" Then
        If ClassID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>δָ����Ŀ������Ŀ����ָ������Ŀ����������Ŀ</li>"
        Else
            ClassID = PE_CLng(ClassID)
            If ClassID <= 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>ָ���˷Ƿ�����Ŀ���ⲿ��Ŀ�򲻴��ڵ���Ŀ��</li>"
            End If
        End If
    End If
    If Trim(Request("ModifySpecialIDFrom")) = "yes" Then
        If SpecialID = "" Then
            SpecialID = 0
        Else
            SpecialID = CLng(SpecialID)
        End If
    End If
    If Trim(Request("ModifyTemplateIDFrom")) = "yes" Then
        If TemplateID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ���������ģ��</li>"
        Else
            TemplateID = PE_CLng(TemplateID)
        End If
    End If
    
    If Trim(Request("ModifyPageFrom")) = "yes" Then
        If PaginationType = "" Then
            PaginationType = 0
        Else
            PaginationType = PE_CLng(PaginationType)
        End If
        
        If MaxCharPerPage = "" Then
            MaxCharPerPage = 0
        Else
            MaxCharPerPage = PE_CLng(MaxCharPerPage)
        End If
        
        If PaginationType = 1 And MaxCharPerPage = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ���Զ���ҳʱ��ÿҳ��Լ�ַ���,�������0</li>"
        End If
    End If

    If Trim(Request("ModifyArticleAttributeFrom")) = "yes" Then
        If IncludePicYn = "yes" Then
            IncludePicYn = True
        Else
            IncludePicYn = False
        End If
        If DefaultPicYn = "yes" Then
            DefaultPicYn = True
        Else
            DefaultPicYn = False
        End If
        If OnTop = "yes" Then
            OnTop = True
        Else
            OnTop = False
        End If
        If Elite = "yes" Then
            Elite = True
        Else
            Elite = False
        End If
        If Stars = "" Then
            Stars = 0
        Else
            Stars = PE_CLng(Stars)
        End If
    End If
    If Trim(Request("ModifyHistFrom")) = "yes" Then
        If Hits <> "" Then
            Hits = PE_CLng(Hits)
        Else
            Hits = 0
        End If
    End If
    If Trim(Request("ModifyTimeFrom")) = "yes" Then
        If UpDateType = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ѡ������¼��ʱ�����ͣ�</li>"
        Else
            UpDateType = CLng(UpDateType)
            If UpDateType = 2 Then
                If IsDate(UpdateTime) = False Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>����¼��ʱ���ʽ����ȷ��</li>"
                Else
                    UpdateTime = CDate(UpdateTime)
                End If
            End If
        End If
    End If
    If Trim(Request("ModifyCommentFrom")) = "yes" Then
        If ShowCommentLink = "yes" Then
            ShowCommentLink = True
        Else
            ShowCommentLink = False
        End If
    End If
    
    If Trim(Request("ModifySkinIDFrom")) = "yes" Then
        If SkinID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ����ɫ���</li>"
        Else
            SkinID = PE_CLng(SkinID)
        End If
    End If
    If Trim(Request("ModifyInfoPoint")) = "yes" Then
        If InfoPoint = "" Then
            InfoPoint = 0
        Else
            InfoPoint = PE_CLng(InfoPoint)
        End If
    End If
    If Trim(Request("ModifyScriptFrom")) = "yes" Then
        If Script_Iframe = "yes" Then
            Script_Iframe = True
        Else
            Script_Iframe = False
        End If
        If Script_Object = "yes" Then
            Script_Object = True
        Else
            Script_Object = False
        End If
        If Script_Script = "yes" Then
            Script_Script = True
        Else
            Script_Script = False
        End If
        If Script_Class = "yes" Then
            Script_Class = True
        Else
            Script_Class = False
        End If
        If Script_Div = "yes" Then
            Script_Div = True
        Else
            Script_Div = False
        End If
        If Script_Span = "yes" Then
            Script_Span = True
        Else
            Script_Span = False
        End If
        If Script_Img = "yes" Then
            Script_Img = True
        Else
            Script_Img = False
        End If
        If Script_Font = "yes" Then
            Script_Font = True
        Else
            Script_Font = False
        End If
        If Script_A = "yes" Then
            Script_A = True
        Else
            Script_A = False
        End If
        If Script_Html = "yes" Then
            Script_Html = True
        Else
            Script_Html = False
        End If
        If Script_Table = "yes" Then
            Script_Table = True
        Else
            Script_Table = False
        End If
        If Script_Tr = "yes" Then
            Script_Tr = True
        Else
            Script_Tr = False
        End If
        If Script_Td = "yes" Then
            Script_Td = True
        Else
            Script_Td = False
        End If
    End If

    If Trim(Request("ModifyCollectionImageFrom")) = "yes" Then
        If SaveFiles = "yes" Then
            SaveFiles = True
        Else
            SaveFiles = False
        End If
        If AddWatermark = "yes" Then
            AddWatermark = True
        Else
            AddWatermark = False
        End If

        If AddThumb = "yes" Then
            AddThumb = True
        Else
            AddThumb = False
        End If

        If SaveFlashUrlToFile = "yes" Then
            SaveFlashUrlToFile = True
        Else
            SaveFlashUrlToFile = False
        End If
    End If

    If CreateImmediate = "yes" Then
        CreateImmediate = True
    Else
        CreateImmediate = False
    End If
    
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from PE_Item where ItemID in (" & BatchItemID & ")"
    rs.Open sql, Conn, 1, 3
    Do While Not rs.EOF
        If Trim(Request("ModifyChannelIDFrom")) = "yes" Then rs("ChannelID") = ChannelID
        If Trim(Request("ModifyClassIDFrom")) = "yes" Then rs("ClassID") = ClassID
        If Trim(Request("ModifySpecialIDFrom")) = "yes" Then rs("SpecialID") = SpecialID
        If Trim(Request("ModifyTemplateIDFrom")) = "yes" Then rs("TemplateID") = SpecialID
        If Trim(Request("ModifyPageFrom")) = "yes" Then
            rs("PaginationType") = PaginationType
            rs("MaxCharPerPage") = MaxCharPerPage
        End If
        If Trim(Request("ModifyArticleAttributeFrom")) = "yes" Then
            rs("IncludePicYn") = IncludePicYn
            rs("DefaultPicYn") = DefaultPicYn
            rs("OnTop") = OnTop
            rs("Elite") = Elite
            rs("Stars") = Stars
        End If
        If Trim(Request("ModifyHistFrom")) = "yes" Then rs("Hits") = Hits
        If Trim(Request("ModifyTimeFrom")) = "yes" Then
            rs("UpDateType") = UpDateType
            If UpDateType = 2 Then
                rs("UpDateTime") = UpdateTime
            End If
        End If
        If Trim(Request("ModifyCommentFrom")) = "yes" Then rs("ShowCommentLink") = ShowCommentLink
        If Trim(Request("ModifySkinIDFrom")) = "yes" Then rs("SkinID") = SkinID
        If Trim(Request("ModifyScriptFrom")) = "yes" Then
            rs("Script_Iframe") = Script_Iframe
            rs("Script_Object") = Script_Object
            rs("Script_Script") = Script_Script
            rs("Script_Class") = Script_Class
            rs("Script_Div") = Script_Div
            rs("Script_Span") = Script_Span
            rs("Script_Img") = Script_Img
            rs("Script_Font") = Script_Font
            rs("Script_A") = Script_A
            rs("Script_Html") = Script_Html
            rs("Script_Table") = Script_Table
            rs("Script_Tr") = Script_Tr
            rs("Script_Td") = Script_Td
        End If
        If Trim(Request("ModifyCollectionNumFrom")) = "yes" Then
            rs("CollectionNum") = CollectionNum
            rs("CollectionType") = CollectionType
        End If

        If Trim(Request("ModifyCollectionImageFrom")) = "yes" Then
            rs("SaveFiles") = SaveFiles
            rs("AddWatermark") = AddWatermark
            rs("AddThumb") = AddThumb
            rs("SaveFlashUrlToFile") = SaveFlashUrlToFile
        End If

        If Trim(Request("ModifyCollectionCompositorFrom")) = "yes" Then
            rs("CollecOrder") = CollecOrder
        End If


        If Trim(Request("ModifyInfoPurview")) = "yes" Then
            rs("InfoPurview") = InfoPurview
            rs("arrGroupID") = arrGroupID
        End If
        If Trim(Request("ModifyInfoPoint")) = "yes" Then
            rs("InfoPoint") = InfoPoint
        End If
        If Trim(Request("ModifyChargeType")) = "yes" Then
            rs("ChargeType") = ChargeType
            rs("PitchTime") = DividePercent
            rs("ReadTimes") = ReadTimes
        End If
        If Trim(Request("ModifyDividePercent")) = "yes" Then
            rs("DividePercent") = DividePercent
        End If

        If Trim(Request("ModifyItemEstate")) = "yes" Then
            rs("Status") = Status
            rs("CreateImmediate") = CreateImmediate
        End If

        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Call WriteSuccessMsg("�����޸Ĳɼ����Գɹ���", ComeUrl)
End Sub
'=================================================
'��������Import
'��  �ã�������Ŀ��һ��
'=================================================
Sub Import()
    Response.Write "<br>" & vbCrLf
    Response.Write "<form name='myform' method='post' action='Admin_CollectionManage.asp?action=Import2'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center'><strong>�ɼ���Ŀ���루��һ����</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;������Ҫ����Ĳɼ����ݿ���ļ�����"
    Response.Write "        <input name='ItemMdb' type='text' id='ItemMdb' value='../Temp/PE_Item.mdb' size='20' maxlength='50'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' ��һ�� '>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='Import2'> </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub
'=================================================
'��������Import2
'��  �ã�������Ŀ�ɼ��ڶ���
'=================================================
Sub Import2()
    On Error Resume Next
    Dim rs, sql
    Dim mdbname, tconn, trs, iCount
    mdbname = Replace(Trim(Request.Form("ItemMdb")), "'", "")
    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д�������ݿ���"
    End If

    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݿ����ʧ��,���Ժ�����,����ԭ��" & Err.Description
        Err.Clear
    End If

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    Response.Write "<br>" & vbCrLf
    Response.Write "<form name='myform' method='post' action='Admin_CollectionManage.asp?action=DoImport'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center'><strong>��Ŀ�ɼ����루�ڶ�����</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100' align='center'>"
    Response.Write "        <br>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr align='center'>"
    Response.Write "            <td><strong>��������Ĳɼ���Ŀ</strong><br>"
    Response.Write "<select name='ItemID' size='2' multiple style='height:300px;width:250px;'>"
    sql = "select * from PE_Item"
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, tconn, 1, 1
    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>û���κβɼ���Ŀ</option>"
        iCount = 0
    Else
        iCount = rs.RecordCount
        Do While Not rs.EOF
            Response.Write "<option value='" & rs("ItemID") & "'>" & rs("ItemName") & "</option>"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Response.Write "</select></td>"
    Response.Write "            <td width='80'><input type='submit' name='Submit' value='����&gt;&gt;' "
    If iCount = 0 Then Response.Write " disabled"
    Response.Write "></td>"
    Response.Write "            <td><strong>ϵͳ���Ѿ����ڵ���Ŀ�ɼ�</strong><br>"
    Response.Write "             <select name='tItemID' size='2' multiple style='height:300px;width:250px;' disabled>"
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>û���κβɼ���Ŀ</option>"
    Else
        Do While Not rs.EOF
            Response.Write "<option value='" & rs("ItemID") & "'>" & rs("ItemName") & "</option>"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Response.Write "              </select></td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "     <br><b>��ʾ����ס��Ctrl����Shift�������Զ�ѡ</b><br>"
    Response.Write "        <input name='ItemMdb' type='hidden' id='ItemMdb' value='" & mdbname & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='DoImport'>"
    Response.Write "        <br>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub
'=================================================
'��������DoImport
'��  �ã�����ɼ���Ŀ����
'=================================================
Sub DoImport()
    On Error Resume Next
    Dim mdbname, tconn, trs, Table_PE_Filters
    Dim ItemID
    Dim rs, sql
    ItemID = Trim(Request("ItemID"))
    mdbname = Replace(Trim(Request.Form("Itemmdb")), "'", "")
    If IsValidID(ItemID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ����Ĳɼ���Ŀ</li>"
    End If
    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д�������ݿ���"
    End If
    
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If
    
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
    If Err.Number <> 0 Then
        ErrMsg = ErrMsg & "<li>���ݿ����ʧ��,���Ժ�����,����ԭ��" & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    Set rs = tconn.Execute(" select * from PE_Item where ItemID in (" & ItemID & ")  order by ItemID")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_Item", Conn, 1, 3
    Do While Not rs.EOF
        trs.addnew
        trs("ItemName") = rs("ItemName")
        trs("ChannelID") = 1
        trs("ClassID") = rs("ClassID")
        trs("SpecialID") = rs("SpecialID")
        trs("WebName") = rs("WebName")
        trs("WebUrl") = rs("WebUrl")
        trs("ItemDoem") = rs("ItemDoem")
        trs("ListStr") = rs("ListStr")
        trs("LsString") = rs("LsString")
        trs("LoString") = rs("LoString")
        trs("ListPaingType") = rs("ListPaingType")
        trs("LPsString") = rs("LPsString")
        trs("LPoString") = rs("LPoString")
        trs("ListPaingStr1") = rs("ListPaingStr1")
        trs("ListPaingStr2") = rs("ListPaingStr2")
        trs("ListPaingID1") = rs("ListPaingID1")
        trs("ListPaingID2") = rs("ListPaingID2")
        trs("ListPaingStr3") = rs("ListPaingStr3")
        trs("HsString") = rs("HsString")
        trs("HoString") = rs("HoString")
        trs("HttpUrlType") = rs("HttpUrlType")
        trs("HttpUrlStr") = rs("HttpUrlStr")
        trs("TsString") = rs("TsString")
        trs("ToString") = rs("ToString")
        trs("CsString") = rs("CsString")
        trs("CoString") = rs("CoString")
        trs("UpDateType") = rs("UpDateType")
        trs("DateType") = rs("DateType")
        trs("DsString") = rs("DsString")
        trs("DoString") = rs("DoString")
        trs("AuthorType") = rs("AuthorType")
        trs("AuthorStr") = rs("AuthorStr")
        trs("AsString") = rs("AsString")
        trs("AoString") = rs("AoString")
        trs("CopyFromType") = rs("CopyFromType")
        trs("FsString") = rs("FsString")
        trs("FoString") = rs("FoString")
        trs("CopyFromStr") = rs("CopyFromStr")
        trs("KeyType") = rs("KeyType")
        trs("KsString") = rs("KsString")
        trs("KoString") = rs("KoString")
        trs("KeyStr") = rs("KeyStr")
        trs("NewsPaingType") = rs("NewsPaingType")
        trs("NpsString") = rs("NpsString")
        trs("NpoString") = rs("NpoString")
        trs("NewsPaingStr1") = rs("NewsPaingStr1")
        trs("NewsPaingStr2") = rs("NewsPaingStr2")
        trs("Flag") = False
        trs("PsString") = rs("PsString")
        trs("PoString") = rs("PoString")
        trs("PhsString") = rs("PhsString")
        trs("PhoString") = rs("PhoString")
        trs("ThumbnailType") = rs("ThumbnailType")
        trs("ThsString") = rs("ThsString")
        trs("ThoString") = rs("ThoString")
        trs.Update
        rs.MoveNext
    Loop
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    '�ж�PE_Filters ���Ƿ����
    Table_PE_Filters = True
    tconn.Execute ("select FilterName from PE_Filters")
    If Err Then
        Table_PE_Filters = False
    End If
    If Table_PE_Filters = True Then
        Set rs = tconn.Execute("select * from PE_Filters")
        Set trs = Server.CreateObject("adodb.recordset")
        trs.Open "select * from PE_Filters", Conn, 1, 3
        If rs.BOF Or rs.EOF Then
        Else
            Do While Not rs.EOF
                trs.addnew
                trs("ItemID") = 0
                trs("FilterName") = rs("FilterName")
                trs("FilterObject") = rs("FilterObject")
                trs("FilterType") = rs("FilterType")
                trs("FilterContent") = rs("FilterContent")
                trs("FisString") = rs("FisString")
                trs("FioString") = rs("FioString")
                trs("FilterRep") = rs("FilterRep")
                trs("Flag") = False
                trs.Update
                rs.MoveNext
            Loop
        End If
        trs.Close
        Set trs = Nothing
        rs.Close
        Set rs = Nothing
    End If
    tconn.Close
    Set tconn = Nothing
    Call WriteSuccessMsg("�Ѿ��ɹ���ָ�������ݿ��е���ѡ�еĲɼ���Ŀ��<br><br>�㻹��Ҫ�����òɼ���Ŀ�����Բ�������ɲɼ�������", "Admin_CollectionManage.asp?action=ItemManage")
End Sub
'=================================================
'��������Export
'��  �ã������ɼ���Ŀ
'=================================================
Sub Export()
    Dim rs, sql, iCount
    sql = "select * from PE_Item"
    Set rs = Conn.Execute(sql)

    Response.Write "<br>" & vbCrLf
    Response.Write "<form name='myform' method='post' action='Admin_CollectionManage.asp?action=DoExport'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>�ɼ���Ŀ����</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'><td height='10'></td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='center'>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr>"
    Response.Write "           <td>"
    Response.Write "            <select name='ItemID' size='2' multiple style='height:300px;width:450px;'>"
    If rs.BOF And rs.EOF Then
        Response.Write "         <option value=''>��û�вɼ���Ŀ��</option>"
        '�ر��ύ��ť
        iCount = 0
    Else
        iCount = rs.RecordCount
        Do While Not rs.EOF
            Response.Write "     <option value='" & rs("ItemID") & "'>" & rs("ItemName") & "</option>"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Response.Write "         </select>"
    Response.Write "       </td>"
    Response.Write "       <td align='left'>&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='Submit' value=' ѡ������ ' onclick='SelectAll()'>"
    Response.Write "       <br><br>&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='Submit' value=' ȡ��ѡ�� ' onclick='UnSelectAll()'><br><br><br><b>&nbsp;��ʾ����ס��Ctrl����Shift�������Զ�ѡ</b></td>"
    Response.Write "      </tr>"
    Response.Write "      <tr height='30'>"
    Response.Write "        <td colspan='2'>Ŀ�����ݿ⣺<input name='Itemmdb' type='text' id='ItemMdb' value='../Temp/PE_Item.mdb' size='20' maxlength='50'>&nbsp;&nbsp;<INPUT TYPE='checkbox' NAME='FormatConn' value='yes' id='id' checked> �����Ŀ�����ݿ�</td>"
    Response.Write "      </tr>"
    Response.Write "      <tr height='50'>"
    Response.Write "         <td colspan='2' align='center'><input type='submit' name='Submit' value='ִ�е�������' onClick=""document.myform.Action.value='DoExport';"">"
    Response.Write "          <input name='Action' type='hidden' id='Action' value='Export'>"
    Response.Write "         </td>"
    Response.Write "        </tr>"
    Response.Write "    </table>"
    Response.Write "   </td>"
    Response.Write " </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.ItemID.length;i++){" & vbCrLf
    Response.Write "    document.myform.ItemID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.ItemID.length;i++){" & vbCrLf
    Response.Write "    document.myform.ItemID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub
'=================================================
'��������DoExport
'��  �ã�������Ŀ�ɼ�����
'=================================================
Sub DoExport()
    On Error Resume Next
    
    Dim rs, sql
    Dim mdbname, tconn, trs
    Dim ItemID, FormatConn, Table_PE_Item
    FormatConn = Request.Form("FormatConn")
    ItemID = Trim(Request("ItemID"))
    mdbname = Replace(Trim(Request.Form("Itemmdb")), "'", "")
    If IsValidID(ItemID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����Ĳɼ���Ŀ</li>"
    End If
    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д�������ݿ���"
    End If
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݿ����ʧ��,���Ժ�����,����ԭ��" & Err.Description
        Err.Clear
    End If
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If
    If FormatConn <> "" Then
        tconn.Execute ("delete from PE_Item")
        tconn.Execute ("delete from PE_Filters")
    End If
    
    Table_PE_Item = True
    tconn.Execute ("select PsString from PE_Item")
    If Err Then
        Table_PE_Item = False
    End If
    '�ж�PE_Item ���Ƿ����
    If Table_PE_Item = False Then
        tconn.Execute ("alter table [PE_Item]  add COLUMN UpDateType int null")
        tconn.Execute ("alter table [PE_Item]  add COLUMN DateType   int null")
        tconn.Execute ("alter table [PE_Item]  add COLUMN DsString   nvarchar(255) null")
        tconn.Execute ("alter table [PE_Item]  add COLUMN DoString   nvarchar(255) null")
        tconn.Execute ("alter table [PE_Item]  add COLUMN ShowCommentLink bit")
        tconn.Execute ("alter table [PE_Item]  add COLUMN Script_Table bit")
        tconn.Execute ("alter table [PE_Item]  add COLUMN Script_Tr  bit")
        tconn.Execute ("alter table [PE_Item]  add COLUMN Script_Td  bit")
        tconn.Execute ("alter table [PE_Item]  add COLUMN PsString  text null")
        tconn.Execute ("alter table [PE_Item]  add COLUMN PoString  text null")
        tconn.Execute ("alter table [PE_Item]  add COLUMN PhsString  text null")
        tconn.Execute ("alter table [PE_Item]  add COLUMN PhoString  text null")
        tconn.Execute ("alter table [PE_Item]  add COLUMN ThumbnailType int null")
        tconn.Execute ("alter table [PE_Item]  add COLUMN ThsString   text null")
        tconn.Execute ("alter table [PE_Item]  add COLUMN ThoString   text null")
        tconn.Execute ("alter table [PE_Item]  drop CollecTest")
        tconn.Execute ("alter table [PE_Item]  drop Content_view")
        tconn.Execute ("alter table [PE_Item]  drop UploadDir")
        tconn.Execute ("alter table [PE_Item]  drop UpFileType")
    End If
    
    Set rs = Conn.Execute("select * from PE_Item where ItemID in (" & ItemID & ")  order by SkinID ")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_Item", tconn, 1, 3
    Do While Not rs.EOF
        trs.addnew

        trs("ItemName") = rs("ItemName")
        trs("ChannelID") = rs("ChannelID")
        trs("ClassID") = rs("ClassID")
        trs("SpecialID") = rs("SpecialID")
        trs("WebName") = rs("WebName")
        trs("WebUrl") = rs("WebUrl")
        trs("ItemDoem") = rs("ItemDoem")
        trs("ListStr") = rs("ListStr")
        trs("LsString") = rs("LsString")
        trs("LoString") = rs("LoString")
        trs("ListPaingType") = rs("ListPaingType")
        trs("LPsString") = rs("LPsString")
        trs("LPoString") = rs("LPoString")
        trs("ListPaingStr1") = rs("ListPaingStr1")
        trs("ListPaingStr2") = rs("ListPaingStr2")
        trs("ListPaingID1") = rs("ListPaingID1")
        trs("ListPaingID2") = rs("ListPaingID2")
        trs("ListPaingStr3") = rs("ListPaingStr3")
        trs("HsString") = rs("HsString")
        trs("HoString") = rs("HoString")
        trs("HttpUrlType") = rs("HttpUrlType")
        trs("HttpUrlStr") = rs("HttpUrlStr")
        trs("TsString") = rs("TsString")
        trs("ToString") = rs("ToString")
        trs("CsString") = rs("CsString")
        trs("CoString") = rs("CoString")
        trs("UpDateType") = rs("UpDateType")
        trs("DateType") = rs("DateType")
        trs("DsString") = rs("DsString")
        trs("DoString") = rs("DoString")
        trs("AuthorType") = rs("AuthorType")
        trs("AuthorStr") = rs("AuthorStr")
        trs("AsString") = rs("AsString")
        trs("AoString") = rs("AoString")
        trs("CopyFromType") = rs("CopyFromType")
        trs("FsString") = rs("FsString")
        trs("FoString") = rs("FoString")
        trs("CopyFromStr") = rs("CopyFromStr")
        trs("KeyType") = rs("KeyType")
        trs("KsString") = rs("KsString")
        trs("KoString") = rs("KoString")
        trs("KeyStr") = rs("KeyStr")
        trs("NewsPaingType") = rs("NewsPaingType")
        trs("NpsString") = rs("NpsString")
        trs("NpoString") = rs("NpoString")
        trs("NewsPaingStr1") = rs("NewsPaingStr1")
        trs("NewsPaingStr2") = rs("NewsPaingStr2")
        trs("Flag") = False
        trs("PsString") = rs("PsString")
        trs("PoString") = rs("PoString")
        trs("PhsString") = rs("PhsString")
        trs("PhoString") = rs("PhoString")
        trs("ThumbnailType") = rs("ThumbnailType")
        trs("ThsString") = rs("ThsString")
        trs("ThoString") = rs("ThoString")
        trs.Update
        rs.MoveNext
    Loop
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    
    Set rs = Conn.Execute("select * from PE_Filters where ItemID in (" & ItemID & ") or ItemId=-1 order by ItemID")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_Filters", tconn, 1, 3
    If rs.BOF Or rs.EOF Then
    Else
        Do While Not rs.EOF
            trs.addnew
            trs("ItemID") = rs("ItemID")
            trs("FilterName") = rs("FilterName")
            trs("FilterObject") = rs("FilterObject")
            trs("FilterType") = rs("FilterType")
            trs("FilterContent") = rs("FilterContent")
            trs("FisString") = rs("FisString")
            trs("FioString") = rs("FioString")
            trs("FilterRep") = rs("FilterRep")
            trs("Flag") = False
            trs.Update
            rs.MoveNext
        Loop
    End If
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    tconn.Close
    Set tconn = Nothing
    Call WriteSuccessMsg("�Ѿ��ɹ�����ѡ�еĲɼ���Ŀ������ָ�������ݿ��У�", ComeUrl)
End Sub

'*************************  ��ģ�����������  *******************************
'*************************  ��ģ����չ��ʼ  *******************************
'==================================================
'��������DataBaseModify
'��  �ã����ݿ�����
'==================================================
Sub DataBaseModify()
    On Error Resume Next
    Dim Table_PE_Item, strsql
        
    Table_PE_Item = True '�����ɰ����ݿ��ֶ�����
    Conn.Execute ("select UpDateType from PE_Item")
    If Err Then
        Table_PE_Item = False
    End If
    If Table_PE_Item = False Then
        If SystemDatabaseType = "SQL" Then
            Conn.Execute ("alter table [PE_Item]  add  UpDateType int null")
            Conn.Execute ("alter table [PE_Item]  add  DateType   int null")
            Conn.Execute ("alter table [PE_Item]  add  DsString   ntext null")
            Conn.Execute ("alter table [PE_Item]  add  DoString   ntext null")
            Conn.Execute ("alter table [PE_Item]  add  ShowCommentLink bit")
            Conn.Execute ("alter table [PE_Item]  add  Script_Table bit")
            Conn.Execute ("alter table [PE_Item]  add  Script_Tr  bit")
            Conn.Execute ("alter table [PE_Item]  add  Script_Td  bit")
            Conn.Execute ("ALTER TABLE [PE_Item]  DROP [DF_PE_Item_CollecTest]")
            Conn.Execute ("ALTER TABLE [PE_Item]  DROP [DF_PE_Item_Content_view]")
            Conn.Execute ("alter table [PE_Item]  drop  COLUMN CollecTest")
            Conn.Execute ("alter table [PE_Item]  drop  COLUMN Content_view")
            Conn.Execute ("alter table [PE_Item]  drop  COLUMN UploadDir")
            Conn.Execute ("alter table [PE_Item]  drop  COLUMN UpFileType")
        Else
            Conn.Execute ("alter table [PE_Item]  add COLUMN UpDateType int null")
            Conn.Execute ("alter table [PE_Item]  add COLUMN DateType   int null")
            Conn.Execute ("alter table [PE_Item]  add COLUMN DsString   text null")
            Conn.Execute ("alter table [PE_Item]  add COLUMN DoString   text null")
            Conn.Execute ("alter table [PE_Item]  add COLUMN ShowCommentLink bit")
            Conn.Execute ("alter table [PE_Item]  add COLUMN Script_Table bit")
            Conn.Execute ("alter table [PE_Item]  add COLUMN Script_Tr  bit")
            Conn.Execute ("alter table [PE_Item]  add COLUMN Script_Td  bit")
            Conn.Execute ("alter table [PE_Item]  drop  CollecTest")
            Conn.Execute ("alter table [PE_Item]  drop  Content_view")
            Conn.Execute ("alter table [PE_Item]  drop  UploadDir")
            Conn.Execute ("alter table [PE_Item]  drop  UpFileType")
        End If
        Conn.Execute "update PE_Item set Flag=" & PE_False
    End If

    Table_PE_Item = True '����2006 ���ֶ�
    Conn.Execute ("select ThumbnailType from PE_Item")
    If Err Then
        Table_PE_Item = False
    End If
    If Table_PE_Item = False Then
        If SystemDatabaseType = "SQL" Then
            Conn.Execute ("alter table [PE_Item]  add  PsString   ntext null")
            Conn.Execute ("alter table [PE_Item]  add  PoString   ntext null")
            Conn.Execute ("alter table [PE_Item]  add  PhsString   ntext null")
            Conn.Execute ("alter table [PE_Item]  add  PhoString   ntext null")
            Conn.Execute ("alter table [PE_Item]  add  ThumbnailType int null")
            Conn.Execute ("alter table [PE_Item]  add  ThsString   ntext null")
            Conn.Execute ("alter table [PE_Item]  add  ThoString   ntext null")

        Else
            Conn.Execute ("alter table [PE_Item]  add  COLUMN PsString  text null")
            Conn.Execute ("alter table [PE_Item]  add  COLUMN PoString  text null")
            Conn.Execute ("alter table [PE_Item]  add  COLUMN PhsString  text null")
            Conn.Execute ("alter table [PE_Item]  add  COLUMN PhoString  text null")
            Conn.Execute ("alter table [PE_Item]  add  COLUMN ThumbnailType int null")
            Conn.Execute ("alter table [PE_Item]  add  COLUMN ThsString   text null")
            Conn.Execute ("alter table [PE_Item]  add  COLUMN ThoString   text null")
        End If
    End If
End Sub
'=================================================
'��������ShowChekcFormVbs
'��  �ã����Դ����Ƿ�Ψһ
'=================================================
Sub ShowChekcFormVbs()
    Response.Write "<script language=""VBScript"">" & vbCrLf
    Response.Write " Sub ceshi(Num)" & vbCrLf
    Response.Write "    Dim content" & vbCrLf
    Response.Write "    Content=document.form1.Content.value" & vbCrLf
    Response.Write "    Select Case Num   " & vbCrLf
    Response.Write "    Case 1" & vbCrLf
    Response.Write "        huoqv=document.form1.LsString.value" & vbCrLf
    Response.Write "    Case 2" & vbCrLf
    Response.Write "        huoqv=document.form1.LoString.value" & vbCrLf
    Response.Write "    Case 3" & vbCrLf
    Response.Write "        huoqv=document.form1.LPsString.value" & vbCrLf
    Response.Write "    Case 4" & vbCrLf
    Response.Write "        huoqv=document.form1.LPoString.value" & vbCrLf
    Response.Write "    Case 5" & vbCrLf
    Response.Write "        huoqv=document.form1.TsString.value" & vbCrLf
    Response.Write "    Case 6" & vbCrLf
    Response.Write "        huoqv=document.form1.ToString.value" & vbCrLf
    Response.Write "    Case 7" & vbCrLf
    Response.Write "        huoqv=document.form1.CsString.value" & vbCrLf
    Response.Write "    Case 8" & vbCrLf
    Response.Write "        huoqv=document.form1.CoString.value" & vbCrLf
    Response.Write "    Case 9" & vbCrLf
    Response.Write "        huoqv=document.form1.AsString.value" & vbCrLf
    Response.Write "    Case 10" & vbCrLf
    Response.Write "        huoqv=document.form1.AoString.value" & vbCrLf
    Response.Write "    Case 11" & vbCrLf
    Response.Write "        huoqv=document.form1.FsString.value" & vbCrLf
    Response.Write "    Case 12" & vbCrLf
    Response.Write "        huoqv=document.form1.FoString.value" & vbCrLf
    Response.Write "    Case 13" & vbCrLf
    Response.Write "        huoqv=document.form1.KsString.value" & vbCrLf
    Response.Write "    Case 14" & vbCrLf
    Response.Write "        huoqv=document.form1.KoString.value" & vbCrLf
    Response.Write "    Case 15" & vbCrLf
    Response.Write "        huoqv=document.form1.NPsString.value" & vbCrLf
    Response.Write "    Case 16" & vbCrLf
    Response.Write "        huoqv=document.form1.NPoString.value" & vbCrLf
    Response.Write "    Case 17" & vbCrLf
    Response.Write "        huoqv=document.form1.DsString.value" & vbCrLf
    Response.Write "    Case 18" & vbCrLf
    Response.Write "        huoqv=document.form1.DoString.value" & vbCrLf
    Response.Write "    Case 19" & vbCrLf
    Response.Write "        huoqv=document.form1.PsString.value" & vbCrLf
    Response.Write "    Case 20" & vbCrLf
    Response.Write "        huoqv=document.form1.PoString.value" & vbCrLf
    Response.Write "    Case 21" & vbCrLf
    Response.Write "        huoqv=document.form1.IsString.value" & vbCrLf
    Response.Write "    Case 22" & vbCrLf
    Response.Write "        huoqv=document.form1.IoString.value" & vbCrLf
    Response.Write "    Case Else" & vbCrLf
    Response.Write "        Exit sub" & vbCrLf
    Response.Write "    End Select" & vbCrLf
    Response.Write "    if huoqv="""" then " & vbCrLf
    Response.Write "       alert(""������Ч������Ϊ�գ�"")" & vbCrLf
    Response.Write "       exit Sub" & vbCrLf
    Response.Write "    End if " & vbCrLf
    Response.Write "    If InStr(Content,huoqv) = 0 Then" & vbCrLf
    Response.Write "       alert(""������Ч����ҳû����Щ���롣"")" & vbCrLf
    Response.Write "    Else" & vbCrLf
    Response.Write "       If InStr(Mid(Content,InStr(Content,huoqv)+LenB(huoqv),LenB(Content)),huoqv) = 0 Then" & vbCrLf
    Response.Write "          alert(""���Գɹ���������ҳ����Ψһ�ġ�"")" & vbCrLf
    Response.Write "       Else" & vbCrLf
    Response.Write "          alert(""����ʧ�ܣ��������ظ�,��ʼ�����������һ��������Ψһ����Ч��"")" & vbCrLf
    Response.Write "       End if" & vbCrLf
    Response.Write "    End if" & vbCrLf
    Response.Write " End Sub" & vbCrLf
    Response.Write " </script>" & vbCrLf
    Response.Write "<script language=""JavaScript"">" & vbCrLf
    Response.Write " <!--" & vbCrLf
    Response.Write " var openurl=0;" & vbCrLf
    Response.Write " function CheckForm(){" & vbCrLf
    Response.Write "    if (document.form1.Content.value.length > 200000){" & vbCrLf '�ύ���ܴ���200K
    Response.Write "        document.form1.Content.value="""";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write " function setFileFields(weburl){   " & vbCrLf
    Response.Write "    if (openurl==0){" & vbCrLf
    Response.Write "        str=""<iframe id='IFrame1' marginwidth=0 marginheight=0 frameborder=0  width='785' height='400' src=""+weburl+""></iframe>"";" & vbCrLf
    Response.Write "        objFiles.innerHTML=str;" & vbCrLf
    Response.Write "        openurl=1" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write " }" & vbCrLf
    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

'**************************************************
'��������GetManagePath
'��  �ã��ɼ���Ŀ����
'��  ����iChannelID ----Ƶ��ID
'����ֵ���ɼ�������
'**************************************************
Function GetManagePath(ByVal iChannelID)
    Dim strPath, sqlPath, rsPath
    iChannelID = PE_CLng(iChannelID)
    strPath = "<IMG SRC='images/img_u.gif' height='12'>�����ڵ�λ�ã��ɼ�" & ItemName & "����&nbsp;&gt;&gt;&nbsp;"

    If iChannelID = 0 Then
        strPath = strPath & "<a href='" & strFileName & "&ChannelID=0'>����Ƶ����Ŀ</a>"
    Else
        sqlPath = "select ChannelID,ChannelName from PE_Channel where ChannelID=" & iChannelID
        Set rsPath = Conn.Execute(sqlPath)

        If rsPath.BOF And rsPath.EOF Then
            strPath = strPath & "�����Ƶ������"
        Else
            strPath = strPath & "<a href='" & strFileName & "&ChannelID=" & rsPath(0) & "'>" & rsPath(1) & "��Ŀ</a>"
        End If

        rsPath.Close
        Set rsPath = Nothing
    End If

    GetManagePath = strPath
End Function
'=================================================
'��������GetChannel_Option
'��  �ã�����Ƶ��
'��  ����iChannelID  ----Ƶ������
'=================================================
Sub GetChannel_Option(iChannelID)
    Dim strChannel, sqlChannel, rsChannel
    sqlChannel = "select ChannelID,ChannelName,Disabled from PE_Channel  where ModuleType=1 and ChannelType<=1 order by OrderID"
    Set rsChannel = Conn.Execute(sqlChannel)
    If rsChannel.BOF And rsChannel.BOF Then
        strChannel = "<option value=''>�������Ƶ��</option>"
    Else
        Do While Not rsChannel.EOF
            If rsChannel(2) = True Then
                strChannel = strChannel & "<option value=''>" & rsChannel(1) & "(��Ƶ���ѱ�����)</option>"
            Else
                If rsChannel(0) = iChannelID Then
                    strChannel = strChannel & "<option value='" & rsChannel(0) & "' selected>" & rsChannel(1) & "</option>"
                Else
                    strChannel = strChannel & "<option value='" & rsChannel(0) & "'>" & rsChannel(1) & "</option>"
                End If
            End If
            rsChannel.MoveNext
        Loop
    End If
    rsChannel.Close
    Set rsChannel = Nothing
    Response.Write strChannel
End Sub

'=================================================
'��������GetSkin_Option
'��  �ã�����������Ŀ���
'��  ����iSkinID  ----��ĿID
'=================================================
Function GetSkin_Option(iSkinID)
    Dim sqlSkin, rsSkin, strSkin
    If IsNull(iSkinID) Then iSkinID = 0
    strSkin = ""
    sqlSkin = "select * from PE_Skin"
    Set rsSkin = Conn.Execute(sqlSkin)
    If rsSkin.BOF And rsSkin.EOF Then
        strSkin = strSkin & "<option value=''>������ӷ��</option>"
    Else
        If iSkinID = 0 Then
            strSkin = strSkin & "<option value='0' selected>ϵͳĬ�Ϸ��</option>"
        Else
            strSkin = strSkin & "<option value='0'>ϵͳĬ�Ϸ��</option>"
        End If
        Do While Not rsSkin.EOF
            strSkin = strSkin & "<option value='" & rsSkin("SkinID") & "'"
            If rsSkin("SkinID") = iSkinID Then
                strSkin = strSkin & " selected"
            End If
            strSkin = strSkin & ">" & rsSkin("SkinName")
            If rsSkin("IsDefault") = True Then
                strSkin = strSkin & "��Ĭ�ϣ�"
            End If
            strSkin = strSkin & "</option>"
            rsSkin.MoveNext
        Loop
    End If
    rsSkin.Close
    Set rsSkin = Nothing
    GetSkin_Option = strSkin
End Function



'==================================================
'��������UrlEncoding
'��  �ã�ת������
'==================================================
Function UrlEncoding(DataStr)
    On Error Resume Next
    Dim StrReturn, Si, ThisChr, InnerCode, Hight8, Low8
    StrReturn = ""
    For Si = 1 To Len(DataStr)
        ThisChr = Mid(DataStr, Si, 1)
        If Abs(Asc(ThisChr)) < &HFF Then
            StrReturn = StrReturn & ThisChr
        Else
            InnerCode = Asc(ThisChr)
            If InnerCode < 0 Then
                InnerCode = InnerCode + &H10000
            End If
            Hight8 = (InnerCode And &HFF00) \ &HFF
            Low8 = InnerCode And &HFF
            StrReturn = StrReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
        End If
    Next
    UrlEncoding = StrReturn
End Function

'==================================================
'��������ReplaceBadChar2
'��  �ã��滻������ʽ�����ַ�
'��  ����strChar-----Ҫ���˵��ַ�
'����ֵ���滻����ַ�
'==================================================
Function ReplaceBadChar2(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceBadChar2 = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "^,(,),*,?,[,],$,+,|,{,}"
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "\" & arrBadChar(i))
    Next
    ReplaceBadChar2 = tempChar
End Function

'**************************************************
'��������CreateKeyWord
'��  �ã��ɸ������ַ������ɹؼ���
'��  ����Constr---Ҫ���ɹؼ��ֵ�ԭ�ַ���
'����ֵ�����ɵĹؼ���
'**************************************************
Function CreateKeyWord(ByVal ConStr, ByVal Num)
    If ConStr = "" Or IsNull(ConStr) = True Or ConStr = "$False$" Or IsNumeric(Num) = False Then
        CreateKeyWord = "$False$"
        Exit Function
    End If
    If CLng(Num) < 2 Then
        Num = 2
    End If
    ConStr = Replace(ConStr, Chr(32), "")
    ConStr = Replace(ConStr, Chr(9), "")
    ConStr = Replace(ConStr, "&nbsp;", "")
    ConStr = Replace(ConStr, " ", "")
    ConStr = Replace(ConStr, "(", "")
    ConStr = Replace(ConStr, ")", "")
    ConStr = Replace(ConStr, "<", "")
    ConStr = Replace(ConStr, ">", "")
    Dim i, ConstrTemp
    If Num >= Len(ConStr) Then
        CreateKeyWord = Left(ConStr, 254)
        Exit Function
    Else
        For i = 1 To Len(ConStr)
            If i + Num > Len(ConStr) Then
                Exit For
            Else
                ConstrTemp = ConstrTemp & "|" & Mid(ConStr, i, Num)
            End If
        Next
    End If
    If Len(ConstrTemp) < 254 Then
        ConstrTemp = ConstrTemp & "|"
    Else
        ConstrTemp = Left(ConstrTemp, 254) & "|"
    End If
    CreateKeyWord = ConstrTemp
End Function

Function IsRadioChecked(ByVal Compare1, ByVal Compare2)
    If Compare1 = Compare2 Then
        IsRadioChecked = " checked"
    Else
        IsRadioChecked = ""
    End If
End Function

Function IsOptionSelected(ByVal Compare1, ByVal Compare2)
    If Compare1 = Compare2 Then
        IsOptionSelected = " selected"
    Else
        IsOptionSelected = ""
    End If
End Function

Function IsStyleDisplay(ByVal Compare1, ByVal Compare2)

    If Compare1 = Compare2 Then
        IsStyleDisplay = " style='display:'"
    Else
        IsStyleDisplay = " style='display:none'"
    End If

End Function
%>
