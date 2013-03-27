<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_ContentEx.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "AD"   '����Ȩ��


Dim ADID, ZoneID
Dim ZoneConfig, ZoneTypeNum, IAB_Size

ZoneID = Trim(Request("ZoneID"))

If Action = "" Then
    Action = "ZoneList"
End If
If ChannelID = 0 Then
    ChannelID = -2
Else
    ChannelID = PE_CLng(ChannelID)
End If
If IsValidID(ZoneID) = False Then
    ZoneID = ""
End If
strFileName = "Admin_Advertisement.asp?Action=" & Action

Response.Write "<html><head><title>������</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("�� վ �� �� �� ��", 10021)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>��������</strong></td>" & vbCrLf
Response.Write "    <td>" & vbCrLf
Response.Write "      <a href='Admin_Advertisement.asp?Action=ZoneList'>����λ����</a>&nbsp;|&nbsp;"
Response.Write "      <a href='Admin_Advertisement.asp?Action=AddZone'>����°�λ</a>&nbsp;|&nbsp;"
Response.Write "      <a href='Admin_Advertisement.asp?Action=ADList'>��վ������</a>&nbsp;|&nbsp;"
Response.Write "      <a href='Admin_Advertisement.asp?Action=AddAD'>����¹��</a>&nbsp;|&nbsp;"
Response.Write "      <a href='Admin_Advertisement.asp?Action=ZoneJSTemplate'>���JSģ��</a>&nbsp;|&nbsp;"
Response.Write "      <a href='Admin_UploadFile.asp?UploadDir=UploadAdPic'>����ϴ�ͼƬ����</a>&nbsp;|&nbsp;"
Response.Write "      <a href='Admin_Advertisement.asp?Action=Import'>�����λ</a>&nbsp;|&nbsp;"
Response.Write "      <a href='Admin_Advertisement.asp?Action=Export'>������λ</a>&nbsp;|&nbsp;"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
If Not fso.FolderExists(Server.MapPath(InstallDir & ADDir)) Then
    Response.Write "<br><li>�Ҳ�����վ���Ŀ¼��������վ�����е�������ʵ�ʵĹ��Ŀ¼�Ƿ�һ�¡�</li>"
    Response.End
End If

Call InitZoneConfig

Select Case Action
Case "AddAD"
    Call AddAD
Case "ModifyAD"
    Call ModifyAD
Case "CopyAD"
    Call CopyAD
Case "Export"
    Call Export
Case "DoExport"
    Call DoExport
Case "Import"
    Call Import
Case "Import2"
    Call Import2
Case "DoImport"
    Call DoImport
Case "SaveAddAD", "SaveModifyAD"
    Call SaveAD
Case "SetADPassed", "CancelADPassed", "MoveAD", "DelAD"
    Call SetADProperty
Case "ADList"
    Call ADList
Case "PreviewAD"
    Call PreviewAD
Case "AddZone"
    Call AddZone
Case "ModifyZone"
    Call ModifyZone
Case "CopyZone"
    Call CopyZone
Case "SaveAddZone", "SaveModifyZone"
    Call SaveZone
Case "SetZoneActive", "CancelZoneActive", "MoveZone", "DelZone"
    Call SetZoneProperty
Case "ClearZone"
    Call ClearZone
Case "PreviewZone"
    Call PreviewZone
Case "ZoneJSCode"
    Call ZoneJSCode
Case "ZoneList"
    Call ZoneList
Case "ZoneJSTemplate"
    Call ZoneJSTemplate
Case "ModifyTemplate"
    Call ModifyTemplate
Case "SaveTemplate"
    Call SaveTemplate
Case "CreateJSZone"
    Call CreateJSZone
Case Else
    Call ZoneList
End Select
If FoundErr = True Then
    Call WriteEntry(2, AdminName, "���������ʧ�ܣ�ʧ��ԭ��" & ErrMsg)
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If

Response.Write "</body></html>"
Call CloseConn

Sub InitZoneConfig()
    ZoneTypeNum = 9
    ReDim ZoneConfig(9, 4)
    ZoneConfig(0, 0) = ""
    ZoneConfig(0, 1) = ""
    ZoneConfig(1, 0) = "Banner"
    ZoneConfig(1, 1) = "���κ��"
    ZoneConfig(1, 2) = "1"
    ZoneConfig(2, 0) = "Pop"
    ZoneConfig(2, 1) = "��������"
    ZoneConfig(2, 2) = "2,1,100,100,0"
    ZoneConfig(3, 0) = "Move"
    ZoneConfig(3, 1) = "�����ƶ�"
    ZoneConfig(3, 2) = "3,15,200,0.015"
    ZoneConfig(4, 0) = "Fixed"
    ZoneConfig(4, 1) = "�̶�λ��"
    ZoneConfig(4, 2) = "4,100,100"
    ZoneConfig(5, 0) = "Float"
    ZoneConfig(5, 1) = "Ư���ƶ�"
    ZoneConfig(5, 2) = "5,1,100,100"
    ZoneConfig(6, 0) = "Code"
    ZoneConfig(6, 1) = "���ִ���"
    ZoneConfig(6, 2) = "6"
    ZoneConfig(7, 0) = "Couplet"
    ZoneConfig(7, 1) = "�������"
    ZoneConfig(7, 2) = "7"
    ZoneConfig(8, 0) = "BottomLeft"
    ZoneConfig(8, 1) = "������"
    ZoneConfig(8, 2) = "8"
    ZoneConfig(9, 0) = "BottomRight"
    ZoneConfig(9, 1) = "���ҹ��"
    ZoneConfig(9, 2) = "9"

    ReDim IAB_Size(15, 1)
    IAB_Size(0, 0) = "468x60"
    IAB_Size(0, 1) = "IAB - 468 x 60 IMU (������)"
    IAB_Size(1, 0) = "234x60"
    IAB_Size(1, 1) = "IAB - 234 x 60 IMU (������)"
    IAB_Size(2, 0) = "88x31"
    IAB_Size(2, 1) = "IAB -��88 x 31 IMU (С��ť)"
    IAB_Size(3, 0) = "120x90"
    IAB_Size(3, 1) = "IAB - 120 x 90 IMU (��ťһ)"
    IAB_Size(4, 0) = "120x60"
    IAB_Size(4, 1) = "IAB - 120 x 60 IMU (��ť��)"
    IAB_Size(5, 0) = "728x90"
    IAB_Size(5, 1) = "IAB - 728 x 90 IMU (ͨ�����) *"
    IAB_Size(6, 0) = "120x240"
    IAB_Size(6, 1) = "IAB - 120 x 240 IMU (�������)"
    IAB_Size(7, 0) = "125x125"
    IAB_Size(7, 1) = "IAB - 125 x 125 IMU (���ΰ�ť)"
    IAB_Size(8, 0) = "180x150"
    IAB_Size(8, 1) = "IAB - 180 x 150 IMU (������) *"
    IAB_Size(9, 0) = "300x250"
    IAB_Size(9, 1) = "IAB - 300 x 250 IMU (�г�����) *"
    IAB_Size(10, 0) = "336x280"
    IAB_Size(10, 1) = "IAB - 336 x 280 IMU (�󳤷���)"
    IAB_Size(11, 0) = "240x400"
    IAB_Size(11, 1) = "IAB - 240 x 400 IMU (��������)"
    IAB_Size(12, 0) = "250x250"
    IAB_Size(12, 1) = "IAB - 250 x 250 IMU (�����ε���)"
    IAB_Size(13, 0) = "120x600"
    IAB_Size(13, 1) = "IAB - 120 x 600 IMU (Ħ���¥)"
    IAB_Size(14, 0) = "160x600"
    IAB_Size(14, 1) = "IAB - 160 x 600 IMU (��Ħ���¥) *"
    IAB_Size(15, 0) = "300x600"
    IAB_Size(15, 1) = "IAB - 300 x 600 IMU (��ҳ���) *"
End Sub

Sub ShowJS_ADMain(ItemName)
    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "function CheckItem(CB){" & vbCrLf
    Response.Write "  if(document.myform.chkAll.checked){" & vbCrLf
    Response.Write "    document.myform.chkAll.checked = document.myform.chkAll.checked&0;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (CB.checked)" & vbCrLf
    Response.Write "    hL(CB);" & vbCrLf
    Response.Write "  else" & vbCrLf
    Response.Write "    dL(CB);" & vbCrLf
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
    Response.Write "  for (var i=0;i<form.elements.length;i++){" & vbCrLf
    Response.Write "    var e = form.elements[i];" & vbCrLf
    Response.Write "    if (e.name != 'chkAll' && e.disabled == false && e.type == 'checkbox') {" & vbCrLf
    Response.Write "      e.checked = form.chkAll.checked;" & vbCrLf
    Response.Write "      if (e.checked)" & vbCrLf
    Response.Write "        hL(e);" & vbCrLf
    Response.Write "      else" & vbCrLf
    Response.Write "        dL(e);" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function ConfirmDel(){" & vbCrLf
    Response.Write "  if(document.myform.Action.value=='DelZone'||document.myform.Action.value=='DelAD'){" & vbCrLf
    Response.Write "    if(confirm('ȷ��Ҫɾ��ѡ�е�" & ItemName & "��'))" & vbCrLf
    Response.Write "      return true;" & vbCrLf
    Response.Write "    else" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "function hL(E){" & vbCrLf
    Response.Write "  while (E.tagName!='TR') {E=E.parentElement;}" & vbCrLf
    Response.Write "  E.className='tdbg2';" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function dL(E){" & vbCrLf
    Response.Write "  while (E.tagName!='TR') {E=E.parentElement;}" & vbCrLf
    Response.Write "  E.className='tdbg';" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
End Sub

Sub ZoneList()
    Dim rsZone, sqlZone
    Call ShowJS_ADMain("��λ")
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22'>" & GetChannelList(ChannelID) & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td height='22'>" & GetZoneManagePath(ChannelID) & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_Advertisement.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "    <td>"
    Response.Write "      <table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "        <tr class='title' height='22'>"
    Response.Write "          <td width='30' align='center'><strong>ѡ��</strong></td>"
    Response.Write "          <td width='30' align='center'><strong>ID</strong></td>"
    Response.Write "          <td align='center'><strong>��λ����</strong></td>"
    Response.Write "          <td width='65' align='center'><strong>��λ����</strong></td>"
    Response.Write "          <td width='35' align='center'><strong>��ʾ</strong></td>"
    Response.Write "          <td width='65' align='center'><strong>��λ�ߴ�</strong></td>"
    Response.Write "          <td width='30' align='center'><strong>�</strong></td>"
    Response.Write "          <td width='120' align='center'><strong>����</strong></td>"
    Response.Write "          <td width='80' align='center'><strong>��λJS</strong></td>"
    Response.Write "        </tr>"

    sqlZone = "select * from PE_AdZone where 1=1"
    If ChannelID = -2 Then
        sqlZone = sqlZone & " and ChannelID = 0" 
    End If	
    If ChannelID >= -1 Then
        sqlZone = sqlZone & " and ChannelID=" & ChannelID
    End If
    If Keyword <> "" Then
        Select Case strField
        Case "ZoneName"
            sqlZone = sqlZone & " and ZoneName like '%" & Keyword & "%' "
        Case "ZoneIntro"
            sqlZone = sqlZone & " and ZoneIntro like '%" & Keyword & "%' "
        Case Else
            sqlZone = sqlZone & " and ZoneName like '%" & Keyword & "%' "
        End Select
    End If
    sqlZone = sqlZone & " order by ZoneID desc"

    Set rsZone = Server.CreateObject("ADODB.Recordset")
    rsZone.Open sqlZone, Conn, 1, 1
    If rsZone.BOF And rsZone.EOF Then
        Response.Write "        <tr class='tdbg'><td colspan='20' align='center'><br>û���κι���λ��<br><br></td></tr>"
    Else
        totalPut = rsZone.RecordCount
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
                rsZone.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Dim ZoneNum
        ZoneNum = 0
        Do While Not rsZone.EOF
            Response.Write "        <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "          <td width='30' align='center'><input name='ZoneID' type='checkbox' onclick='CheckItem(this)' value='" & rsZone("ZoneID") & "'></td>"
            Response.Write "          <td width='30' align='center'>" & rsZone("ZoneID") & "</td>"
            Response.Write "          <td>"
            If ChannelID = -2 Then
                Response.Write "[" & GetChannelName(rsZone("ChannelID")) & "]"
            End If
            Response.Write "            <a href='Admin_Advertisement.asp?Action=ADList&ZoneID=" & rsZone("ZoneID") & "' title='" & rsZone("ZoneIntro") & "'>" & rsZone("ZoneName") & "</a>"
            Response.Write "          </td>"
            Response.Write "          <td width='65' align='center'>" & ZoneConfig(rsZone("ZoneType"), 1) & "</td>"
            Response.Write "          <td width='35' align='center'>"
            If rsZone("ShowType") = 2 Then
                Response.Write "����"
            ElseIf rsZone("ShowType") = 3 Then
                Response.Write "ѭ��"
            Else
                Response.Write "���"
            End If
            Response.Write "          </td>"
            Response.Write "          <td width='65' align='center'>" & rsZone("ZoneWidth") & " x " & rsZone("ZoneHeight") & "</td>"
            Response.Write "          <td width='30' align='center'>"
            If rsZone("Active") = True Then
                Response.Write "<b>��</b>"
            Else
                Response.Write "<font color=red><b>��</b></font>"
            End If
            Response.Write "          </td>"
            Response.Write "          <td width='120' align='center'>"
            Response.Write "            <a href='Admin_Advertisement.asp?Action=AddAD&ZoneID=" & rsZone("ZoneID") & "'>���</a>&nbsp;"
            Response.Write "            <a href='Admin_Advertisement.asp?Action=ModifyZone&ZoneID=" & rsZone("ZoneID") & "'>�޸�</a>&nbsp;"
            Response.Write "            <a href='Admin_Advertisement.asp?Action=CopyZone&ZoneID=" & rsZone("ZoneID") & "'>����</a>"
            Response.Write "<br>"
            Response.Write "            <a href='Admin_Advertisement.asp?Action=DelZone&ZoneID=" & rsZone("ZoneID") & "' onClick=""return confirm('ȷ��Ҫɾ���˰�λ��');"">ɾ��</a>&nbsp;"
            Response.Write "            <a href='Admin_Advertisement.asp?Action=ClearZone&ZoneID=" & rsZone("ZoneID") & "' onClick=""return confirm('ȷ��Ҫ��մ˰�λ����պ�ԭ�������ڴ˰�λ�Ĺ�潫�������ڰ�λ��');"">���</a>&nbsp;"
            If rsZone("Active") = False Then
                Response.Write "            <a href='Admin_Advertisement.asp?Action=SetZoneActive&ZoneID=" & rsZone("ZoneID") & "'>�</a>"
            Else
                Response.Write "            <a href='Admin_Advertisement.asp?Action=CancelZoneActive&ZoneID=" & rsZone("ZoneID") & "'>��ͣ</a>"
            End If
            Response.Write "          </td>"
            Response.Write "          <td width='80' align='center'>"
            Response.Write "            <a href='Admin_Advertisement.asp?Action=CreateJSZone&ZoneID=" & rsZone("ZoneID") & "'>ˢ��</a>&nbsp;"
            Response.Write "            <a href='Admin_Advertisement.asp?Action=PreviewZone&ZoneID=" & rsZone("ZoneID") & "'>Ԥ��</a>"
            Response.Write "<br>"
            Response.Write "            <a href='Admin_Advertisement.asp?Action=ZoneJSCode&ZoneID=" & rsZone("ZoneID") & "'>JS���ô���</a>"
            Response.Write "          </td>"
            Response.Write "        </tr>"

            ZoneNum = ZoneNum + 1
            If ZoneNum >= MaxPerPage Then Exit Do
            rsZone.MoveNext
        Loop
    End If
    rsZone.Close
    Set rsZone = Nothing
    Response.Write "      </table>"
    Response.Write "      <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "        <tr>"
    Response.Write "          <td width='130' height='30'>"
    Response.Write "            <input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>ѡ�����еİ�λ"
    Response.Write "          </td>"
    Response.Write "          <td>"
    Response.Write "            <input type='submit' value='ɾ��ѡ����λ' name='submit' onClick=""document.myform.Action.value='DelZone'"">&nbsp;"
    Response.Write "            <input type='submit' value='��Ϊ���λ' name='submit' onClick=""document.myform.Action.value='SetZoneActive'"">&nbsp;"
    Response.Write "            <input type='submit' value='��ͣ��λ��ʾ' name='submit' onClick=""document.myform.Action.value='CancelZoneActive'"">&nbsp;"
    Response.Write "            <input type='submit' value='ˢ�°�λJS' name='submit' onClick=""document.myform.Action.value='CreateJSZone'"">&nbsp;"
    Response.Write "            <input type='submit' value='�ƶ���λ ->' name='submit' onClick=""document.myform.Action.value='MoveZone'""><select name='ChannelID' id='ChannelID'>" & GetChannel_Option(-1) & "</select>"
    Response.Write "            <input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </form>"
    Response.Write "  </tr>"
    Response.Write "</table>"

    If totalPut > 0 Then
        Response.Write ShowPage(strFileName & "&ChannelID=" & ChannelID, totalPut, MaxPerPage, CurrentPage, True, True, "����λ", True)
    End If
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='80' align='right'><strong>��λ������</strong></td>"
    Response.Write "   <td>" & GetZoneSearchForm(FileName) & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br><b>˵����</b><br>&nbsp;&nbsp;&nbsp;&nbsp;�����ã�������ò��������ƹ���λJS���룬Ȼ����ģ������Ӧ��λ�ò��뼴�ɡ�<br><br>"
End Sub

Function GetZoneSearchForm(Action)
    Dim strForm
    strForm = "<table border='0' cellpadding='0' cellspacing='0'>"
    strForm = strForm & "<form method='Get' name='SearchForm' action='" & Action & "'>"
    strForm = strForm & "<tr><td height='28' align='center'> "
    strForm = strForm & "<select name='Field' size='1'>"
    strForm = strForm & "<option value='ZoneName' selected>��λ����</option>"
    strForm = strForm & "<option value='ZoneIntro'>��λ����</option>"
    strForm = strForm & "</select> "
    strForm = strForm & "<input type='text' name='keyword'  size='20' value='�ؼ���' maxlength='50' onFocus='this.select();'> "
    strForm = strForm & "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    strForm = strForm & "<input name='Action' type='hidden' id='Action' value='ZoneList'>"
    strForm = strForm & "<input type='submit' name='Submit'  value='����'>"
    strForm = strForm & "</td></tr></form></table>"
    GetZoneSearchForm = strForm
End Function

Sub ShowJS_Zone()
    Response.Write "<script language=JavaScript>" & vbCrLf
    Response.Write "<!--" & vbCrLf
    Response.Write "function CheckForm()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  if(myform.ZoneName.value==''){" & vbCrLf
    Response.Write "    alert('��λ���Ʋ���Ϊ�գ�');" & vbCrLf
    Response.Write "    myform.ZoneName.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(document.myform.ZoneType[document.myform.ZoneType.length-1].checked == false){" & vbCrLf
    Response.Write "    if(myform.ZoneWidth.value==''){" & vbCrLf
    Response.Write "      alert('��λ��Ȳ���Ϊ�գ�');" & vbCrLf
    Response.Write "      myform.ZoneWidth.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if(myform.ZoneHeight.value==''){" & vbCrLf
    Response.Write "      alert('��λ�߶Ȳ���Ϊ�գ�');" & vbCrLf
    Response.Write "      myform.ZoneHeight.focus();" & vbCrLf
    Response.Write "      return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "function Change_Setting()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  if(document.myform.ZoneType[0].checked == false) {" & vbCrLf
    Response.Write "    document.myform.ShowType[2].disabled = true;" & vbCrLf
    Response.Write "    if (document.myform.ShowType[2].checked == true)" & vbCrLf
    Response.Write "    document.myform.ShowType[0].checked = true;" & vbCrLf
    Response.Write "  } else" & vbCrLf
    Response.Write "    document.myform.ShowType[2].disabled = false;" & vbCrLf
    Response.Write "  if(document.myform.ZoneType[document.myform.ZoneType.length-4].checked == false)" & vbCrLf
    Response.Write "    Zone_EnableSize();" & vbCrLf
    Response.Write "  else" & vbCrLf
    Response.Write "    Zone_DisableSize();" & vbCrLf
    Response.Write "  for (var j=0;j<document.myform.ZoneType.length;j++){" & vbCrLf
    Response.Write "    var ot = eval('document.all.ZoneType' + (j + 1) + '_Setting');" & vbCrLf
    Response.Write "    if(document.myform.ZoneType[j].checked==true)" & vbCrLf
    Response.Write "      ot.style.display = '';" & vbCrLf
    Response.Write "    else" & vbCrLf
    Response.Write "      ot.style.display = 'none';" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function Zone_DisableSize()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  document.myform.SizeType[0].disabled = true;" & vbCrLf
    Response.Write "  document.myform.SizeType[1].disabled = true;" & vbCrLf
    Response.Write "  document.myform.ZoneSize.disabled = true;" & vbCrLf
    Response.Write "  document.myform.ZoneWidth.disabled = true;" & vbCrLf
    Response.Write "  document.myform.ZoneHeight.disabled = true;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function Zone_EnableSize()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  document.myform.SizeType[0].disabled = false;" & vbCrLf
    Response.Write "  document.myform.SizeType[1].disabled = false;" & vbCrLf
    Response.Write "  document.myform.ZoneSize.disabled = false;" & vbCrLf
    Response.Write "  document.myform.ZoneWidth.disabled = false;" & vbCrLf
    Response.Write "  document.myform.ZoneHeight.disabled = false;" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function Zone_SelectSize(o)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  size = o.options[o.selectedIndex].value;" & vbCrLf
    Response.Write "  if (size != '0x0')" & vbCrLf
    Response.Write "  {" & vbCrLf
    Response.Write "    sarray = size.split('x');" & vbCrLf
    Response.Write "    height = sarray.pop();" & vbCrLf
    Response.Write "    width  = sarray.pop();" & vbCrLf
    Response.Write "    document.myform.ZoneWidth.value = width;" & vbCrLf
    Response.Write "    document.myform.ZoneHeight.value = height;" & vbCrLf
    Response.Write "    document.myform.SizeType[0].checked = true;" & vbCrLf
    Response.Write "    document.myform.SizeType[1].checked = false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else" & vbCrLf
    Response.Write "  {" & vbCrLf
    Response.Write "    document.myform.SizeType[0].checked = false;" & vbCrLf
    Response.Write "    document.myform.SizeType[1].checked = true;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function Zone_EditSize()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  document.myform.SizeType[0].checked = false;" & vbCrLf
    Response.Write "  document.myform.SizeType[1].checked = true;" & vbCrLf
    Response.Write "  document.myform.ZoneSize.selectedIndex = document.myform.ZoneSize.options.length - 1;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub AddZone()
    Dim i
    Call ShowJS_Zone
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Advertisement.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>�� �� �� λ</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>����Ƶ�����ࣺ</strong><br>�˷���ֻ�������ְ�λ���ڵ�λ�á�</td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <select name='ChannelID' id='ChannelID'>" & GetChannel_Option(-1) & "</select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>��λ���ƣ�</strong></td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <input name='ZoneName' type='text' id='ZoneName' size='60' maxlength='60' value=''> <font color='red'>*</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>����JS�ļ�����</strong></td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <input name='ZoneJSName' type='text' id='ZoneJSName' size='60' maxlength='100' value='" & GetCurrentZoneJSName() & "'> <font color='red'>*</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>��λ������</strong></td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <textarea name='ZoneIntro' cols='50' rows='3' id='ZoneIntro'></textarea>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>��λ���ͣ�</strong><br>ѡ������ڴ˰�λ�Ĺ�����͡�</td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <table>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    For i = 1 To ZoneTypeNum
        Response.Write "              <input type='radio' name='ZoneType' value='" & i & "' onclick='Change_Setting();' " & IsRadioChecked(1, i) & "> " & ZoneConfig(i, 1) & "&nbsp;"
        If i Mod 5 = 0 Then Response.Write "<br>"
    Next
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' valign='top'><strong>��λ���ã�</strong><br>�԰�λ����ϸ�����������á�</td>"
    Response.Write "      <td width='600' valign='top'>"
    Response.Write "        <table width='100%' height='40' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <input name='DefaultSetting' type='radio' value='1' onClick=""ZoneSetting.style.display='none'"" checked> Ĭ������&nbsp;"
    Response.Write "              <input name='DefaultSetting' type='radio' value='0' onClick=""ZoneSetting.style.display=''""> �Զ�������&nbsp;"
    Response.Write "            </td>"
    Response.Write "          <tr>"
    Response.Write "        </table>"
    Response.Write "        <table id='ZoneSetting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:none'>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <table id='ZoneType1_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(1, 2), 1)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType2_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:none'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(2, 2), 2)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType3_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:none'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(3, 2), 3)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType4_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:none'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(4, 2), 4)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType5_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:none'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(5, 2), 5)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType6_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:none'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(6, 2), 6)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType7_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:none'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(7, 2), 7)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType8_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:none'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(8, 2), 8)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType9_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:none'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(9, 2), 9)
    Response.Write "              </td></tr></table>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>��λ�ߴ磺</strong><br>IAB��������������ϻ��׼�ߴ硣<br>��*�ŵ�Ϊ�����ӵı�׼���ߴ硣</td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <table>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <input type='radio' name='SizeType' value='default' checked>"
    Response.Write "              <select name='ZoneSize' onchange='Zone_SelectSize(this)'>"
    For i = 0 To 15
        Response.Write "<option value='" & IAB_Size(i, 0) & "' " & IsOptionSelected(i, 0) & ">" & IAB_Size(i, 1) & "</option>"
    Next
    Response.Write "<option value='0x0'>�Զ����С</option>"
    Response.Write "              </select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <input type='radio' name='SizeType' value='custom' onclick='Zone_EditSize()'>"
    Response.Write "              ���: "
    Response.Write "              <input name='ZoneWidth' size='5' maxlength='4' onkeydown='Zone_EditSize()' value='468'>&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "              �߶�:"
    Response.Write "              <input name='ZoneHeight' size='5' maxlength='4' onkeydown='Zone_EditSize()' value='60'>"
    Response.Write "              <font color='red'>*</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>��ʾ��ʽ��</strong><br>����λ���ж�����ʱ���մ��趨������ʾ�����ݹ���Ȩ�أ���</td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <input name='ShowType' type='radio' value='1' checked> ��Ȩ�������ʾ&nbsp;&nbsp;Ȩ��Խ����ʾ����Խ��<br>"
    Response.Write "        <input name='ShowType' type='radio' value='2'> ��Ȩ��������ʾ&nbsp;&nbsp;��ʾȨ��ֵ���Ĺ�档<br>"
    Response.Write "        <input name='ShowType' type='radio' value='3'> ��˳��ѭ����ʾ&nbsp;&nbsp;�˷�ʽ���Ծ��κ����Ч��"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>��λ״̬��</strong><br>��Ϊ��İ�λ������ǰ̨��ʾ��</td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <input name='Active' type='checkbox' id='Active' value='yes' checked> ���λ"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveAddZone'>"
    Response.Write "        <input  type='submit' name='Submit' value=' �� �� '>&nbsp;&nbsp;"
    Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Advertisement.asp?Action=ZoneList'"" style='cursor:hand;'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub ModifyZone()
    If ZoneID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵİ�λID</li>"
        Exit Sub
    Else
        ZoneID = PE_CLng(ZoneID)
    End If
    Dim rsZone, sqlZone
    sqlZone = "select * from PE_AdZone where ZoneID=" & ZoneID
    Set rsZone = Conn.Execute(sqlZone)
    If rsZone.BOF And rsZone.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���İ�λ</li>"
        rsZone.Close
        Set rsZone = Nothing
        Exit Sub
    End If

    Dim ZoneSize, IsIABSize, strDisabled, i
    ZoneSize = rsZone("ZoneWidth") & "x" & rsZone("ZoneHeight")
    IsIABSize = False
    For i = 0 To 15
        If ZoneSize = IAB_Size(i, 0) Then
            IsIABSize = True
        End If
    Next
    If rsZone("ZoneType") = 4 Then strDisabled = " disabled"
    ZoneConfig(rsZone("ZoneType"), 2) = rsZone("ZoneSetting")

    Call ShowJS_Zone
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Advertisement.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>�� �� �� λ</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>����Ƶ�����ࣺ</strong><br>�˷���ֻ�������ְ�λ���ڵ�λ�á�</td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <select name='ChannelID' id='ChannelID'>" & GetChannel_Option(rsZone("ChannelID")) & "</select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>��λ���ƣ�</strong></td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <input name='ZoneName' type='text' id='ZoneName' size='60' maxlength='60' value='" & rsZone("ZoneName") & "'> <font color='red'>*</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>����JS�ļ�����</strong></td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <input name='ZoneJSName' type='text' id='ZoneJSName' size='60' maxlength='100' value='" & rsZone("ZoneJSName") & "'> <font color='red'>*</font>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>��λ������</strong></td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <textarea name='ZoneIntro' cols='50' rows='3' id='ZoneIntro'>" & PE_ConvertBR(rsZone("ZoneIntro")) & "</textarea>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>��λ���ͣ�</strong><br>ѡ������ڴ˰�λ�Ĺ�����͡�</td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <table>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    For i = 1 To ZoneTypeNum
        Response.Write "              <input type='radio' name='ZoneType' value='" & i & "' onclick='Change_Setting();' " & IsRadioChecked(rsZone("ZoneType"), i) & "> " & ZoneConfig(i, 1) & "&nbsp;"
        If i Mod 5 = 0 Then Response.Write "<br>"
    Next
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200' valign='top'><strong>��λ���ã�</strong><br>�԰�λ����ϸ�����������á�</td>"
    Response.Write "      <td width='600' valign='top'>"
    Response.Write "        <table width='100%' height='40' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <input name='DefaultSetting' type='radio' value='1' onClick=""ZoneSetting.style.display='none'"" " & IsRadioChecked(rsZone("DefaultSetting"), True) & "> Ĭ������&nbsp;"
    Response.Write "              <input name='DefaultSetting' type='radio' value='0' onClick=""ZoneSetting.style.display=''"" " & IsRadioChecked(rsZone("DefaultSetting"), False) & "> �Զ�������&nbsp;"
    Response.Write "            </td>"
    Response.Write "          <tr>"
    Response.Write "        </table>"

    Response.Write "        <table id='ZoneSetting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:" & StyleDisplay(rsZone("DefaultSetting"), False) & "'>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <table id='ZoneType1_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:" & StyleDisplay(rsZone("ZoneType"), 1) & "'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(1, 2), 1)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType2_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:" & StyleDisplay(rsZone("ZoneType"), 2) & "'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(2, 2), 2)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType3_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:" & StyleDisplay(rsZone("ZoneType"), 3) & "'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(3, 2), 3)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType4_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:" & StyleDisplay(rsZone("ZoneType"), 4) & "'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(4, 2), 4)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType5_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:" & StyleDisplay(rsZone("ZoneType"), 5) & "'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(5, 2), 5)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType6_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:" & StyleDisplay(rsZone("ZoneType"), 6) & "'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(6, 2), 6)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType7_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:" & StyleDisplay(rsZone("ZoneType"), 7) & "'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(7, 2), 7)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType8_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:" & StyleDisplay(rsZone("ZoneType"), 8) & "'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(8, 2), 8)
    Response.Write "              </td></tr></table>"
    Response.Write "              <table id='ZoneType9_Setting' width='100%' border='0' cellpadding='0' cellspacing='0' style='display:" & StyleDisplay(rsZone("ZoneType"), 9) & "'><tr><td>"
    Call ShowZoneSetting(ZoneConfig(9, 2), 9)
    Response.Write "              </td></tr></table>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"

    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>��λ�ߴ磺</strong><br>IAB��������������ϻ��׼�ߴ硣<br>��*�ŵ�Ϊ�����ӵı�׼���ߴ硣</td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <table>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <input type='radio' name='SizeType' value='default' " & IsRadioChecked(IsIABSize, True) & ">"
    Response.Write "              <select name='ZoneSize' onchange='Zone_SelectSize(this)' " & strDisabled & ">"
    For i = 0 To 15
        Response.Write "<option value='" & IAB_Size(i, 0) & "' " & IsOptionSelected(ZoneSize, IAB_Size(i, 0)) & ">" & IAB_Size(i, 1) & "</option>"
    Next
    Response.Write "<option value='0x0' " & IsOptionSelected(IsIABSize, False) & ">�Զ����С</option>"
    Response.Write "              </select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <input type='radio' name='SizeType' value='custom' onclick='Zone_EditSize()' " & IsRadioChecked(IsIABSize, False) & ">"
    Response.Write "              ���: "
    Response.Write "              <input name='ZoneWidth' size='5' maxlength='4' onkeydown='Zone_EditSize()' value='" & rsZone("ZoneWidth") & "' " & strDisabled & ">&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "              �߶�:"
    Response.Write "              <input name='ZoneHeight' size='5' maxlength='4' onkeydown='Zone_EditSize()' value='" & rsZone("ZoneHeight") & "' " & strDisabled & ">"
    Response.Write "              <font color='red'>*</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>��ʾ��ʽ��</strong><br>����λ���ж�����ʱ���մ��趨������ʾ�����ݹ���Ȩ�أ���</td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <input name='ShowType' type='radio' value='1' " & IsRadioChecked(rsZone("ShowType"), 1) & "> ��Ȩ�������ʾ&nbsp;&nbsp;Ȩ��Խ����ʾ����Խ��<br>"
    Response.Write "        <input name='ShowType' type='radio' value='2' " & IsRadioChecked(rsZone("ShowType"), 2) & "> ��Ȩ��������ʾ&nbsp;&nbsp;��ʾȨ��ֵ���Ĺ�档<br>"
    Response.Write "        <input name='ShowType' type='radio' value='3' " & IsRadioChecked(rsZone("ShowType"), 3)
    If rsZone("ZoneType") <> 1 Then Response.Write " disabled"
    Response.Write "> ��˳��ѭ����ʾ&nbsp;&nbsp;�˷�ʽ���Ծ��κ����Ч��"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='200'><strong>��λ״̬��</strong><br>��Ϊ��İ�λ������ǰ̨��ʾ��</td>"
    Response.Write "      <td width='600'>"
    Response.Write "        <input name='Active' type='checkbox' id='Active' value='yes' " & IsRadioChecked(rsZone("Active"), True) & "> ���λ"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='ZoneID' type='hidden' id='ZoneID' value='" & rsZone("ZoneID") & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModifyZone'>"
    Response.Write "        <input  type='submit' name='Submit' value=' �� �� '>&nbsp;&nbsp;"
    Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Advertisement.asp?Action=ZoneList'"" style='cursor:hand;'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
    rsZone.Close
    Set rsZone = Nothing
End Sub

Sub ShowZoneSetting(Setting, ZoneType)
    Select Case ZoneType
    Case 1
        Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff'>"
        Response.Write "  <tr align='center' class='tdbg2'>"
        Response.Write "    <td colspan='2'><strong>��λ��������--" & ZoneConfig(1, 1) & "</strong></td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td colspan='2' align='center'>�������ް�λ�������ã�</td>"
        Response.Write "  </tr>"
        Response.Write "</table>"
    Case 2
        Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff'>"
        Response.Write "  <tr align='center' class='tdbg2'>"
        Response.Write "    <td colspan='2'><strong>��λ��������--" & ZoneConfig(2, 1) & "</strong></td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='200'><strong>������ʽ��</strong></td>"
        Response.Write "    <td>"
        Response.Write "      <select name='Pop_PopType'>"
        Response.Write "        <option value='1' " & IsOptionSelected(GetSettingItem(Setting, 2, 1), "1") & ">ǰ�ô���</option>"
        Response.Write "        <option value='2' " & IsOptionSelected(GetSettingItem(Setting, 2, 1), "2") & ">���ô���</option>"
        Response.Write "        <option value='3' " & IsOptionSelected(GetSettingItem(Setting, 2, 1), "3") & ">��ҳ�Ի���</option>"
        Response.Write "        <option value='4' " & IsOptionSelected(GetSettingItem(Setting, 2, 1), "4") & ">��Ͷ���</option>"
        Response.Write "      </select>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='200'><strong>����λ�ã��󣩣�</strong></td>"
        Response.Write "    <td>"
        Response.Write "      <input name='Pop_Left' type='text' id='Pop_Left' size='5' maxlength='4' value='" & GetSettingItem(Setting, 2, 2) & "'>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='200'><strong>����λ�ã��ϣ���</strong></td>"
        Response.Write "    <td>"
        Response.Write "      <input name='Pop_Top' type='text' id='Pop_Top' size='5' maxlength='4' value='" & GetSettingItem(Setting, 2, 3) & "'>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='200'><strong>ʱ������</strong><br>��ʱ�����ڲ��ظ�������</td>"
        Response.Write "    <td>"
        Response.Write "      <input name='Pop_CookieHour' type='text' id='Pop_CookieHour' size='5' maxlength='2' value='" & GetSettingItem(Setting, 2, 4) & "'> Сʱ ����<font color='blue'>��Ϊ0ʱ���ǵ���</font>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "</table>"
    Case 3
        Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff'>"
        Response.Write "  <tr align='center' class='tdbg2'>"
        Response.Write "    <td colspan='2'><strong>��λ��������--" & ZoneConfig(3, 1) & "</strong></td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='200'><strong>���λ�ã��󣩣�</strong></td>"
        Response.Write "    <td>"
        Response.Write "      <input name='Move_Left' type='text' id='Move_Left' size='5' maxlength='4' value='" & GetSettingItem(Setting, 3, 1) & "'>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='200'><strong>���λ�ã��ϣ���</strong></td>"
        Response.Write "    <td>"
        Response.Write "      <input name='Move_Top' type='text' id='Move_Top' size='5' maxlength='4' value='" & GetSettingItem(Setting, 3, 2) & "'>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='200'><strong>�ƶ�ƽ���ȣ�</strong></td>"
        Response.Write "    <td>"
        Response.Write "      <input name='Move_Delta' type='text' id='Move_Delta' size='7' maxlength='7' value='" & GetSettingItem(Setting, 3, 3) & "'> ��ȡֵ��0.001��1֮�䣩"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "</table>"
    Case 4
        Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff'>"
        Response.Write "  <tr align='center' class='tdbg2'>"
        Response.Write "    <td colspan='2'><strong>��λ��������--" & ZoneConfig(4, 1) & "</strong></td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='200'><strong>���λ�ã��󣩣�</strong></td>"
        Response.Write "    <td>"
        Response.Write "      <input name='Fixed_Left' type='text' id='Fixed_Left' size='5' maxlength='4' value='" & GetSettingItem(Setting, 4, 1) & "'>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='200'><strong>���λ�ã��ϣ���</strong></td>"
        Response.Write "    <td>"
        Response.Write "      <input name='Fixed_Top' type='text' id='Fixed_Top' size='5' maxlength='4' value='" & GetSettingItem(Setting, 4, 2) & "'>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "</table>"
    Case 5
        Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff'>"
        Response.Write "  <tr align='center' class='tdbg2'>"
        Response.Write "    <td colspan='2'><strong>��λ��������--" & ZoneConfig(5, 1) & "</strong></td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='200'><strong>Ư�����ͣ�</strong></td>"
        Response.Write "    <td>"
        Response.Write "      <select name='Float_Type'>"
        Response.Write "        <option value='1' " & IsOptionSelected(GetSettingItem(Setting, 5, 1), "1") & ">����Ư��</option>"
        Response.Write "        <option value='2' " & IsOptionSelected(GetSettingItem(Setting, 5, 1), "2") & ">����Ư��</option>"
        Response.Write "        <option value='3' " & IsOptionSelected(GetSettingItem(Setting, 5, 1), "3") & ">����Ư��</option>"
        Response.Write "        <option value='4' " & IsOptionSelected(GetSettingItem(Setting, 5, 1), "4") & ">����Ư��</option>"
        Response.Write "      </select>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='200'><strong>��ʼλ�ã��󣩣�</strong></td>"
        Response.Write "    <td>"
        Response.Write "      <input name='Float_Left' type='text' id='Float_Left' size='5' maxlength='4' value='" & GetSettingItem(Setting, 5, 2) & "'>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='200'><strong>��ʼλ�ã��ϣ���</strong></td>"
        Response.Write "    <td>"
        Response.Write "      <input name='Float_Top' type='text' id='Float_Top' size='5' maxlength='4' value='" & GetSettingItem(Setting, 5, 3) & "'>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "</table>"
    Case 6
        Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff'>"
        Response.Write "  <tr align='center' class='tdbg2'>"
        Response.Write "    <td colspan='2'><strong>��λ��������--" & ZoneConfig(6, 1) & "</strong></td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td colspan='2' align='center'>�������ް�λ�������ã�</td>"
        Response.Write "  </tr>"
        Response.Write "</table>"
    Case 7
        Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff'>"
        Response.Write "  <tr align='center' class='tdbg2'>"
        Response.Write "    <td colspan='2'><strong>��λ��������--" & ZoneConfig(7, 1) & "</strong></td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td colspan='2' align='center'>�������ް�λ�������ã�</td>"
        Response.Write "  </tr>"
        Response.Write "</table>"
    Case 8
        Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff'>"
        Response.Write "  <tr align='center' class='tdbg2'>"
        Response.Write "    <td colspan='2'><strong>��λ��������--" & ZoneConfig(8, 1) & "</strong></td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td colspan='2' align='center'>�������ް�λ�������ã�</td>"
        Response.Write "  </tr>"
        Response.Write "</table>"
    Case 9
        Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff'>"
        Response.Write "  <tr align='center' class='tdbg2'>"
        Response.Write "    <td colspan='2'><strong>��λ��������--" & ZoneConfig(9, 1) & "</strong></td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td colspan='2' align='center'>�������ް�λ�������ã�</td>"
        Response.Write "  </tr>"
        Response.Write "</table>"
    End Select
End Sub

Sub SaveZone()
    Dim ZoneID, ChannelID, ZoneName, ZoneJSName, ZoneIntro, ZoneType, DefaultSetting, ZoneSetting, ZoneWidth, ZoneHeight, ShowType, Active
    Dim rsZone, sqlZone
    ZoneID = PE_CLng(Trim(Request.Form("ZoneID")))
    ChannelID = PE_CLng(Trim(Request.Form("ChannelID")))
    ZoneName = Trim(Request.Form("ZoneName"))
    ZoneJSName = Trim(Request.Form("ZoneJSName"))
    ZoneIntro = Trim(Request.Form("ZoneIntro"))
    ZoneType = PE_CLng(Trim(Request.Form("ZoneType")))
    DefaultSetting = CBool(Trim(Request.Form("DefaultSetting")))
    ZoneWidth = PE_CLng(Trim(Request.Form("ZoneWidth")))
    ZoneHeight = PE_CLng(Trim(Request.Form("ZoneHeight")))
    ShowType = PE_CLng(Trim(Request.Form("ShowType")))
    Active = Trim(Request.Form("Active"))

    If ZoneName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��λ���Ʋ���Ϊ�գ�</li>"
    End If
    If CheckZoneJSName(ZoneJSName) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�����JS�ļ���������Ҫ��<br>&nbsp;&nbsp;�ļ�����ֻ�ܰ���Ӣ����ĸ�����֡��»��߼���-������<br>&nbsp;&nbsp;ֻ֧��һ��·��������Ҫʹ�����·����ַ</li>"
    End If
    If FoundErr = True Then Exit Sub

    If DefaultSetting = True Then
        ZoneSetting = ZoneConfig(ZoneType, 2)
    Else
        Select Case ZoneType
        Case 1
            ZoneSetting = ZoneType
        Case 2
            ZoneSetting = ZoneType & "," & RequestSetting("Pop_PopType") & "," & RequestSetting("Pop_Left") & "," & RequestSetting("Pop_Top") & "," & RequestSetting("Pop_CookieHour")
        Case 3
            ZoneSetting = ZoneType & "," & RequestSetting("Move_Left") & "," & RequestSetting("Move_Top") & "," & RequestSetting("Move_Delta")
        Case 4
            ZoneSetting = ZoneType & "," & RequestSetting("Fixed_Left") & "," & RequestSetting("Fixed_Top")
        Case 5
            ZoneSetting = ZoneType & "," & RequestSetting("Float_Type") & "," & RequestSetting("Float_Left") & "," & RequestSetting("Float_Top")
        Case 6
            ZoneSetting = ZoneType
        End Select
    End If

    ZoneName = PE_HTMLEncode(ZoneName)
    ZoneIntro = PE_HTMLEncode(ZoneIntro)
    ZoneSetting = PE_HTMLEncode(ZoneSetting)
    Active = CBool(Active = "yes")
    If (ShowType = 3 And ZoneType <> 1) Or ShowType = 0 Then ShowType = 1

    Set rsZone = Server.CreateObject("adodb.recordset")
    If Action = "SaveAddZone" Then
        sqlZone = "select top 1 * from PE_AdZone"
        rsZone.Open sqlZone, Conn, 1, 3
        rsZone.addnew
        ZoneID = PE_CLng(Conn.Execute("select max(ZoneID) from PE_AdZone")(0)) + 1
        rsZone("ZoneID") = ZoneID
        rsZone("UpdateTime") = Now()
    ElseIf Action = "SaveModifyZone" Then
        If ZoneID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����ȷ����λID��ֵ</li>"
            Exit Sub
        End If
        sqlZone = "select * from PE_AdZone where ZoneID=" & ZoneID
        rsZone.Open sqlZone, Conn, 1, 3
        If rsZone.BOF And rsZone.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�Ҳ���ָ���İ�λ��</li>"
            rsZone.Close
            Set rsZone = Nothing
            Exit Sub
        End If
    End If
    rsZone("ChannelID") = ChannelID
    rsZone("ZoneName") = ZoneName
    rsZone("ZoneJSName") = ZoneJSName
    rsZone("ZoneIntro") = ZoneIntro
    rsZone("ZoneType") = ZoneType
    rsZone("DefaultSetting") = DefaultSetting
    rsZone("ZoneSetting") = ZoneSetting
    rsZone("ZoneWidth") = ZoneWidth
    rsZone("ZoneHeight") = ZoneHeight
    rsZone("ShowType") = ShowType
    rsZone("Active") = Active
    rsZone.Update
    rsZone.Close
    Set rsZone = Nothing

    Call WriteEntry(2, AdminName, "�������λ���óɹ���" & ZoneName)

    Call CreateJSZoneID(ZoneID)

    Call ClearSiteCache(0)
    Call CloseConn
    Response.Redirect "Admin_Advertisement.asp?Action=ZoneList&ChannelID=" & ChannelID
End Sub

Sub SetZoneProperty()
    If ZoneID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ����λID</li>"
        Exit Sub
    End If
    If Action = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
        Exit Sub
    End If

    Dim sqlProperty, rsProperty
    Dim MoveChannelID
    If InStr(ZoneID, ",") > 0 Then
        sqlProperty = "select * from PE_AdZone where ZoneID in (" & ZoneID & ")"
    Else
        sqlProperty = "select * from PE_AdZone where ZoneID=" & ZoneID
    End If
    Set rsProperty = Server.CreateObject("ADODB.Recordset")
    rsProperty.Open sqlProperty, Conn, 1, 3
    Do While Not rsProperty.EOF
        Select Case Action
        Case "SetZoneActive"
            rsProperty("Active") = True
        Case "CancelZoneActive"
            rsProperty("Active") = False
        Case "MoveZone"
            MoveChannelID = Trim(Request("ChannelID"))
            If MoveChannelID = "" Then
                MoveChannelID = -1
            Else
                MoveChannelID = PE_CLng(MoveChannelID)
            End If
            rsProperty("ChannelID") = MoveChannelID
        Case "DelZone"
            Call DelZoneID_AD(rsProperty("IncludeADID"), rsProperty("ZoneID"))
            Dim ZoneJSFile
            ZoneJSFile = GetZoneJSName(rsProperty("ZoneJSName"), rsProperty("ZoneID"), rsProperty("UpdateTime"))
            If fso.FileExists(Server.MapPath(ZoneJSFile)) Then
                fso.DeleteFile Server.MapPath(ZoneJSFile)
            End If
            rsProperty.Delete
        End Select
        rsProperty.Update
        rsProperty.MoveNext
    Loop
    rsProperty.Close
    Set rsProperty = Nothing

    If Action = "SetZoneActive" Or Action = "CancelZoneActive" Then
        Call CreateJSZoneID(ZoneID)
    End If
    Call WriteEntry(2, AdminName, "���ù���λ���Գɹ�����λID��" & ZoneID)

    Call ClearSiteCache(0)
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub CopyZone()
    Dim MaxZoneID
    ZoneID = PE_CLng(ZoneID)
    If ZoneID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
        Exit Sub
    End If
    
    MaxZoneID = PE_CLng(Conn.Execute("select max(ZoneID) from PE_AdZone")(0)) + 1
    Conn.Execute ("insert into PE_AdZone select " & MaxZoneID & " as ZoneID,ChannelID,'' as IncludeADID,'���� '+ZoneName as ZoneName,'" & GetCurrentZoneJSName() & "' as ZoneJSName,ZoneIntro,ZoneType,DefaultSetting,ZoneSetting,ZoneWidth,ZoneHeight,Active,ShowType,UpdateTime from PE_AdZone where ZoneID=" & ZoneID)
    Call WriteEntry(2, AdminName, "���ƹ���λ�ɹ�����λID��" & ZoneID)
    Response.Redirect ComeUrl
End Sub

Sub ClearZone()
    If ZoneID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
        Exit Sub
    Else
        ZoneID = PE_CLng(ZoneID)
    End If

    Dim rs, IncludeADID
    Set rs = Conn.Execute("select IncludeADID from PE_AdZone where ZoneID=" & ZoneID)
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��λ�����ڣ������Ѿ���ɾ��</li>"
    Else
        IncludeADID = rs(0)
    End If
    rs.Close
    Set rs = Nothing
    If FoundErr = True Then Exit Sub

    Conn.Execute ("update PE_AdZone set IncludeADID='' where ZoneID=" & ZoneID)
    Call DelZoneID_AD(IncludeADID, ZoneID)

    Call CreateJSZoneID(ZoneID)
    Call WriteEntry(2, AdminName, "�������λ�µĹ�棬��λID��" & ZoneID)

    Call ClearSiteCache(0)
    Call WriteSuccessMsg("�Ѿ��ɹ���ձ���λ�µĹ�档", ComeUrl)
End Sub

Sub DelZoneID_AD(arrADID, iZoneID)
    If iZoneID = "" Or IsNull(iZoneID) Then
        Exit Sub
    Else
        iZoneID = PE_CLng(iZoneID)
    End If
    If IsValidID(arrADID) = True Then
        Dim sqlAD, rsAD
        arrADID = ReplaceBadChar(arrADID)
        sqlAD = "select ZoneID from PE_Advertisement where ADID in (" & arrADID & ")"
        Set rsAD = Server.CreateObject("Adodb.RecordSet")
        rsAD.Open sqlAD, Conn, 1, 3
        Do While Not rsAD.EOF
            rsAD(0) = RemoveStr(rsAD(0), CStr(iZoneID), ",")
            rsAD.Update
            rsAD.MoveNext
        Loop
        rsAD.Close
        Set rsAD = Nothing
    End If
End Sub

Sub PreviewZone()
    Dim ID, sqlJs, rsJs
    If ZoneID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������ʧ��</li>"
        Exit Sub
    Else
        ZoneID = PE_CLng(ZoneID)
    End If
    sqlJs = "select ZoneID,IncludeADID,ZoneName,UpdateTime,ZoneJSName from PE_AdZone where ZoneID=" & ZoneID
    Set rsJs = Conn.Execute(sqlJs)
    If rsJs.BOF And rsJs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���İ�λ��</li>"
        rsJs.Close
        Set rsJs = Nothing
        Exit Sub
    End If

    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='2' align='center'><strong>Ԥ����λJSЧ��----" & rsJs("ZoneName") & "</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg2'>"
    Response.Write "    <td height='25' align='center'>"
    Response.Write "      <a href='javascript:this.location.reload();'>ˢ��ҳ��</a>&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "      <a href='" & ComeUrl & "'>������ҳ</a>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr valign='top'>"
    If IsNull(rsJs("IncludeADID")) Or rsJs("IncludeADID") = "" Then
        Response.Write "    <td height='200' align='center'><br><br><br><br>��λ����ʱ��δ��ӹ�棬����Ӻ��ٽ���Ԥ����</td>"
    Else
        Response.Write "    <td height='800'><script language='javascript' src='" & GetZoneJSName(rsJs("ZoneJSName"), rsJs("ZoneID"), rsJs("UpdateTime")) & "'></script></td>"
    End If
    Response.Write "  </tr>"
    Response.Write "</table>"

    rsJs.Close
    Set rsJs = Nothing
End Sub

Sub ZoneJSCode()
    Dim ID, sqlJs, rsJs
    If ZoneID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������ʧ��</li>"
        Exit Sub
    Else
        ZoneID = PE_CLng(ZoneID)
    End If
    sqlJs = "select ZoneID,ZoneName,UpdateTime,ZoneJSName from PE_AdZone where ZoneID=" & ZoneID
    Set rsJs = Conn.Execute(sqlJs)
    If rsJs.BOF And rsJs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���İ�λ��</li>"
        rsJs.Close
        Set rsJs = Nothing
        Exit Sub
    End If

    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='2' align='center'><strong>��λJS���ô���----" & rsJs("ZoneName") & "</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg2'>"
    Response.Write "    <td height='25' align='center'>���÷�����������Ĵ�����뵽��ҳ��Ԥ���Ĺ��λ��</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='100' align='center'>"
    Response.Write "      <textarea name='ZoneJSCode' cols='100' rows='5' id='ZoneJSCode'><script language='javascript' src='" & GetZoneJSNameCode(rsJs("ZoneJSName")) & "'></script></textarea>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='25' align='center'><a href='" & ComeUrl & "'>����</a></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"

    rsJs.Close
    Set rsJs = Nothing
End Sub


Sub CreateJSZone()
    If ZoneID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ����λID</li>"
        Exit Sub
    End If
    Call CreateJSZoneID(ZoneID)
    Call WriteSuccessMsg("ˢ�°�λJS�ɹ���", ComeUrl)
End Sub

Sub CreateJSZoneID(ZoneID)
    If ObjInstalled_FSO = False Then
        Exit Sub
    End If

    Dim arrZoneID, j, rsZone, sqlZone, ZoneJS_Path, ZoneJS_Name, strADJS
    Dim ZoneType, IncludeADID, ShowType, ZoneWidth, ZoneHeight, ZoneSetting, Active

    If IsValidID(ZoneID) = False Then
        Exit Sub
    Else
        ZoneID = Replace(ZoneID, " ", "")
    End If

    arrZoneID = Split(ZoneID, ",")
    For j = 0 To UBound(arrZoneID)
        sqlZone = "select * from PE_AdZone where ZoneID=" & arrZoneID(j)
        Set rsZone = Conn.Execute(sqlZone)
        If Not rsZone.BOF And Not rsZone.EOF Then
            ZoneJS_Path = InstallDir & ADDir & "/" & GetZoneJS_Path(rsZone("ZoneJSName"), rsZone("UpdateTime"))
            ZoneJS_Name = ZoneJS_Path & GetZoneJS_Name(rsZone("ZoneJSName"), rsZone("ZoneID"))
            ZoneType = PE_CLng(rsZone("ZoneType"))
            IncludeADID = rsZone("IncludeADID")
            ShowType = rsZone("ShowType")
            ZoneWidth = rsZone("ZoneWidth")
            ZoneHeight = rsZone("ZoneHeight")
            If rsZone("DefaultSetting") = True Then
                ZoneSetting = ZoneConfig(ZoneType, 2)
            Else
                ZoneSetting = rsZone("ZoneSetting")
            End If
            Active = rsZone("Active")

            strADJS = ""
            If IsValidID(IncludeADID) = True And Active = True Then
                Dim rsAD, sqlAD
                sqlAD = "select * from PE_Advertisement where Passed=" & PE_True & " and ADID in (" & IncludeADID & ") order by Priority desc, ADID desc"
                Set rsAD = Conn.Execute(sqlAD)
                If Not (rsAD.BOF And rsAD.EOF) Then
                    strADJS = strADJS & GetZoneJSTemplate(ZoneType)
                    strADJS = strADJS & vbCrLf
                    strADJS = strADJS & "var ZoneAD_" & arrZoneID(j) & " = new " & ZoneConfig(ZoneType, 0) & "ZoneAD(""ZoneAD_" & arrZoneID(j) & """);" & vbCrLf
                    strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".ZoneID      = " & PE_CLng(arrZoneID(j)) & ";" & vbCrLf
                    strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".ZoneWidth   = " & PE_CLng(ZoneWidth) & ";" & vbCrLf
                    strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".ZoneHeight  = " & PE_CLng(ZoneHeight) & ";" & vbCrLf
                    strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".ShowType    = " & PE_CLng(ShowType) & ";" & vbCrLf
                    Select Case ZoneType
                    Case 1

                    Case 2
                        strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".PopType     = " & PE_CLng(GetSettingItem(ZoneSetting, 2, 1)) & ";" & vbCrLf
                        strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".Left        = " & PE_CLng(GetSettingItem(ZoneSetting, 2, 2)) & ";" & vbCrLf
                        strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".Top         = " & PE_CLng(GetSettingItem(ZoneSetting, 2, 3)) & ";" & vbCrLf
                        strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".CookieHour  = " & PE_CLng(GetSettingItem(ZoneSetting, 2, 4)) & ";" & vbCrLf
                    Case 3
                        strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".Left        = " & PE_CLng(GetSettingItem(ZoneSetting, 3, 1)) & ";" & vbCrLf
                        strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".Top         = " & PE_CLng(GetSettingItem(ZoneSetting, 3, 2)) & ";" & vbCrLf
                        If GetSettingItem(ZoneSetting, 3, 3) <> "" Then
                            strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".Delta       = " & GetSettingItem(ZoneSetting, 3, 3) & ";" & vbCrLf
                        End If
                    Case 4
                        strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".Left        = " & PE_CLng(GetSettingItem(ZoneSetting, 4, 1)) & ";" & vbCrLf
                        strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".Top         = " & PE_CLng(GetSettingItem(ZoneSetting, 4, 2)) & ";" & vbCrLf
                    Case 5
                        strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".FloatType     = " & PE_CLng(GetSettingItem(ZoneSetting, 5, 1)) & ";" & vbCrLf
                        strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".Left        = " & PE_CLng(GetSettingItem(ZoneSetting, 5, 2)) & ";" & vbCrLf
                        strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".Top         = " & PE_CLng(GetSettingItem(ZoneSetting, 5, 3)) & ";" & vbCrLf
                    Case 6

                    End Select
                    Do While Not rsAD.EOF
                        strADJS = strADJS & vbCrLf
                        strADJS = strADJS & "var objAD = new ObjectAD();" & vbCrLf
                        strADJS = strADJS & "objAD.ADID           = " & rsAD("ADID") & ";" & vbCrLf
                        strADJS = strADJS & "objAD.ADType         = " & PE_CLng(rsAD("ADType")) & ";" & vbCrLf
                        strADJS = strADJS & "objAD.ADName         = """ & rsAD("ADName") & """;" & vbCrLf
                        
                        If SiteUrlType=1 then
                        Dim ImgUrl
                        ImgUrl= rsAD("ImgUrl")
                        ImgUrl=Right(ImgUrl,Len(ImgUrl)-Len(InstallDir)+1)
                        strADJS = strADJS & "objAD.ImgUrl         = """ & ImgUrl & """;" & vbCrLf
                        strADJS = strADJS & "objAD.InstallDir     = """ & strInstallDir & """;" & vbCrLf
                        Else
                        strADJS = strADJS & "objAD.ImgUrl         = """ & rsAD("ImgUrl") & """;" & vbCrLf
                        strADJS = strADJS & "objAD.InstallDir     = """ & InstallDir & """;" & vbCrLf
                        End If
                        
                        strADJS = strADJS & "objAD.ImgWidth       = " & PE_CLng(rsAD("ImgWidth")) & ";" & vbCrLf
                        strADJS = strADJS & "objAD.ImgHeight      = " & PE_CLng(rsAD("ImgHeight")) & ";" & vbCrLf
                        strADJS = strADJS & "objAD.FlashWmode     = " & Abs(PE_CLng(rsAD("FlashWmode"))) & ";" & vbCrLf
                        If PE_CLng(rsAD("ADType")) = 3 Then
                            strADJS = strADJS & "objAD.ADIntro        = """ & Html2Js(PE_HTMLEncode(rsAD("ADIntro"))) & """;" & vbCrLf
                        Else
                            strADJS = strADJS & "objAD.ADIntro        = """ & Html2Js(rsAD("ADIntro")) & """;" & vbCrLf
                        End If
                        strADJS = strADJS & "objAD.LinkUrl        = """ & rsAD("LinkUrl") & """;" & vbCrLf
                        strADJS = strADJS & "objAD.LinkTarget     = " & Abs(PE_CLng(rsAD("LinkTarget"))) & ";" & vbCrLf
                        strADJS = strADJS & "objAD.LinkAlt        = """ & rsAD("LinkAlt") & """;" & vbCrLf
                        strADJS = strADJS & "objAD.Priority       = " & PE_CLng(rsAD("Priority")) & ";" & vbCrLf
                        strADJS = strADJS & "objAD.CountView      = " & Abs(PE_CLng(rsAD("CountView"))) & ";" & vbCrLf
                        strADJS = strADJS & "objAD.CountClick     = " & Abs(PE_CLng(rsAD("CountClick"))) & ";" & vbCrLf
                        
                        strADJS = strADJS & "objAD.ADDIR          = """ & ADDir & """;" & vbCrLf
                        strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".AddAD(objAD);" & vbCrLf
                        rsAD.MoveNext
                    Loop
                    strADJS = strADJS & vbCrLf
                    strADJS = strADJS & "ZoneAD_" & arrZoneID(j) & ".Show();" & vbCrLf
                End If
                rsAD.Close
                Set rsAD = Nothing
            End If
            If Not fso.FolderExists(Server.MapPath(ZoneJS_Path)) Then
                fso.CreateFolder Server.MapPath(ZoneJS_Path)
            End If
            Call WriteToFile(ZoneJS_Name, strADJS)
        End If
        rsZone.Close
        Set rsZone = Nothing
    Next
End Sub

Sub ZoneJSTemplate()
    Dim i
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td height='22'>�����ڵ�λ�ã���վ������&nbsp;&gt;&gt;&nbsp;���JSģ��</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td>"
    Response.Write "      <table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "        <tr class='title'>"
    Response.Write "          <td width='50' height='22' align='center'><strong>����ID</strong></td>"
    Response.Write "          <td width='150' height='22' align='center'><strong>ģ����������</strong></td>"
    Response.Write "          <td height='22' align='center'><strong>ģ���ļ�����·��</strong></td>"
    Response.Write "          <td width='130' height='22' align='center'><strong>ģ���ļ���С</strong></td>"
    Response.Write "          <td width='130' height='22' align='center'><strong>����</strong></td>"
    Response.Write "        </tr>"
    For i = 1 To ZoneTypeNum
        Response.Write "        <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
        Response.Write "          <td width='50' align='center'>" & i & "</td>"
        Response.Write "          <td width='150' align='center'>" & ZoneConfig(i, 1) & "</td>"
        Response.Write "          <td align='center'>" & GetTemplateName(i) & "</td>"
        Response.Write "          <td width='130' align='center'>"
        Set hf = fso.GetFile(Server.MapPath(GetTemplateName(i)))
        Response.Write Round(hf.Size / 1024, 1) & " KB"
        Response.Write "          </td>"
        Response.Write "          <td width='130' align='center'>"
        Response.Write "            <a href='Admin_Advertisement.asp?Action=ModifyTemplate&ZoneType=" & i & "'>�޸�ģ������</a>"
        Response.Write "          </td>"
        Response.Write "        </tr>"
    Next
    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub

Sub ModifyTemplate()
    Dim ZoneType
    ZoneType = Trim(Request("ZoneType"))
    If ZoneType = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ����λ����ID</li>"
        Exit Sub
    Else
        ZoneType = PE_CLng(ZoneType)
    End If

    Response.Write "<br>"
    Response.Write "<table width='100%' class='border' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "  <form name='myform' method='Post' action='Admin_Advertisement.asp'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' align='center'><strong>�޸�ģ������</strong>--" & ZoneConfig(ZoneType, 1) & "</td>"
    Response.Write "  </tr>"

    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='350' align='center'>"
    Response.Write "      <textarea name='TemplateContent' cols='110' rows='20' wrap='off' id='TemplateContent'>" & GetZoneJSTemplate(ZoneType) & "</textarea>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td height='50' align='center'>"
    Response.Write "      <input name='ZoneType' type='hidden' id='ZoneType' value='" & ZoneType & "'>"
    Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveTemplate'>"
    Response.Write "      <input type='submit' name='Submit2' value=' �����޸Ľ�� '>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  </form>"
    Response.Write "</table>"
End Sub

Sub SaveTemplate()
    If ObjInstalled_FSO = False Then
        Exit Sub
    End If

    Dim ZoneType, TemplateFile, TemplateContent
    ZoneType = Trim(Request("ZoneType"))
    TemplateContent = Request("TemplateContent")
    
    TemplateFile = GetTemplateName(ZoneType)
    Call WriteToFile(TemplateFile, TemplateContent)
    Call WriteSuccessMsg("����ģ�����óɹ���", ComeUrl)
End Sub



Sub ADList()
    Dim rsAD, sqlAD
    Dim strAD, strADLink
    Call ShowJS_Tooltip
    Call ShowJS_ADMain("���")
    strFileName = "Admin_Advertisement.asp?Action=ADList&ZoneID="&PE_Clng(Request("ZoneID"))
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td height='22'>" & GetADManagePath() & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_Advertisement.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "    <td>"
    Response.Write "      <table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "        <tr class='title' height='22'>"
    Response.Write "          <td width='30' align='center'><strong>ѡ��</strong></td>"
    Response.Write "          <td width='30' align='center'><strong>ID</strong></td>"
    Response.Write "          <td width='65' align='center'><strong>���Ԥ��</strong></td>"
    Response.Write "          <td align='center'><strong>�������</strong></td>"
    Response.Write "          <td width='60' align='center'><strong>�������</strong></td>"
    Response.Write "          <td width='40' align='center'><strong>Ȩ��</strong></td>"
    Response.Write "          <td width='50' align='center'><strong>�����</strong></td>"
    Response.Write "          <td width='50' align='center'><strong>�����</strong></td>"
    Response.Write "          <td width='50' align='center'><strong>�����</strong></td>"
    Response.Write "          <td width='40' align='center'><strong>�����</strong></td>"
    Response.Write "          <td width='150' align='center'><strong>����</strong></td>"
    Response.Write "        </tr>"

    sqlAD = "select * from PE_Advertisement where 1=1"
    If ZoneID <> "" Then
        Dim tZone
        Set tZone = Conn.Execute("select IncludeADID from PE_AdZone where ZoneID=" & PE_CLng(ZoneID))
        If Not (tZone.BOF And tZone.EOF) Then
            If Not IsNull(tZone(0)) And tZone(0) <> "" Then
                sqlAD = sqlAD & " and ADID in (" & tZone(0) & ") "
            Else
                sqlAD = sqlAD & " and 1=0 "
            End If
        End If
        Set tZone = Nothing
    End If
    If Keyword <> "" Then
        Select Case strField
        Case "ZoneName"
            sqlAD = sqlAD & " and ADName like '%" & Keyword & "%' "
        Case "ZoneIntro"
            sqlAD = sqlAD & " and ADIntro like '%" & Keyword & "%' "
        Case Else
            sqlAD = sqlAD & " and ADName like '%" & Keyword & "%' "
        End Select
    End If
    sqlAD = sqlAD & " order by ADID desc"

    Set rsAD = Server.CreateObject("ADODB.Recordset")
    rsAD.Open sqlAD, Conn, 1, 1
    If rsAD.BOF And rsAD.EOF Then
        Response.Write "        <tr class='tdbg'><td colspan='20' align='center'><br>û���κι�棡<br><br></td></tr>"
    Else
        totalPut = rsAD.RecordCount
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
                rsAD.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Dim ADNum
        ADNum = 0
        Do While Not rsAD.EOF
            Response.Write "        <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "          <td width='30' align='center'><input name='ADID' type='checkbox' onclick='CheckItem(this)' value='" & rsAD("ADID") & "'></td>"
            Response.Write "          <td width='30' align='center'>" & rsAD("ADID") & "</td>"
            Response.Write "          <td width='65' align='center'>"
            If rsAD("ADType") = 4 Then
                Response.Write "            <a onmouseover=""ShowADPreview('&nbsp;����������Ԥ��&nbsp;')"" onmouseout=""hideTooltip('dHTMLADPreview')"" href='Admin_Advertisement.asp?Action=PreviewAD&ADID=" & rsAD("ADID") & "'>Ԥ��</a>"
            Else
                Response.Write "            <a onmouseover=""ShowADPreview('" & FixJs(GetADContent(rsAD("ADID"))) & "')"" onmouseout=""hideTooltip('dHTMLADPreview')"" href='Admin_Advertisement.asp?Action=PreviewAD&ADID=" & rsAD("ADID") & "'>Ԥ��</a>"
            End If
            Response.Write "          </td>"
            Response.Write "          <td><a href='Admin_Advertisement.asp?Action=ModifyAD&ADID=" & rsAD("ADID") & "'>" & rsAD("ADName") & "</a></td>"
            Response.Write "          <td width='60' align='center'>" & GetADType(rsAD("ADType")) & "</td>"
            Response.Write "          <td width='40' align='center'>" & rsAD("Priority") & "</td>"
            Response.Write "          <td width='50' align='center'>"
            If rsAD("CountClick") = True Then
                Response.Write rsAD("Clicks")
            Else
                Response.Write "<font color='#999999'>��ͳ��</font>"
            End If
            Response.Write "          </td>"
            Response.Write "          <td width='50' align='center'>"
            If rsAD("CountView") = True Then
                Response.Write rsAD("Views")
            Else
                Response.Write "<font color='#999999'>��ͳ��</font>"
            End If
            Response.Write "          </td>"
            Response.Write "          <td width='50' align='center'>"
            If rsAD("Views") > 1 Then
                Response.Write Round(rsAD("Clicks") / rsAD("Views"), 3)
            End If
            Response.Write "</td>"
            Response.Write "          <td width='40' align='center'>"
            If rsAD("Passed") = True Then
                Response.Write "<b>��</b>"
            Else
                Response.Write "<font color=red><b>��</b></font>"
            End If
            Response.Write "          </td>"
            Response.Write "          <td width='150' align='center'>"
            Response.Write "            <a href='Admin_Advertisement.asp?Action=ModifyAD&ADID=" & rsAD("ADID") & "'>�޸�</a>"
            Response.Write "            <a href='Admin_Advertisement.asp?Action=CopyAD&ADID=" & rsAD("ADID") & "'>����</a>"
            Response.Write "            <a href='Admin_Advertisement.asp?Action=DelAD&ADID=" & rsAD("ADID") & "' onClick=""return confirm('ȷ��Ҫɾ���˹����');"">ɾ��</a>"
            If rsAD("Passed") = False Then
                Response.Write "            <a href='Admin_Advertisement.asp?Action=SetADPassed&ADID=" & rsAD("ADID") & "'>ͨ�����</a>"
            Else
                Response.Write "            <a href='Admin_Advertisement.asp?Action=CancelADPassed&ADID=" & rsAD("ADID") & "'>ȡ��ͨ��</a>"
            End If
            Response.Write "          </td>"
            Response.Write "        </tr>"

            ADNum = ADNum + 1
            If ADNum >= MaxPerPage Then Exit Do
            rsAD.MoveNext
        Loop
    End If
    rsAD.Close
    Set rsAD = Nothing
    Response.Write "      </table>"
    Response.Write "      <table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "        <tr>"
    Response.Write "          <td width='130' height='30'>"
    Response.Write "            <input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>ѡ�����еĹ��"
    Response.Write "          </td>"
    Response.Write "          <td>"
    Response.Write "            <input type='submit' value='ɾ��ѡ�����' name='submit' onClick=""document.myform.Action.value='DelAD'"">&nbsp;"
    Response.Write "            <input type='submit' value='���ͨ��ѡ�����' name='submit' onClick=""document.myform.Action.value='SetADPassed'"">&nbsp;"
    Response.Write "            <input type='submit' value='ȡ�����ѡ�����' name='submit' onClick=""document.myform.Action.value='CancelADPassed'"">&nbsp;"
    Response.Write "            <input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "    </td>"
    Response.Write "  </form>"
    Response.Write "  </tr>"
    Response.Write "</table>"

    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "�����", True)
    End If
    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='80' align='right'><strong>���������</strong></td>"
    Response.Write "   <td>" & GetADSearchForm(strFileName) & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"
End Sub

Function GetADSearchForm(Action)
    Dim strForm
    strForm = "<table border='0' cellpadding='0' cellspacing='0'>"
    strForm = strForm & "<form method='Get' name='SearchForm' action='" & Action & "'>"
    strForm = strForm & "<tr><td height='28' align='center'> "
    strForm = strForm & "<select name='Field' size='1'>"
    strForm = strForm & "<option value='ADName' selected>�������</option>"
    strForm = strForm & "<option value='ADIntro'>�����</option>"
    strForm = strForm & "</select> "
    strForm = strForm & "<input type='text' name='keyword'  size='20' value='�ؼ���' maxlength='50' onFocus='this.select();'> "
    strForm = strForm & "<input name='Action' type='hidden' id='Action' value='ADList'>"
    strForm = strForm & "<input type='submit' name='Submit'  value='����'>"
    strForm = strForm & "</td></tr></form></table>"
    GetADSearchForm = strForm
End Function

Sub ShowJS_AD()
    Response.Write "<script language=JavaScript>" & vbCrLf
    Response.Write "<!--" & vbCrLf
    Response.Write "function CheckForm()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  if(myform.ADName.value==''){" & vbCrLf
    Response.Write "    alert('������Ʋ���Ϊ�գ�');" & vbCrLf
    Response.Write "    myform.ADName.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(myform.ADType[0].checked == true && myform.ImgUrl.value==''){" & vbCrLf
    Response.Write "    alert('ͼƬ��ַ����Ϊ�գ�');" & vbCrLf
    Response.Write "    myform.ImgUrl.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(myform.ADType[1].checked == true && myform.FlashUrl.value==''){" & vbCrLf
    Response.Write "    alert('������ַ����Ϊ�գ�');" & vbCrLf
    Response.Write "    myform.FlashUrl.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(myform.ADType[2].checked == true && myform.ADText.value==''){" & vbCrLf
    Response.Write "    alert('����ı�����Ϊ�գ�');" & vbCrLf
    Response.Write "    myform.ADText.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(myform.ADType[3].checked == true && myform.ADCode.value==''){" & vbCrLf
    Response.Write "    alert('�����벻��Ϊ�գ�');" & vbCrLf
    Response.Write "    myform.ADCode.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if(myform.Priority.value==''){" & vbCrLf
    Response.Write "    alert('���Ȩ�ز���Ϊ�գ�');" & vbCrLf
    Response.Write "    myform.Priority.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "function ADTypeChecked(i)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  document.myform.ADType[i].checked = true;" & vbCrLf
    Response.Write "  Change_ADType();" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "function Change_ADType()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  for (var j=0;j<document.myform.ADType.length;j++){" & vbCrLf
    Response.Write "    var ot = eval('document.all.ADContent_' + (j + 1) + '');" & vbCrLf
    Response.Write "    if(document.myform.ADType[j].checked==true){" & vbCrLf
    Response.Write "      ot.style.display = '';" & vbCrLf
    Response.Write "      if(j==0){" & vbCrLf
    Response.Write "        document.myform.CountClick.disabled = false;" & vbCrLf
    Response.Write "        document.myform.Clicks.disabled = false;" & vbCrLf
    Response.Write "      }" & vbCrLf
    Response.Write "      else{" & vbCrLf
    Response.Write "        document.myform.CountClick.disabled = true;" & vbCrLf
    Response.Write "        document.myform.Clicks.disabled = true;" & vbCrLf
    Response.Write "      }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    else{" & vbCrLf
    Response.Write "      ot.style.display = 'none';" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

'=================================================
'��������Export
'��  �ã�������λ
'=================================================
Sub Export()

    Dim rs, sql, iCount

    sql = "select * from PE_AdZone"
    Set rs = Conn.Execute(sql)
 
    Response.Write "<form name='myform' method='post' action='Admin_Advertisement.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>������λ</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'><td height='10'></td></tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='center'>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr>"
    Response.Write "           <td>"
    Response.Write "            <select name='ZoneID' size='2' multiple style='height:300px;width:450px;'>"
    
    If rs.BOF And rs.EOF Then
        Response.Write "         <option value=''>��û�а�λ��</option>"
        '�ر��ύ��ť
        iCount = 0
    Else
        iCount = rs.RecordCount
        Do While Not rs.EOF
            Response.Write "     <option value='" & rs("ZoneID") & "'>" & rs("ZoneName") & "</option>"
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
    Response.Write "        <td colspan='2'>Ŀ�����ݿ⣺<input name='AdZoneMdb' type='text' id='AdZoneMdb' value='../Temp/AdZone.mdb' size='20' maxlength='50'>&nbsp;&nbsp;<INPUT TYPE='checkbox' NAME='FormatConn' value='yes' id='id' checked> �����Ŀ�����ݿ�</td>"
    Response.Write "      </tr>"
    Response.Write "      <tr height='50'>"
    Response.Write "         <td colspan='2' align='center'><input type='submit' name='Submit' value='ִ�е�������' onClick=""document.myform.Action.value='DoExport';"">"
    Response.Write "                  <input name='Action' type='hidden' id='Action' value='Export'>"
    Response.Write "         </td>"
    Response.Write "        </tr>"
    Response.Write "    </table>"
    Response.Write "   </td>"
    Response.Write " </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.ZoneID.length;i++){" & vbCrLf
    Response.Write "    document.myform.ZoneID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.ZoneID.length;i++){" & vbCrLf
    Response.Write "    document.myform.ZoneID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub
'=================================================
'��������Import
'��  �ã������λ��һ��
'=================================================
Sub Import()
    Response.Write "<form name='myform' method='post' action='Admin_Advertisement.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center'><strong>����루��һ����</strong></td>"
    Response.Write "    </tr>"
    Response.Write " <tr class='tdbg'>"
    Response.Write "      <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;������Ҫ����İ�λ���ݿ���ļ�����"
    Response.Write "        <input name='AdZoneMdb' type='text' id='AdZoneMdb' value='../Temp/AdZone.mdb' size='20' maxlength='50'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' ��һ�� '>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='Import2'> </td>"
    Response.Write " </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub
'=================================================
'��������Import2
'��  �ã������λ�ڶ���
'=================================================
Sub Import2()
    Dim rs, sql
    Dim mdbname, tconn, trs, iCount
    mdbname = Replace(Trim(Request.Form("AdZonemdb")), "'", "")
    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д�����λ���ݿ���"
        Exit Sub
    End If
    
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
    If Err.Number <> 0 Then
        ErrMsg = ErrMsg & "<li>���ݿ����ʧ�ܣ����Ժ����ԣ�����ԭ��" & Err.Description
        Err.Clear
        Exit Sub
    End If

    Response.Write "<form name='myform' method='post' action='Admin_Advertisement.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' align='center'><strong>�����λ���ڶ�����</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100' align='center'>"
    Response.Write "        <br>"
    Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "          <tr align='center'>"
    Response.Write "            <td><strong>��������İ�λ</strong><br>"
    Response.Write "<select name='ZoneID' size='2' multiple style='height:300px;width:250px;'>"

    sql = "select * from PE_AdZone"
    Set rs = Server.CreateObject("Adodb.RecordSet")
    rs.Open sql, tconn, 1, 1
    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>û���κΰ�λ</option>"
        iCount = 0
    Else
        iCount = rs.RecordCount
        Do While Not rs.EOF
            Response.Write "<option value='" & rs("ZoneID") & "'>" & rs("ZoneName") & "</option>"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Response.Write "</select></td>"
    Response.Write "            <td width='80'><input type='submit' name='Submit' value='����&gt;&gt;' "
    If iCount = 0 Then Response.Write " disabled"
    Response.Write "></td>"
    Response.Write "            <td><strong>ϵͳ���Ѿ����ڵİ�λ</strong><br>"
    Response.Write "             <select name='tZoneID' size='2' multiple style='height:300px;width:250px;' disabled>"

    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
        Response.Write "<option value='0'>û���κ�ģ��</option>"
    Else
        Do While Not rs.EOF
            Response.Write "<option value='" & rs("ZoneID") & "'>" & rs("ZoneName") & "</option>"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing

    Response.Write "              </select></td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "     <br><b>��ʾ����ס��Ctrl����Shift�������Զ�ѡ</b><br>"
    Response.Write "        <input name='AdZoneMdb' type='hidden' id='AdZoneMdb' value='" & mdbname & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='DoImport'>"
    Response.Write "        <br>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

'=================================================
'��������DoExport
'��  �ã�������λ����
'=================================================
Sub DoExport()
    On Error Resume Next
    Dim rs, rsMax
    Dim mdbname, tconn, trs
    Dim ZoneID, FormatConn
    FormatConn = Request.Form("FormatConn")
    ZoneID = Trim(Request("ZoneID"))
    mdbname = Replace(Trim(Request.Form("AdZonemdb")), "'", "")
    
    If IsValidID(ZoneID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����İ�λ</li>"
    End If
    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д������λ���ݿ���"
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If
    
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݿ����ʧ�ܣ����Ժ����ԣ�����ԭ��" & Err.Description
        Err.Clear
        Exit Sub
    End If

    If FoundErr = True Then
        Exit Sub
    End If

    If FormatConn <> "" Then
        tconn.Execute ("delete from PE_AdZone")
    End If
    Set rs = Conn.Execute("select * from PE_AdZone where ZoneID in (" & ZoneID & ")  order by ZoneID ")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_AdZone", tconn, 1, 3
    Do While Not rs.EOF
        Set rsMax = tconn.Execute("select max(ZoneId) from PE_AdZone")
        trs.addnew
        If IsNull(rsMax(0)) Then
            trs("ZoneID") = 1
        Else
            trs("ZoneID") = rsMax(0) + 1
        End If
        trs("ChannelID") = rs("ChannelID")
        trs("ZoneName") = rs("ZoneName")
        trs("ZoneJSName") = rs("ZoneJSName")
        trs("ZoneIntro") = rs("ZoneIntro")
        trs("ZoneType") = rs("ZoneType")
        trs("DefaultSetting") = rs("DefaultSetting")
        trs("ZoneSetting") = rs("ZoneSetting")
        trs("ZoneWidth") = rs("ZoneWidth")
        trs("ZoneHeight") = rs("ZoneHeight")
        trs("ShowType") = rs("ShowType")
        trs("Active") = rs("Active")
        trs("UpdateTime") = rs("UpdateTime")
        trs.Update
        rs.MoveNext
    Loop
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    rsMax.Close
    Set rsMax = Nothing
    tconn.Close
    Set tconn = Nothing
    Call WriteSuccessMsg("�Ѿ��ɹ�����ѡ�еİ�λ���õ�����ָ�������ݿ��У�", ComeUrl)
End Sub

'=================================================
'��������DoImport
'��  �ã������λ����
'=================================================
Sub DoImport()
    On Error Resume Next
    Dim mdbname, tconn, trs, rsMax
    Dim ZoneID
    Dim rs
    ZoneID = Trim(Request("ZoneID"))
    mdbname = Replace(Trim(Request.Form("AdZonemdb")), "'", "")
    If IsValidID(ZoneID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ����İ�λ</li>"
    End If
    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д�����λ���ݿ���"
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If
    
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
    If Err.Number <> 0 Then
        ErrMsg = ErrMsg & "<li>���ݿ����ʧ�ܣ����Ժ����ԣ�����ԭ��" & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    Set rs = tconn.Execute(" select * from PE_AdZone where ZoneID in (" & ZoneID & ")  order by ZoneID")
    Set trs = Server.CreateObject("adodb.recordset")
    trs.Open "select * from PE_AdZone", Conn, 1, 3
    Do While Not rs.EOF
        Set rsMax = Conn.Execute("select max(ZoneId) from PE_AdZone")
        trs.addnew
        If IsNull(rsMax(0)) Then
            trs("ZoneID") = 1
        Else
            trs("ZoneID") = rsMax(0) + 1
        End If
        trs("ChannelID") = rs("ChannelID")
        trs("ZoneName") = rs("ZoneName")
        trs("ZoneJSName") = rs("ZoneJSName")
        trs("ZoneIntro") = rs("ZoneIntro")
        trs("ZoneType") = rs("ZoneType")
        trs("DefaultSetting") = rs("DefaultSetting")
        trs("ZoneSetting") = rs("ZoneSetting")
        trs("ZoneWidth") = rs("ZoneWidth")
        trs("ZoneHeight") = rs("ZoneHeight")
        trs("ShowType") = rs("ShowType")
        trs("Active") = rs("Active")
        trs("UpdateTime") = rs("UpdateTime")
        trs.Update
        rs.MoveNext
    Loop
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
    tconn.Close
    Set tconn = Nothing
    Call WriteSuccessMsg("�Ѿ��ɹ���ָ�������ݿ��е���ѡ�еİ�λ��", ComeUrl)
    Call CreateJSZoneID(ZoneID)
    Call ClearSiteCache(0)
End Sub

Sub AddAD()
    Call ShowJS_AD
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Advertisement.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>�� �� �� ��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='center'>"
    Response.Write "      <td class='tdbg' valign='top' width='255'>"
    Response.Write "        <table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <tr>"
    Response.Write "            <td align='center'><b>������λ</b></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <select name='ZoneID' size='2' multiple style='height:360px;width:250px;'>"
    Response.Write GetZone_Option(ZoneID)
    Response.Write "              </select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "      <td valign='top'>"
    Response.Write "        <table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='70' align='right'><strong>������ƣ�</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='ADName' type='text' id='ADName' size='58' maxlength='255' value=''>"
    Response.Write "              <font color='red'>*</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='70' align='right'><strong>������ͣ�</strong></td>"
    Response.Write "            <td>"
    Response.Write "              " & GetADType_Option(1)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='70' align='right'><strong>������ݣ�</strong></td>"
    Response.Write "            <td height='265' valign='top'>"
    Response.Write "              <table id='ADContent_1' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff' style='display:'>"
    Response.Write "                <tr align='center' class='tdbg2'>"
    Response.Write "                  <td colspan='2'><strong>�����������--ͼƬ</strong></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>ͼƬ��ַ��</td>"
    Response.Write "                  <td>"
    Response.Write "                    <input name='ImgUrl' type='text' id='ImgUrl' size='58' maxlength='255' value=''>"
    Response.Write "                    <font color='red'>*</font>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>ͼƬ�ϴ���</td>"
    Response.Write "                  <td> <iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=AdPic' frameborder=0 scrolling=no width='360' height='25'></iframe>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>ͼƬ�ߴ磺</td>"
    Response.Write "                  <td>"
    Response.Write "                    ��<input name='ImgWidth' type='text' id='ImgWidth' size='6' maxlength='5' value=''>"
    Response.Write "                    ����&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "                    �ߣ�<input name='ImgHeight' type='text' id='ImgHeight' size='6' maxlength='5' value=''>"
    Response.Write "                    ����"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>���ӵ�ַ��</td>"
    Response.Write "                  <td>"
    Response.Write "                    <input name='LinkUrl' type='text' id='LinkUrl' value='http://' size='58' maxlength='255'>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>������ʾ��</td>"
    Response.Write "                  <td>"
    Response.Write "                    <input name='LinkAlt' type='text' id='LinkAlt' value='' size='58' maxlength='255'>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>����Ŀ�꣺</td>"
    Response.Write "                  <td>"
    Response.Write "                    <input name='LinkTarget' type='radio' id='LinkTarget' value='1' checked>�´���"
    Response.Write "                    <input name='LinkTarget' type='radio' id='LinkTarget' value='0'>ԭ����"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>����飺</td>"
    Response.Write "                  <td>"
    Response.Write "                    <textarea name='ADIntro' cols='48' rows='4' id='ADIntro'></textarea>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "              <table id='ADContent_2' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff' style='display:none'>"
    Response.Write "                <tr align='center' class='tdbg2'>"
    Response.Write "                  <td colspan='2'><strong>�����������--����</strong></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>������ַ��</td>"
    Response.Write "                  <td>"
    Response.Write "                    <input name='FlashUrl' type='text' id='FlashUrl' size='58' maxlength='255' value=''>"
    Response.Write "                    <font color='red'>*</font>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>�����ϴ���</td>"
    Response.Write "                  <td> <iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=AdPic' frameborder=0 scrolling=no width='360' height='25'></iframe>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>�����ߴ磺</td>"
    Response.Write "                  <td>"
    Response.Write "                    ��<input name='FlashWidth' type='text' id='FlashWidth' size='6' maxlength='5' value=''>"
    Response.Write "                    ����&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "                    �ߣ�<input name='FlashHeight' type='text' id='FlashHeight' size='6' maxlength='5' value=''>"
    Response.Write "                    ����"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>����͸����</td>"
    Response.Write "                  <td>"
    Response.Write "                    <input type='radio' name='FlashWmode' value='0' checked> ��͸��&nbsp;&nbsp;"
    Response.Write "                    <input type='radio' name='FlashWmode' value='1'> ͸��&nbsp;&nbsp;"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "              <table id='ADContent_3' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff' style='display:none'>"
    Response.Write "                <tr align='center' class='tdbg2'>"
    Response.Write "                  <td colspan='2'><strong>�����������--�ı�</strong></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td colspan='2' align='center'>"
    Response.Write "                    <textarea name='ADText' cols='64' rows='15' id='ADText'></textarea>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "              <table id='ADContent_4' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff' style='display:none'>"
    Response.Write "                <tr align='center' class='tdbg2'>"
    Response.Write "                  <td colspan='2'><strong>�����������--����</strong></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td colspan='2' align='center'>"
    Response.Write "                    <textarea name='ADCode' cols='64' rows='15' id='ADCode'></textarea>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "              <table id='ADContent_5' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff' style='display:none'>"
    Response.Write "                <tr align='center' class='tdbg2'>"
    Response.Write "                  <td colspan='2'><strong>�����������--ҳ��</strong></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>ҳ���ַ��</td>"
    Response.Write "                  <td>"
    Response.Write "                    <textarea name='WebFileUrl' cols='48' rows='4' id='WebFileUrl'></textarea>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='70' align='right'><strong>���Ȩ�أ�</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Priority' type='text' id='Priority' size='4' maxlength='3' value='1'> <font color='red'>*</font> ����Ϊ��λ��������ʾʱ������Ȩ��Ȩ��Խ����ʾ����Խ��"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='70' align='right'><strong>���ͳ�ƣ�</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='CountView' type='checkbox' id='CountView' value='yes'> ͳ�������  �������<input name='Views' type='text' id='Views' size='5' maxlength='6' value=''>"
    Response.Write "              &nbsp;&nbsp;&nbsp;&nbsp;<input name='CountClick' type='checkbox' id='CountClick' value='yes'> ͳ�Ƶ����  �������<input name='Clicks' type='text' id='Clicks' size='5' maxlength='6' value=''>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='70' align='right'><strong>���״̬��</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Passed' type='checkbox' id='Passed' value='yes' checked> ͨ�����"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"
    Response.Write "    <tr>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveAddAD'>"
    Response.Write "        <input type='submit' name='Submit' value=' �� �� '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub ModifyAD()
    Dim ADID, rsAD, sqlAD
    ADID = Trim(Request("ADID"))
    If ADID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ҫ�޸ĵĹ��ID</li>"
        Exit Sub
    Else
        ADID = PE_CLng(ADID)
    End If
    sqlAD = "select * from PE_Advertisement where ADID=" & ADID
    Set rsAD = Conn.Execute(sqlAD)
    If rsAD.BOF And rsAD.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���Ĺ��</li>"
        rsAD.Close
        Set rsAD = Nothing
        Exit Sub
    End If

    Dim ADType
    Dim ImgUrl, ImgWidth, ImgHeight, LinkUrl, LinkAlt, LinkTarget, ADIntro
    Dim FlashUrl, FlashWidth, FlashHeight, FlashWmode
    Dim ADText, ADCode, WebFileUrl, strDisabled

    ADType = rsAD("ADType")
    Select Case ADType
    Case 1
        ImgUrl = rsAD("ImgUrl")
        ImgWidth = rsAD("ImgWidth")
        ImgHeight = rsAD("ImgHeight")
        LinkUrl = rsAD("LinkUrl")
        LinkAlt = rsAD("LinkAlt")
        LinkTarget = rsAD("LinkTarget")
        ADIntro = rsAD("ADIntro")
    Case 2
        FlashUrl = rsAD("ImgUrl")
        FlashWidth = rsAD("ImgWidth")
        FlashHeight = rsAD("ImgHeight")
        FlashWmode = rsAD("FlashWmode")
    Case 3
        ADText = rsAD("ADIntro")
    Case 4
        ADCode = rsAD("ADIntro")
    Case 5
        WebFileUrl = rsAD("ADIntro")
    End Select
    If ADType <> 1 Then strDisabled = " disabled"

    Call ShowJS_AD
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Advertisement.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>�� �� �� ��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='center'>"
    Response.Write "      <td class='tdbg' valign='top' width='255'>"
    Response.Write "        <table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <tr>"
    Response.Write "            <td align='center'><b>������λ</b></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr>"
    Response.Write "            <td>"
    Response.Write "              <select name='ZoneID' size='2' multiple style='height:360px;width:250px;'>"
    Response.Write GetZone_Option(rsAD("ZoneID"))
    Response.Write "              </select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "      <td valign='top'>"
    Response.Write "        <table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='70' align='right'><strong>������ƣ�</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='ADName' type='text' id='ADName' size='58' maxlength='255' value='" & rsAD("ADName") & "'>"
    Response.Write "              <font color='red'>*</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='70' align='right'><strong>������ͣ�</strong></td>"
    Response.Write "            <td>"
    Response.Write "              " & GetADType_Option(rsAD("ADType"))
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='70' align='right'><strong>������ݣ�</strong></td>"
    Response.Write "            <td height='265' valign='top'>"
    Response.Write "              <table id='ADContent_1' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff' style='display:" & StyleDisplay(rsAD("ADType"), 1) & "'>"
    Response.Write "                <tr align='center' class='tdbg2'>"
    Response.Write "                  <td colspan='2'><strong>�����������--ͼƬ</strong></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>ͼƬ��ַ��</td>"
    Response.Write "                  <td>"
    Response.Write "                    <input name='ImgUrl' type='text' id='ImgUrl' size='58' maxlength='255' value='" & ImgUrl & "'>"
    Response.Write "                    <font color='red'>*</font>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>ͼƬ�ϴ���</td>"
    Response.Write "                  <td> <iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=AdPic' frameborder=0 scrolling=no width='360' height='25'></iframe>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>ͼƬ�ߴ磺</td>"
    Response.Write "                  <td>"
    Response.Write "                    ��<input name='ImgWidth' type='text' id='ImgWidth' size='6' maxlength='5' value='" & ImgWidth & "'>"
    Response.Write "                    ����&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "                    �ߣ�<input name='ImgHeight' type='text' id='ImgHeight' size='6' maxlength='5' value='" & ImgHeight & "'>"
    Response.Write "                    ����"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>���ӵ�ַ��</td>"
    Response.Write "                  <td>"
    Response.Write "                    <input name='LinkUrl' type='text' id='LinkUrl' value='" & LinkUrl & "' size='58' maxlength='255'>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>������ʾ��</td>"
    Response.Write "                  <td>"
    Response.Write "                    <input name='LinkAlt' type='text' id='LinkAlt' value='" & LinkAlt & "' size='58' maxlength='255'>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>����Ŀ�꣺</td>"
    Response.Write "                  <td>"
    Response.Write "                    <input name='LinkTarget' type='radio' id='LinkTarget' value='1' " & IsRadioChecked(LinkTarget, 1) & ">�´���"
    Response.Write "                    <input name='LinkTarget' type='radio' id='LinkTarget' value='0' " & IsRadioChecked(LinkTarget, 0) & ">ԭ����"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>����飺</td>"
    Response.Write "                  <td>"
    Response.Write "                    <textarea name='ADIntro' cols='48' rows='4' id='ADIntro'>" & PE_ConvertBR(ADIntro) & "</textarea>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "              <table id='ADContent_2' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff' style='display:" & StyleDisplay(rsAD("ADType"), 2) & "'>"
    Response.Write "                <tr align='center' class='tdbg2'>"
    Response.Write "                  <td colspan='2'><strong>�����������--����</strong></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>������ַ��</td>"
    Response.Write "                  <td>"
    Response.Write "                    <input name='FlashUrl' type='text' id='FlashUrl' size='58' maxlength='255' value='" & FlashUrl & "'>"
    Response.Write "                    <font color='red'>*</font>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>�����ϴ���</td>"
    Response.Write "                  <td> <iframe style='top:2px' ID='uploadPhoto' src='Upload.asp?dialogtype=AdPic' frameborder=0 scrolling=no width='360' height='25'></iframe>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>�����ߴ磺</td>"
    Response.Write "                  <td>"
    Response.Write "                    ��<input name='FlashWidth' type='text' id='FlashWidth' size='6' maxlength='5' value='" & FlashWidth & "'>"
    Response.Write "                    ����&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "                    �ߣ�<input name='FlashHeight' type='text' id='FlashHeight' size='6' maxlength='5' value='" & FlashHeight & "'>"
    Response.Write "                    ����"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>����͸����</td>"
    Response.Write "                  <td>"
    Response.Write "                    <input type='radio' name='FlashWmode' value='0' " & IsRadioChecked(FlashWmode, 0) & "> ��͸��&nbsp;&nbsp;"
    Response.Write "                    <input type='radio' name='FlashWmode' value='1' " & IsRadioChecked(FlashWmode, 1) & "> ͸��&nbsp;&nbsp;"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "              <table id='ADContent_3' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff' style='display:" & StyleDisplay(rsAD("ADType"), 3) & "'>"
    Response.Write "                <tr align='center' class='tdbg2'>"
    Response.Write "                  <td colspan='2'><strong>�����������--�ı�</strong></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td colspan='2' align='center'>"
    Response.Write "                    <textarea name='ADText' cols='64' rows='15' id='ADText'>" & ADText & "</textarea>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "              <table id='ADContent_4' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff' style='display:" & StyleDisplay(rsAD("ADType"), 4) & "'>"
    Response.Write "                <tr align='center' class='tdbg2'>"
    Response.Write "                  <td colspan='2'><strong>�����������--����</strong></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td colspan='2' align='center'>"
    Response.Write "                    <textarea name='ADCode' cols='64' rows='15' id='ADCode'>" & ADCode & "</textarea>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "              <table id='ADContent_5' width='100%' border='0' cellpadding='2' cellspacing='1' bgcolor='#ffffff' style='display:" & StyleDisplay(rsAD("ADType"), 5) & "'>"
    Response.Write "                <tr align='center' class='tdbg2'>"
    Response.Write "                  <td colspan='2'><strong>�����������--ҳ��</strong></td>"
    Response.Write "                </tr>"
    Response.Write "                <tr class='tdbg'>"
    Response.Write "                  <td width='80' align='right'>ҳ���ַ��</td>"
    Response.Write "                  <td>"
    Response.Write "                    <textarea name='WebFileUrl' cols='48' rows='4' id='WebFileUrl'>" & PE_ConvertBR(WebFileUrl) & "</textarea>"
    Response.Write "                  </td>"
    Response.Write "                </tr>"
    Response.Write "              </table>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='70' align='right'><strong>���Ȩ�أ�</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Priority' type='text' id='Priority' size='4' maxlength='3' value='" & rsAD("Priority") & "'> <font color='red'>*</font> ����Ϊ��λ��������ʾʱ������Ȩ��Ȩ��Խ����ʾ����Խ��"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='70' align='right'><strong>���ͳ�ƣ�</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='CountView' type='checkbox' id='CountView' value='yes' " & IsRadioChecked(rsAD("CountView"), True) & "> ͳ�������  �������<input name='Views' type='text' id='Views' size='5' maxlength='6' value='" & rsAD("Views") & "'>"
    Response.Write "              &nbsp;&nbsp;&nbsp;&nbsp;<input name='CountClick' type='checkbox' id='CountClick' value='yes' " & IsRadioChecked(rsAD("CountClick"), True) & " " & strDisabled & "> ͳ�Ƶ����  �������<input name='Clicks' type='text' id='Clicks' size='5' maxlength='6' value='" & rsAD("Clicks") & "' " & strDisabled & ">"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='70' align='right'><strong>���״̬��</strong></td>"
    Response.Write "            <td>"
    Response.Write "              <input name='Passed' type='checkbox' id='Passed' value='yes' " & IsRadioChecked(rsAD("Passed"), True) & "> ͨ�����"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"
    Response.Write "    <tr>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='ADID' type='hidden' id='ADID' value='" & rsAD("ADID") & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModifyAD'>"
    Response.Write "        <input type='submit' name='Submit' value=' �� �� '>&nbsp;&nbsp;"
    Response.Write "        <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Advertisement.asp?Action=ADList'"" style='cursor:hand;'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
    rsAD.Close
    Set rsAD = Nothing
End Sub

Sub SaveAD()
    Dim ADID, ADType, ADName, ZoneID, Priority, Passed, ADSetting, CountView, Views, CountClick, Clicks
    Dim ImgUrl, ImgWidth, ImgHeight, FlashWmode, LinkUrl, LinkAlt, LinkTarget, ADIntro
    Dim rsAD, sqlAD, OldZoneID

    ADID = PE_CLng(Trim(Request("ADID")))
    ADType = PE_CLng(Trim(Request("ADType")))
    ADName = Trim(Request("ADName"))
    ZoneID = Trim(Request("ZoneID"))
    Priority = PE_CLng(Trim(Request("Priority")))
    Passed = Trim(Request("Passed"))
    CountView = Trim(Request("CountView"))
    Views = PE_CLng(Trim(Request("Views")))
    CountClick = Trim(Request("CountClick"))
    Clicks = PE_CLng(Trim(Request("Clicks")))

    ImgUrl = ""
    ImgWidth = 0
    ImgHeight = 0
    FlashWmode = 0
    LinkUrl = ""
    LinkAlt = ""
    LinkTarget = 1
    ADIntro = ""
    Select Case ADType
    Case 1
        ImgUrl = Trim(Request("ImgUrl"))
        ImgWidth = PE_CLng(Trim(Request("ImgWidth")))
        ImgHeight = PE_CLng(Trim(Request("ImgHeight")))
        LinkUrl = Trim(Request("LinkUrl"))
        If LinkUrl = "http://" Then LinkUrl = ""
        LinkAlt = Trim(Request("LinkAlt"))
        LinkTarget = PE_CLng(Trim(Request("LinkTarget")))
        ADIntro = Trim(Request("ADIntro"))
    Case 2
        ImgUrl = Trim(Request("FlashUrl"))
        ImgWidth = PE_CLng(Trim(Request("FlashWidth")))
        ImgHeight = PE_CLng(Trim(Request("FlashHeight")))
        FlashWmode = PE_CLng(Trim(Request("FlashWmode")))
    Case 3
        ADIntro = Trim(Request("ADText"))
    Case 4
        ADIntro = Trim(Request("ADCode"))
    Case 5
        ADIntro = Trim(Request("WebFileUrl"))
    End Select

    If ADName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������Ʋ���Ϊ�գ�</li>"
    End If
    If Priority = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���Ȩ�ر�����д��</li>"
    End If
    Select Case ADType
    Case 1
        If ImgUrl = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ͼƬ��ַ����Ϊ�գ�</li>"
        End If
    Case 2
        If ImgUrl = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>������ַ����Ϊ�գ�</li>"
        End If
    Case 3
        If ADIntro = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����ı�����Ϊ�գ�</li>"
        End If
    Case 4
        If ADIntro = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�����벻��Ϊ�գ�</li>"
        End If
    Case 5
        If ADIntro = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>ҳ���ַ����Ϊ�գ�</li>"
        End If
    End Select
    If FoundErr = True Then
        Exit Sub
    End If
    
    ADName = PE_HTMLEncode(ADName)
    ImgUrl = PE_HTMLEncode(ImgUrl)
    LinkUrl = PE_HTMLEncode(LinkUrl)
    LinkAlt = PE_HTMLEncode(LinkAlt)
    If ADType <> 3 And ADType <> 4 Then
        ADIntro = PE_HTMLEncode(ADIntro)
    End If
    If IsValidID(ZoneID) = False Then
        ZoneID = ""
    Else
        ZoneID = ReplaceBadChar(ZoneID)
    End If
    CountView = CBool(CountView = "yes")
    CountClick = CBool(CountClick = "yes")
    Passed = CBool(Passed = "yes")

    Set rsAD = Server.CreateObject("adodb.recordset")
    If Action = "SaveAddAD" Then
        sqlAD = "select top 1 * from PE_Advertisement"
        rsAD.Open sqlAD, Conn, 1, 3
        rsAD.addnew
        ADID = PE_CLng(Conn.Execute("select max(ADID) from PE_Advertisement")(0)) + 1
        rsAD("ADID") = ADID
        Call AddADID_Zone(ZoneID, ADID)
    ElseIf Action = "SaveModifyAD" Then
        If ADID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>����ȷ�����ID��ֵ</li>"
            Exit Sub
        End If
        sqlAD = "select * from PE_Advertisement where ADID=" & ADID
        rsAD.Open sqlAD, Conn, 1, 3
        If rsAD.BOF And rsAD.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�Ҳ���ָ���Ĺ�棡</li>"
            rsAD.Close
            Set rsAD = Nothing
            Exit Sub
        End If
        OldZoneID = rsAD("ZoneID")
    End If
    rsAD("ADType") = ADType
    rsAD("ADName") = ADName
    rsAD("ZoneID") = ZoneID
    rsAD("ImgUrl") = ImgUrl
    rsAD("ImgWidth") = ImgWidth
    rsAD("ImgHeight") = ImgHeight
    rsAD("FlashWmode") = FlashWmode
    rsAD("LinkUrl") = LinkUrl
    rsAD("LinkAlt") = LinkAlt
    rsAD("LinkTarget") = LinkTarget
    rsAD("ADIntro") = ADIntro
    rsAD("Priority") = Priority
    rsAD("CountView") = CountView
    rsAD("Views") = Views
    rsAD("CountClick") = CountClick
    rsAD("Clicks") = Clicks
    rsAD("Passed") = Passed
    rsAD.Update
    rsAD.Close
    Set rsAD = Nothing

    If ZoneID <> OldZoneID Then
        Call DelADID_Zone(OldZoneID, ADID)
        Call AddADID_Zone(ZoneID, ADID)
        Call CreateJSZoneID(OldZoneID)
    End If
    Call CreateJSZoneID(ZoneID)
    
    Call WriteEntry(2, AdminName, "��ӹ��ɹ���" & ADName)

    Call ClearSiteCache(0)
    Call CloseConn
    Response.Redirect "Admin_Advertisement.asp?Action=ADList"
End Sub

Sub CopyAD()
    Dim MaxADID
    ADID = PE_CLng(Trim(Request("ADID")))
    If ADID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������ʧ��</li>"
        Exit Sub
    End If
    
    MaxADID = PE_CLng(Conn.Execute("select max(ADID) from PE_Advertisement")(0)) + 1
    Conn.Execute ("insert into PE_Advertisement select " & MaxADID & " as ADID,UserID,ZoneID,ADType, '���� '+ADName as ADName,ImgUrl,ImgWidth,ImgHeight,FlashWmode,ADIntro,LinkUrl,LinkTarget,LinkAlt,Priority,Setting,CountView,Views,CountClick,Clicks,Passed from PE_Advertisement where ADID=" & ADID)
    Call WriteEntry(2, AdminName, "���ƹ��ɹ������ID��" & ADID)
    Response.Redirect ComeUrl
End Sub

Sub AddADID_Zone(arrZoneID, iADID)
    If iADID = "" Or IsNull(iADID) Then
        Exit Sub
    Else
        iADID = PE_CLng(iADID)
    End If
    If IsValidID(arrZoneID) = True Then
        Dim sqlZone, rsZone
        arrZoneID = ReplaceBadChar(arrZoneID)
        sqlZone = "select IncludeADID from PE_AdZone where ZoneID in (" & arrZoneID & ")"
        Set rsZone = Server.CreateObject("Adodb.RecordSet")
        rsZone.Open sqlZone, Conn, 1, 3
        Do While Not rsZone.EOF
            rsZone(0) = AppendStr(rsZone(0), CStr(iADID), ",")
            rsZone.Update
            rsZone.MoveNext
        Loop
        rsZone.Close
        Set rsZone = Nothing
    End If
End Sub

Sub DelADID_Zone(arrZoneID, iADID)
    If iADID = "" Or IsNull(iADID) Then
        Exit Sub
    Else
        iADID = PE_CLng(iADID)
    End If
    If IsValidID(arrZoneID) = True Then
        Dim sqlZone, rsZone
        sqlZone = "select IncludeADID from PE_AdZone where ZoneID in (" & arrZoneID & ")"
        Set rsZone = Server.CreateObject("Adodb.RecordSet")
        rsZone.Open sqlZone, Conn, 1, 3
        Do While Not rsZone.EOF
            rsZone(0) = RemoveStr(rsZone(0), CStr(iADID), ",")
            rsZone.Update
            rsZone.MoveNext
        Loop
        rsZone.Close
        Set rsZone = Nothing
    End If
End Sub

Sub SetADProperty()
    Dim ADID, sqlProperty, rsProperty
    Dim MoveChannelID
    ADID = Trim(Request("ADID"))
    If IsValidID(ADID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ�����ID</li>"
        Exit Sub
    End If
    If Action = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�������㣡</li>"
        Exit Sub
    End If

    If InStr(ADID, ",") > 0 Then
        sqlProperty = "select * from PE_Advertisement where ADID in (" & ADID & ")"
    Else
        sqlProperty = "select * from PE_Advertisement where ADID=" & ADID
    End If
    Set rsProperty = Server.CreateObject("ADODB.Recordset")
    rsProperty.Open sqlProperty, Conn, 1, 3
    Do While Not rsProperty.EOF
        Select Case Action
        Case "SetADPassed"
            rsProperty("Passed") = True
        Case "CancelADPassed"
            rsProperty("Passed") = False
        Case "DelAD"
            Call DelADID_Zone(rsProperty("ZoneID"), rsProperty("ADID"))
            rsProperty.Delete
        End Select
        rsProperty.Update
        rsProperty.MoveNext
    Loop
    rsProperty.Close
    Set rsProperty = Nothing
    Call WriteEntry(2, AdminName, "���ù�����Գɹ�")
    Call ClearSiteCache(0)

    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub PreviewAD()
    Dim ADID, sqlAD, rsAD
    ADID = Trim(Request("ADID"))
    If ADID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>������ʧ��</li>"
        Exit Sub
    Else
        ADID = PE_CLng(ADID)
    End If
    sqlAD = "select * from PE_Advertisement where ADID=" & ADID
    Set rsAD = Conn.Execute(sqlAD)
    If rsAD.BOF And rsAD.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���Ĺ�棡</li>"
        rsAD.Close
        Set rsAD = Nothing
        Exit Sub
    End If

    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='2' align='center'><strong>Ԥ�����----" & rsAD("ADName") & "</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg2'>"
    Response.Write "    <td height='25' align='center'>"
    Response.Write "      <a href='javascript:this.location.reload();'>ˢ��ҳ��</a>&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "      <a href='" & ComeUrl & "'>������ҳ</a>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr valign='top'>"
    Response.Write "    <td height='300'>" & GetADContent(rsAD("ADID")) & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"

    rsAD.Close
    Set rsAD = Nothing
End Sub

Function ShowJS_Tooltip()
    Response.Write "<div id=dHTMLADPreview style='Z-INDEX: 1000; LEFT: 0px; VISIBILITY: hidden; WIDTH: 10px; POSITION: absolute; TOP: 0px; HEIGHT: 10px'></DIV>"
    Response.Write "<SCRIPT language = 'JavaScript'>" & vbCrLf
    Response.Write "<!--" & vbCrLf
    Response.Write "var tipTimer;" & vbCrLf
    Response.Write "function locateObject(n, d)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "   var p,i,x;" & vbCrLf
    Response.Write "   if (!d) d=document;" & vbCrLf
    Response.Write "   if ((p=n.indexOf('?')) > 0 && parent.frames.length)" & vbCrLf
    Response.Write "   {d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}" & vbCrLf
    Response.Write "   if (!(x=d[n])&&d.all) x=d.all[n]; " & vbCrLf
    Response.Write "   for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];" & vbCrLf
    Response.Write "   for (i=0;!x&&d.layers&&i<d.layers.length;i++) x=locateObject(n,d.layers[i].document); return x;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function ShowADPreview(ADContent)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  showTooltip('dHTMLADPreview',event, ADContent, '#ffffff','#000000','#000000','6000')" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function showTooltip(object, e, tipContent, backcolor, bordercolor, textcolor, displaytime)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "   window.clearTimeout(tipTimer)" & vbCrLf
    Response.Write "   if (document.all) {" & vbCrLf
    Response.Write "       locateObject(object).style.top=document.body.scrollTop+event.clientY+20" & vbCrLf
    Response.Write "       locateObject(object).innerHTML='<table style=""font-family:����; font-size: 9pt; border: '+bordercolor+'; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px; background-color: '+backcolor+'"" width=""10"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td nowrap><font style=""font-family:����; font-size: 9pt; color: '+textcolor+'"">'+unescape(tipContent)+'</font></td></tr></table> '" & vbCrLf
    Response.Write "       if ((e.x + locateObject(object).clientWidth) > (document.body.clientWidth + document.body.scrollLeft)) {" & vbCrLf
    Response.Write "           locateObject(object).style.left = (document.body.clientWidth + document.body.scrollLeft) - locateObject(object).clientWidth-10;" & vbCrLf
    Response.Write "       } else {" & vbCrLf
    Response.Write "           locateObject(object).style.left=document.body.scrollLeft+event.clientX" & vbCrLf
    Response.Write "       }" & vbCrLf
    Response.Write "       locateObject(object).style.visibility='visible';" & vbCrLf
    Response.Write "       tipTimer=window.setTimeout(""hideTooltip('""+object+""')"", displaytime);" & vbCrLf
    Response.Write "       return true;" & vbCrLf
    Response.Write "   } else if (document.layers) {" & vbCrLf
    Response.Write "       locateObject(object).document.write('<table width=""10"" border=""0"" cellspacing=""1"" cellpadding=""1""><tr bgcolor=""'+bordercolor+'""><td><table width=""10"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr bgcolor=""'+backcolor+'""><td nowrap><font style=""font-family:����; font-size: 9pt; color: '+textcolor+'"">'+unescape(tipContent)+'</font></td></tr></table></td></tr></table>')" & vbCrLf
    Response.Write "       locateObject(object).document.close()" & vbCrLf
    Response.Write "       locateObject(object).top=e.y+20" & vbCrLf
    Response.Write "       if ((e.x + locateObject(object).clip.width) > (window.pageXOffset + window.innerWidth)) {" & vbCrLf
    Response.Write "           locateObject(object).left = window.innerWidth - locateObject(object).clip.width-10;" & vbCrLf
    Response.Write "       } else {" & vbCrLf
    Response.Write "           locateObject(object).left=e.x;" & vbCrLf
    Response.Write "       }" & vbCrLf
    Response.Write "       locateObject(object).visibility='show';" & vbCrLf
    Response.Write "       tipTimer=window.setTimeout(""hideTooltip('""+object+""')"", displaytime);" & vbCrLf
    Response.Write "       return true;" & vbCrLf
    Response.Write "   } else {" & vbCrLf
    Response.Write "       return true;" & vbCrLf
    Response.Write "   }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function hideTooltip(object) {" & vbCrLf
    Response.Write "    if (document.all) {" & vbCrLf
    Response.Write "        locateObject(object).style.visibility = 'hidden';" & vbCrLf
    Response.Write "        locateObject(object).style.left = 1;" & vbCrLf
    Response.Write "        locateObject(object).style.top = 1;" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    } else {" & vbCrLf
    Response.Write "        if (document.layers) {" & vbCrLf
    Response.Write "            locateObject(object).visibility = 'hide';" & vbCrLf
    Response.Write "            locateObject(object).left = 1;" & vbCrLf
    Response.Write "            locateObject(object).top = 1;" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        } else {" & vbCrLf
    Response.Write "            return true;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
End Function


'******************************************************************************************
'����Ϊ����λ�������õ��ĺ�������ȡ��صĹ���·��������Ƶ������λ���õ���Ϣ��
'******************************************************************************************

Function GetZoneManagePath(iChannelID)
    Dim strPath, sqlPath, rsPath
    strPath = "�����ڵ�λ�ã���վ������&nbsp;&gt;&gt;&nbsp;��λ����&nbsp;&gt;&gt;&nbsp;"
    If iChannelID = -1 Then
        strPath = strPath & "��վͨ�ð�λ"
    ElseIf iChannelID = -2 Then
        strPath = strPath & "��վ��ҳ��λ"
    Else
        sqlPath = "select ChannelName from PE_Channel where ChannelID=" & iChannelID
        Set rsPath = Conn.Execute(sqlPath)
        If rsPath.BOF And rsPath.EOF Then
            strPath = strPath & "�����Ƶ������"
        Else
            strPath = strPath & rsPath(0) & "��λ"
        End If
        rsPath.Close
        Set rsPath = Nothing
    End If
    strPath = strPath & "&nbsp;&gt;&gt;&nbsp;"
    If Keyword = "" Then
        strPath = strPath & "���а�λ"
    Else
        Select Case strField
            Case "ZoneName"
                strPath = strPath & "�����к��� <font color=red>" & Keyword & "</font> �İ�λ"
            Case "ZoneIntro"
                strPath = strPath & "����к��� <font color=red>" & Keyword & "</font> �İ�λ"
            Case Else
                strPath = strPath & "�����к��� <font color=red>" & Keyword & "</font> �İ�λ"
        End Select
    End If
    GetZoneManagePath = strPath
End Function


Function GetSettingItem(strSetting, sType, suf)
    Dim arrSetting
    GetSettingItem = ""
    If IsNull(strSetting) Or strSetting = "" Or Not IsNumeric(suf) Then
        Exit Function
    End If
    arrSetting = Split(strSetting, ",")
    If arrSetting(0) = CStr(sType) Then
        If suf > 0 And suf <= UBound(arrSetting) Then
            GetSettingItem = arrSetting(suf)
        End If
    End If
End Function

Function GetCurrentZoneJSName()
    GetCurrentZoneJSName = Year(Date) & Right("0" & Month(Date), 2) & "/" & PE_CLng(Conn.Execute("select max(ZoneID) from PE_AdZone")(0)) + 1 & ".js"
End Function

Function GetZoneJSName(ZoneJSName, iZoneID, UpdateTime)
    If CheckZoneJSName(ZoneJSName) = False Then
        GetZoneJSName = InstallDir & ADDir & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/" & iZoneID & ".js"
    Else
        GetZoneJSName = InstallDir & ADDir & "/" & ZoneJSName
    End If
End Function

Function GetZoneJS_Path(ZoneJSName, UpdateTime)
    GetZoneJS_Path = Year(UpdateTime) & Right("0" & Month(UpdateTime), 2)
    If CheckZoneJSName(ZoneJSName) = False Then Exit Function
    Dim arrJSName
    arrJSName = Split(ZoneJSName, "/")
    If UBound(arrJSName) = 1 Then
        GetZoneJS_Path = arrJSName(0) & "/"
    End If
End Function

Function GetZoneJS_Name(ZoneJSName, iZoneID)
    If CheckZoneJSName(ZoneJSName) = False Then
        GetZoneJS_Name = iZoneID & ".js"
        Exit Function
    End If
    Dim arrJSName
    arrJSName = Split(ZoneJSName, "/")
    If UBound(arrJSName) = 1 Then
        GetZoneJS_Name = arrJSName(1)
    Else
        GetZoneJS_Name = ZoneJSName
    End If
End Function

Function GetZoneJSNameCode(ZoneJSName)
    GetZoneJSNameCode = "{$InstallDir}{$ADDir}/" & ZoneJSName
End Function

Function GetTemplateName(iZoneType)
    GetTemplateName = InstallDir & ADDir & "/ADTemplate/Template_" & ZoneConfig(iZoneType, 0) & ".js"
End Function

Function GetZoneJSTemplate(ZoneType)
    Dim TemplateFile
    GetZoneJSTemplate = ""
    If ObjInstalled_FSO = False Then
        Exit Function
    End If

    TemplateFile = GetTemplateName(ZoneType)
    GetZoneJSTemplate = ReadFileContent(TemplateFile)
End Function

Function RequestSetting(name)
    Dim str
    str = Trim(Request.Form(name))
    If Not IsNull(str) Then
        RequestSetting = Replace(str, ",", "��")
    Else
        RequestSetting = ""
    End If
End Function

Function CheckZoneJSName(ZoneJSName)
    CheckZoneJSName = False
    If ZoneJSName = "" Or IsNull(ZoneJSName) Then Exit Function
    Dim retVal
    regEx.Pattern = "^[\w-]+/?\w+\.js$"
    retVal = regEx.Test(ZoneJSName)
    If retVal Then CheckZoneJSName = True
End Function


'******************************************************************************************
'����Ϊ���������õ��ĺ�������ȡ��صĹ���·�����������õ���Ϣ��
'******************************************************************************************

Function GetADManagePath()
    Dim strPath, sqlPath, rsPath
    strPath = "�����ڵ�λ�ã���վ������&nbsp;&gt;&gt;&nbsp;"
    If ZoneID <> "" Then
        Dim nZone
        Set nZone = Conn.Execute("select ZoneName from PE_AdZone where ZoneID=" & PE_CLng(ZoneID))
        If Not nZone.BOF And Not nZone.EOF Then
            strPath = strPath & nZone("ZoneName") & "&nbsp;&gt;&gt;&nbsp;"
        End If
    End If
    If Keyword = "" Then
        strPath = strPath & "���й��"
    Else
        Select Case strField
            Case "ADName"
                strPath = strPath & "�����к��� <font color=red>" & Keyword & "</font> �Ĺ��"
            Case "ADIntro"
                strPath = strPath & "����к��� <font color=red>" & Keyword & "</font> �Ĺ��"
            Case Else
                strPath = strPath & "�����к��� <font color=red>" & Keyword & "</font> �Ĺ��"
        End Select
    End If
    GetADManagePath = strPath
End Function

Function GetZone_Option(arrZoneID)
    Dim strTemp, rsZone, sqlZone
    sqlZone = "select ZoneID,ChannelID,ZoneName from PE_AdZone order by ChannelID asc,ZoneID desc"
    Set rsZone = Conn.Execute(sqlZone)
    Do While Not rsZone.EOF
        strTemp = strTemp & "<option value='" & rsZone("ZoneID") & "'"
        If FoundInArr(arrZoneID, rsZone("ZoneID"), ",") = True Then strTemp = strTemp & " selected"
        strTemp = strTemp & ">��" & GetChannelName(rsZone("ChannelID")) & "��" & rsZone("ZoneName") & "</option>"
        rsZone.MoveNext
    Loop
    rsZone.Close
    Set rsZone = Nothing
    GetZone_Option = strTemp
End Function

Function GetADType(iADType)
    If iADType = 1 Then
        GetADType = "ͼƬ"
    ElseIf iADType = 2 Then
        GetADType = "����"
    ElseIf iADType = 3 Then
        GetADType = "�ı�"
    ElseIf iADType = 4 Then
        GetADType = "����"
    ElseIf iADType = 5 Then
        GetADType = "ҳ��"
    End If
End Function

Function GetADType_Option(iADType)
    Dim strADType
    strADType = strADType & "<input type='radio' name='ADType' value='1' onclick='Change_ADType();'"
    If iADType = 1 Then strADType = strADType & " checked"
    strADType = strADType & "> ͼƬ&nbsp;&nbsp;"
    strADType = strADType & "<input type='radio' name='ADType' value='2' onclick='Change_ADType();'"
    If iADType = 2 Then strADType = strADType & " checked"
    strADType = strADType & "> ����&nbsp;&nbsp;"
    strADType = strADType & "<input type='radio' name='ADType' value='3' onclick='Change_ADType();'"
    If iADType = 3 Then strADType = strADType & " checked"
    strADType = strADType & "> �ı�&nbsp;&nbsp;"
    strADType = strADType & "<input type='radio' name='ADType' value='4' onclick='Change_ADType();'"
    If iADType = 4 Then strADType = strADType & " checked"
    strADType = strADType & "> ����&nbsp;&nbsp;"
    strADType = strADType & "<input type='radio' name='ADType' value='5' onclick='Change_ADType();'"
    If iADType = 5 Then strADType = strADType & " checked"
    strADType = strADType & "> ҳ��&nbsp;&nbsp;"
    GetADType_Option = strADType
End Function

Function GetADContent(ADID)
    If IsNull(ADID) Or ADID = "" Then
        Exit Function
    Else
        ADID = PE_CLng(ADID)
    End If

    Dim rsAD, sqlAD, strAD, strTemp
    sqlAD = "select * from PE_Advertisement where ADID=" & ADID
    Set rsAD = Server.CreateObject("ADODB.Recordset")
    rsAD.Open sqlAD, Conn, 1, 1
    If Not (rsAD.BOF And rsAD.EOF) Then
        Select Case rsAD("ADType")
        Case 1
            strTemp = strTemp & "<img src='" & GetADPicUrl(rsAD("ImgUrl")) & "'"
            If rsAD("ImgWidth") > 0 Then strTemp = strTemp & " width='" & rsAD("ImgWidth") & "'"
            If rsAD("ImgHeight") > 0 Then strTemp = strTemp & " height='" & rsAD("ImgHeight") & "'"
            strTemp = strTemp & " border='0'>"
            If IsNull(rsAD("LinkUrl")) Or rsAD("LinkUrl") = "" Then
                strAD = strTemp
            Else
                strAD = "<a href='" & rsAD("LinkUrl") & ""
                If rsAD("LinkTarget") = 0 Then
                    strAD = strAD & " target='_self'"
                Else
                    strAD = strAD & " target='_blank'"
                End If
                If rsAD("LinkAlt") <> "" Then strAD = strAD & " title='" & rsAD("LinkAlt") & "'"
                strAD = strAD & ">" & strTemp & "</a>"
            End If
        Case 2
            strAD = "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0'"
            If rsAD("ImgWidth") > 0 Then strAD = strAD & " width='" & rsAD("ImgWidth") & "'"
            If rsAD("ImgHeight") > 0 Then strAD = strAD & " height='" & rsAD("ImgHeight") & "'"
            strAD = strAD & "><param name='movie' value='" & GetADPicUrl(rsAD("ImgUrl")) & "'>"
            If rsAD("FlashWmode") = 1 Then strAD = strAD & "<param name='wmode' value='transparent'>"
            strAD = strAD & "<param name='quality' value='autohigh'>"
            strAD = strAD & "<embed src='" & GetADPicUrl(rsAD("ImgUrl")) & "' quality='autohigh'"
            strAD = strAD & " pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'"
            If rsAD("FlashWmode") = 1 Then strAD = strAD & " wmode='transparent'"
            If rsAD("ImgWidth") > 0 Then strAD = strAD & " width='" & rsAD("ImgWidth") & "'"
            If rsAD("ImgHeight") > 0 Then strAD = strAD & " height='" & rsAD("ImgHeight") & "'"
            strAD = strAD & "></embed></object>"
        Case 3
            strAD = PE_HTMLEncode(rsAD("ADIntro"))
        Case 4
            strAD = rsAD("ADIntro")
        Case 5
            strAD = rsAD("ADIntro")
        End Select
    End If
    rsAD.Close
    Set rsAD = Nothing
    GetADContent = strAD
End Function

Function GetADPicUrl(ADPicUrl)
    If LCase(Left(ADPicUrl, Len("UploadADPic"))) = "uploadadpic" Then
        GetADPicUrl = InstallDir & ADDir & "/" & ADPicUrl
    Else
        GetADPicUrl = ADPicUrl
    End If
End Function
%>
