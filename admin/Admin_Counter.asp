<!--#include file="Admin_Common.asp"-->
<!--#include file="../count/conn_counter.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 2      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 0   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = "Counter"   '����Ȩ��

Private rs, sql
Private Search, strGuide, TitleRight
Private RegCount_Fill
Private MasterTimeZone, OnlineTime, IntervalNum, VisitRecord, KillRefresh, OldTotalNum, OldTotalView
Private QDay, QYear, QMonth, QWeek, SYear, SMonth
Private TotalNum, StatItem, Item, ItemNum, Percent, Barwidth, MaxWidth, Assay, Rows, i, DispRow


QDay = Request("QYear") & "-" & Request("QMonth") & "-" & Request("QDay")
QMonth = Request("QYear") & "-" & Request("QMonth")
QYear = Request("QYear")
Select Case Request("Type")
Case 1
    Action = "StatDay"
Case 2
    Action = "StatMonth"
Case 3
    Action = "StatYear"
End Select

strFileName = "Admin_Counter.asp?Action=" & Action
If Request("page") <> "" Then
    CurrentPage = PE_CLng1(Trim(Request("page")))
Else
    CurrentPage = 1
End If

MaxWidth = 220      '����ͳ�����ı��Ŀ��
TotalNum = 0

'�����ݿ�
Call OpenConn_Counter
If FoundErr = True Then Response.End

sql = "select * from PE_StatInfoList"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, Conn_Counter, 1, 1
If Not rs.BOF And Not rs.EOF Then
    MasterTimeZone = rs("MasterTimeZone")
    OnlineTime = rs("OnlineTime")
    RegCount_Fill = rs("RegFields_Fill")
End If
rs.Close
Set rs = Nothing

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>��վͳ�ƹ���</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "<script>" & vbCrLf
Response.Write "function change_type()" & vbCrLf
Response.Write "{ " & vbCrLf
Response.Write "    select_type=form1.type.options[form1.type.selectedIndex].text;" & vbCrLf
Response.Write "    switch(select_type)" & vbCrLf
Response.Write "    { " & vbCrLf
Response.Write "        case '�ձ���' :form1.qmonth.disabled=0;form1.qday.disabled=0;break;" & vbCrLf
Response.Write "        case '�±���' :form1.qmonth.disabled=0;form1.qday.disabled=1;break;" & vbCrLf
Response.Write "        case '�걨��' :form1.qmonth.disabled=1;form1.qday.disabled=1;break;" & vbCrLf
Response.Write "    } " & vbCrLf
Response.Write "} " & vbCrLf
Response.Write "function change_it()" & vbCrLf
Response.Write "{ " & vbCrLf
Response.Write "    select_type=form1.type.options[form1.type.selectedIndex].text;" & vbCrLf
Response.Write "    if (select_type=='�ձ���')" & vbCrLf
Response.Write "    {" & vbCrLf
Response.Write "        select_item_y=form1.qyear.options[form1.qyear.selectedIndex].text;" & vbCrLf
Response.Write "        month29=select_item_y%4;" & vbCrLf
Response.Write "        select_item_m=form1.qmonth.options[form1.qmonth.selectedIndex].text;" & vbCrLf
Response.Write "        switch(select_item_m)" & vbCrLf
Response.Write "        { " & vbCrLf
Response.Write "            case '2' :if (month29==0) {MD(29)}  else {MD(28)};break;" & vbCrLf
Response.Write "            case '4' : " & vbCrLf
Response.Write "            case '6' : " & vbCrLf
Response.Write "            case '9' : " & vbCrLf
Response.Write "            case '11' : MD(30);break; " & vbCrLf
Response.Write "            default : MD(31);break; " & vbCrLf
Response.Write "        }" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "} " & vbCrLf
Response.Write "function MD(days)" & vbCrLf
Response.Write "{ " & vbCrLf
Response.Write "    j=form1.qday.options.length; " & vbCrLf
Response.Write "    for(k=0;k<j;k++) form1.qday.options.remove(0); " & vbCrLf
Response.Write "    for(i=0;i<days;i++)" & vbCrLf
Response.Write "    { " & vbCrLf
Response.Write "        var day=document.createElement('OPTION'); " & vbCrLf
Response.Write "        form1.qday.options.add(day); " & vbCrLf
Response.Write "        day.innerText=i+1; " & vbCrLf
Response.Write "        form1.qday.selectedIndex=0" & vbCrLf
Response.Write "    } " & vbCrLf
Response.Write "} " & vbCrLf
Response.Write "function setFileFileds(num){    " & vbCrLf
Response.Write "     var str="""";" & vbCrLf
Response.Write "     if (num==1){" & vbCrLf
Response.Write "     str=str+=""<s""+ ""c"" + ""r"" + ""i"" + ""pt src='{$InstallDir}Count/CounterLink.asp?style=simple'></sc"" + ""ri"" +""pt>"";" & vbCrLf
Response.Write "     }" & vbCrLf
Response.Write "     else if(num==2){" & vbCrLf
Response.Write "     str=str+=""<s""+ ""c"" + ""r"" + ""i"" + ""pt src='{$InstallDir}Count/CounterLink.asp?style=common'></sc"" + ""ri"" +""pt>"";" & vbCrLf
Response.Write "     }" & vbCrLf
Response.Write "     else if(num==3){" & vbCrLf
Response.Write "     str=str+=""<s""+ ""c"" + ""r"" + ""i"" + ""pt src='{$InstallDir}Count/CounterLink.asp?style=all'></sc"" + ""ri"" +""pt>"";" & vbCrLf
Response.Write "     }" & vbCrLf
Response.Write "     else if(num==4){" & vbCrLf
Response.Write "     str=str+=""<s""+ ""c"" + ""r"" + ""i"" + ""pt src='{$InstallDir}Count/CounterLink.asp?style=none'></sc"" + ""ri"" +""pt>"";" & vbCrLf
Response.Write "     }" & vbCrLf
Response.Write "     document.form1.selectKey.value=str;" & vbCrLf
Response.Write "}" & vbCrLf
    
Response.Write "</script>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>"
If Action = "ShowConfig" Or Action = "SaveConfig" Or Action = "Init" Or Action = "DoInit" Or Action = "IPAdd" Or Action = "IPManage" Or Action = "SaveIPAdd" Or Action = "SearchIP" Or Action = "editIP" Or Action = "SaveIPedit" Or Action = "delIP" Or Action = "Compact" Or Action = "CompactData" Or Action = "Import" Or Action = "DoImport" Or Action = "Export" Or Action = "DoExport" Then
    Call ShowPageTitle("�� վ ͳ �� �� ��", 10025)
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='70' height='30'><strong>��������</strong></td>" & vbCrLf
    Response.Write "    <td height='30'><a href='Admin_Counter.asp?Action=ShowConfig'>��վͳ����Ϣ����</a> | <a href='Admin_Counter.asp?Action=IPAdd'>ͳ��IP�����</a> | <a href='Admin_Counter.asp?Action=IPManage'>ͳ��IP�����</a> | <a href='Admin_Counter.asp?Action=Compact'>ѹ��ͳ�����ݿ�</a> | <a href='Admin_Counter.asp?Action=Init'>ͳ�����ݳ�ʼ��</a>  | <a href='Admin_Counter.asp?Action=Export'>����IP���ݿ�</a> | <a href='Admin_Counter.asp?Action=Import'>����IP���ݿ�</a>    </td>" & vbCrLf
    Response.Write "  </tr>"
Else
    Call ShowPageTitle("�� վ ͳ �� �� ��", 10025)
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td height='30'>"
    Response.Write "    <a href='Admin_Counter.asp?Action=Infolist'>�ۺ�ͳ��</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=FVisitor'>���ʼ�¼</a>&nbsp;|"
    If FoundInArr(RegCount_Fill, "FVisit", ",") = True Then
    Response.Write "    <a href='Admin_Counter.asp?Action=FCounter'>���ʴ���</a>&nbsp;|"
    End If
    Response.Write "    <a href='Admin_Counter.asp?Action=StatYear'>�� �� ��</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=StatAllYear'>ȫ �� ��</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=StatMonth'>�� �� ��</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=StatAllMonth'>ȫ �� ��</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=StatWeek'>�� �� ��</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=StatAllWeek'>ȫ �� ��</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=StatDay'>�� �� ��</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=StatAllDay'>ȫ �� ��</a>&nbsp;|<br>"
    If FoundInArr(RegCount_Fill, "IsCountOnline", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FOnline'>�����û�</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FIP", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FIp'>IP �� ַ</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FAddress", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FAddress'>��ַ����</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FTimezone", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FTimezone'>ʱ������</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FKeyword", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FKeyword'>�� �� ��</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FWeburl", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FWeburl'>������վ</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FRefer", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FReferer'>����ҳ��</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FSystem", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FSystem'>����ϵͳ</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FBrowser", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FBrowser'>� �� ��</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FMozilla", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FMozilla'>�ִ�����</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FScreen", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FScreen'>��Ļ��С</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FColor", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FColor'>��Ļɫ��</a>&nbsp;|"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
End If
Response.Write "</table>"
Response.Write "<br>"
  
Select Case Action
Case "ShowConfig"
    Call ShowConfig
Case "SaveConfig"
    Call SaveConfig
Case "Infolist"
    Call Infolist
Case "FVisitor"
    Call FVisitor
Case "FCounter"
    Call FCounter
Case "StatYear"
    Call StatYear
Case "StatAllYear"
    Call StatAllYear
Case "StatMonth"
    Call StatMonth
Case "StatAllMonth"
    Call StatAllMonth
Case "StatWeek"
    Call StatWeek
Case "StatAllWeek"
    Call StatAllWeek
Case "StatDay"
    Call StatDay
Case "StatAllDay"
    Call StatAllDay
Case "FIp"
    Call FIP
Case "FOnline"
    Call FOnline
Case "FAddress"
    Call FAddress
Case "FTimezone"
    Call FTimezone
Case "FWeburl"
    Call FWeburl
Case "FKeyword"
    Call FKeyword
Case "FReferer"
    Call FReferer
Case "FSystem"
    Call FSystem
Case "FBrowser"
    Call FBrowser
Case "FMozilla"
    Call FMozilla
Case "FScreen"
    Call FScreen
Case "FColor"
    Call FColor
Case "Init"
    Call Init
Case "DoInit"
    Call DoInit
Case "ClientDetail"
    Call ClientDetail
Case "IPAdd"
    Call IPAdd
Case "SaveIPAdd"
    Call SaveIPAdd
Case "IPManage", "SearchIP"
    Call IPManage
Case "editIP"
    Call editIP
Case "SaveIPedit"
    Call SaveIPedit
Case "delIP"
    Call delIP
Case "Compact"
    Call ShowCompact
Case "CompactData"
    Call CompactData
'Case "AutoAnalyse"
    'Call AutoAnalyse
'Case "DoAutoAnalyse"
    'Call DoAutoAnalyse
Case "Export"
    Call Export
Case "Import"
    Call Import			
Case "DoImport"
    Call DoImport	
Case "DoExport"
    Call DoExport			
Case Else
    Call Infolist
End Select

Call CloseConn_Counter

If Not (Action = "ShowConfig" Or Action = "SaveConfig" Or Action = "Init" Or Action = "DoInit" Or Action = "ClientDetail" Or Action = "IPAdd" Or Action = "IPManage" Or Action = "SaveIPAdd" Or Action = "SearchIP" Or Action = "editIP" Or Action = "SaveIPedit" Or Action = "delIP" Or Action = "Compact" Or Action = "CompactData") Then
    Call HistoryList
End If

If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub ShowConfig()
    sql = "select * from PE_StatInfoList"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, Conn_Counter, 1, 3
    If rs.BOF And rs.EOF Then
        Response.Write "<li>��վͳ���������ݶ�ʧ��"
        Exit Sub
    Else
        MasterTimeZone = rs("MasterTimeZone")
        OnlineTime = rs("OnlineTime")
        IntervalNum = rs("IntervalNum")
        VisitRecord = rs("VisitRecord")
        KillRefresh = rs("KillRefresh")
        OldTotalNum = rs("OldTotalNum")
        OldTotalView = rs("OldTotalView")
        RegCount_Fill = rs("RegFields_Fill")
    End If
    rs.Close
    Set rs = Nothing
    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "function ConfirmModify(){" & vbCrLf
    Response.Write "  if(confirm('ǿ�ҽ��龡��ѡ���ٵ�ͳ�ƹ�����Ŀ�����һ���������ã�����'))" & vbCrLf
    Response.Write "      return true;" & vbCrLf
    Response.Write "  else" & vbCrLf
    Response.Write "      return false;" & vbCrLf
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
    Response.Write "</SCRIPT>" & vbCrLf
    Response.Write "<table width='100%'><tr><td align='left'>�����ڵ�λ�ã���վͳ����Ϣ����</td></tr></table>"
    Response.Write "<form method='POST' action='Admin_Counter.asp?Action=SaveConfig' id='form1' name='form1' onsubmit='return ConfirmModify();'>" & vbCrLf

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>������Ϣ</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>��ʼ������</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>������Ŀ</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>���ô���</td>" & vbCrLf

    Response.Write "    <td>&nbsp;</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf

    Response.Write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='0' class='border'>"
    Response.Write "  <tr align='center'>"
    Response.Write "    <td class='tdbg' height='200' valign='top'>"
    Response.Write "      <table width='98%' border='0' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>"
    Response.Write "        <tbody id='Tabs' style='display:'>" & vbCrLf
    'Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>" & vbCrLf
    'Response.Write "    <tr class='topbg'> " & vbCrLf
    'Response.Write "      <td height='22' colspan='4'> <a name='SiteCountInfo'></a><strong>��վͳ����Ϣ����</strong></td>" & vbCrLf
    'Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong> ����������ʱ����</strong></td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='MasterTimeZone' type='text' id='MasterTimeZone' value='" & MasterTimeZone & "' size='20' maxlength='50'>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong>�����û��ı���ʱ�䣺</strong><br>" & vbCrLf
    Response.Write "      �û��л�ҳ����������վ���߹ر������������������������ʱ����ɾ�����û���������ԽС����վͳ�Ƶĵ�ǰʱ����������Խ׼ȷ��������Խ����վͳ�Ƶ���������Խ�ࡣ" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='OnlineTime' type='text' id='OnlineTime' value='" & OnlineTime & "' size='20' maxlength='50'>" & vbCrLf
    Response.Write "      ��      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong>�Զ�������߼����</strong><br>" & vbCrLf
    Response.Write "      �ͻ����������ÿ������ʱ����������ύһ��������Ϣ��ͬʱ������������Ϊ���ߣ�������ԽС����������Ҫ���������Խ�ࡣ</td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='Interval' type='text' id='Interval' value='60' size='20' maxlength='50' disabled>" & vbCrLf
    Response.Write "        ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong>�Զ�������߼��ѭ��������</strong><br>" & vbCrLf
    Response.Write "      ����Ϊ�˷�ֹ�û�����ҳ������ʱ�����κλ�����á��ͻ����������������ύ������Ϣ���������˴���������ֹͣ�ύ��" & vbCrLf
    Response.Write "</td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='IntervalNum' type='text' id='IntervalNum' value='" & IntervalNum & "' size='20' maxlength='50'>" & vbCrLf
    Response.Write "        ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong>�������ʼ�¼����</strong><br>" & vbCrLf
    Response.Write " ���������ϸ(������)��Ŀ����</td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='VisitRecord' type='text' id='VisitRecord' value='" & VisitRecord & "' size='20' maxlength='50'>" & vbCrLf
    Response.Write "        ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong> ��������IP��(����20С��800������)�� </strong><br>" & vbCrLf
    Response.Write "      �������á���������ͳ�ơ�����ʱ��ϵͳ���Ա���������IP�ķ�ʽ����ֹˢ�£���ͬһ��IP���ʶ�λ�������վ���л�ҳ�棬��ֻ������������������������    </td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='KillRefresh' type='text' id='KillRefresh' value='" & KillRefresh & "' size='20' maxlength='50'>" & vbCrLf
    Response.Write "        ��" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf

    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong>ʹ�ñ�ϵͳǰ�ķ�������</strong>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='OldTotalNum' type='text' id='OldTotalNum' value='" & OldTotalNum & "' size='20' maxlength='9'>" & vbCrLf
    Response.Write "        �˴�" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong> ʹ�ñ�ϵͳǰ���������</strong>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='OldTotalView' type='text' id='OldTotalView' value='" & OldTotalView & "' size='20' maxlength='9'>" & vbCrLf
    Response.Write "        �˴�" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf

    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong>������Ŀ:</strong><br>" & vbCrLf
    Response.Write "      ͳ��̫�����Ŀ����������ٶȣ��ķ�̫����վ��Դ��һ��ʱ�䲻������Ĺ�����Ŀ���鲻Ҫ���ã�<br><font color='red'>ǿ�ҽ��龡��ѡ���ٵĹ�����Ŀ�����һ���������ã�����</font><br>" & vbCrLf
    Response.Write "      </td><td>" & vbCrLf
    Response.Write "        <table width='100%'><tr>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "      <input name='RegFields_Fill' type='checkbox' value='IsCountOnline'" & vbCrLf
    If FoundInArr(RegCount_Fill, "IsCountOnline", ",") = True Then Response.Write " checked"
    Response.Write "      >���á���������ͳ�ơ�����</td><td><input name='RegFields_Fill' type='checkbox' value='FIP'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FIP", ",") = True Then Response.Write " checked"
    Response.Write "      >�ͻ���IP��ַ����</td><td><input name='RegFields_Fill' type='checkbox' value='FAddress'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FAddress", ",") = True Then Response.Write " checked"
    Response.Write "      >�ͻ��˵�ַ���� </td></tr><tr class='tdbg'><td><input name='RegFields_Fill' type='checkbox' value='FRefer'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FRefer", ",") = True Then Response.Write " checked"
    Response.Write "      >�ͻ�������ҳ����� </td><td><input name='RegFields_Fill' type='checkbox' value='FTimezone'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FTimezone", ",") = True Then Response.Write " checked"
    Response.Write "      >�ͻ���ʱ������ </td><td><input name='RegFields_Fill' type='checkbox' value='FWeburl'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FWeburl", ",") = True Then Response.Write " checked"
    Response.Write "      >�ͻ���������վ����  </td></tr><tr class='tdbg'><td><input name='RegFields_Fill' type='checkbox' value='FBrowser'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FBrowser", ",") = True Then Response.Write " checked"
    Response.Write "      >�ͻ������������ </td><td><input name='RegFields_Fill' type='checkbox' value='FMozilla'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FMozilla", ",") = True Then Response.Write " checked"
    Response.Write "      >�ͻ����ִ����� </td><td><input name='RegFields_Fill' type='checkbox' value='FSystem'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FSystem", ",") = True Then Response.Write " checked"
    Response.Write "      >�ͻ��˲���ϵͳ���� </td></tr> <tr class='tdbg'><td><input name='RegFields_Fill' type='checkbox' value='FScreen'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FScreen", ",") = True Then Response.Write " checked"
    Response.Write "      >�ͻ�����Ļ��С���� </td><td><input name='RegFields_Fill' type='checkbox' value='FColor'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FColor", ",") = True Then Response.Write " checked"
    Response.Write "      >�ͻ�����Ļɫ�ʷ���  </td><td><input name='RegFields_Fill' type='checkbox' value='FKeyword'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FKeyword", ",") = True Then Response.Write " checked"
    Response.Write "      >�����ؼ��ʷ��� </td></tr> <tr class='tdbg'><td><input name='RegFields_Fill' type='checkbox' value='FVisit'"
    If FoundInArr(RegCount_Fill, "FVisit", ",") = True Then Response.Write " checked"
    Response.Write "      >���ʴ���ͳ�Ʒ��� </td><td><input name='RegFields_Fill' type='checkbox' value='FYesterDay'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FYesterDay", ",") = True Then Response.Write " checked"
    Response.Write "      >��������ͳ��  </td><td>" & vbCrLf

    Response.Write "      </td></tr>" & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "   </td></tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong>ͳ�Ƽ����������ͣ�</strong><br>" & vbCrLf
    Response.Write "      [����ѡ������Ҫ�������Ϣ����]</td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "      <select name='select'  onChange='setFileFileds(this.value)'>" & vbCrLf
    Response.Write "        <option value='1' selected>��ʾ����ʽ��Ϣ</option>" & vbCrLf
    Response.Write "        <option value='2'>��ʾ��ͨ��ʽ��Ϣ</option>" & vbCrLf
    Response.Write "        <option value='3'>��ʾ������ʽ��Ϣ</option>" & vbCrLf
    Response.Write "        <option value='4'>ͳ�Ƶ�����ʾ��Ϣ</option>" & vbCrLf
    Response.Write "      </select>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong> ��ʾ���ݴ��룺</strong><br>" & vbCrLf
    Response.Write "      �뽫�˴��뿽��������Ҫ��ͳ�Ƶ�ҳ�棬�˴��벻������������˴˴����ҳ�����ͳ�����ݣ����һ��Ը�ҳ�������<br></td>" & vbCrLf
    Response.Write "      <td colspan='3'><textarea name='selectKey' cols='50' rows='5' id='selectKey'><script src='{$InstallDir}Count/CounterLink.asp?style=simple'></script></textarea>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong> ǰ̨��ʾ�������Ӵ��룺</strong><br>" & vbCrLf
    Response.Write "      �뽫�˴��뿽��������Ҫ��ʾ�����б����ӵ�ģ���У��˴��������������˴˴����ҳ����ʾ�����б����ӣ������Ը�ҳ�������<br></td>" & vbCrLf
    Response.Write "      <td colspan='3'><textarea name='LinkContent' cols='50' rows='5' id='LinkContent'><a href='{$InstallDir}Count/ShowOnline.asp' target='_blank'>��վ���������ϸ�б�</a></textarea>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    
    Response.Write "  <p align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveConfig'>" & vbCrLf
    Response.Write "        <input name='cmdSave' type='submit' id='cmdSave' value=' �������� '>" & vbCrLf
    Response.Write "      </p>" & vbCrLf

    Response.Write "</form>" & vbCrLf

End Sub
Sub SaveConfig()
    Dim sqlConfig, rsConfig
    sqlConfig = "select * from PE_StatInfoList"
    Set rsConfig = Server.CreateObject("ADODB.Recordset")
    rsConfig.Open sqlConfig, Conn_Counter, 1, 3
    If rsConfig.BOF And rsConfig.EOF Then
        rsConfig.AddNew
    End If
    rsConfig("MasterTimeZone") = PE_CLng(Trim(Request("MasterTimeZone")))
    rsConfig("OnlineTime") = PE_CLng(Trim(Request("OnlineTime")))
    rsConfig("VisitRecord") = PE_CLng(Trim(Request("VisitRecord")))
    rsConfig("IntervalNum") = PE_CLng(Trim(Request("IntervalNum")))
    rsConfig("KillRefresh") = PE_CLng(Trim(Request("KillRefresh")))
    rsConfig("OldTotalNum") = PE_CLng(Trim(Request("OldTotalNum")))
    rsConfig("OldTotalView") = PE_CLng(Trim(Request("OldTotalView")))
    rsConfig("RegFields_Fill") = ReplaceBadChar(Trim(Request("RegFields_Fill")))
    rsConfig.Update
    rsConfig.Close
    Set rsConfig = Nothing
    Call WriteSuccessMsg("��վͳ�����ñ���ɹ���", ComeUrl)
End Sub


Sub Infolist()
    Dim StartDate, StatDayNum, AllNum, TotalView, CountNum, AveDayNum, DayNum
    Dim MonthMaxNum, MonthMaxDate, DayMaxNum, DayMaxDate, HourMaxNum, HourMaxTime, ZoneNum, ChinaNum, OtherNum
    Dim MaxBrw, MaxBrwNum, MaxSys, MaxSysNum, MaxScr, MaxScrNum, MaxAre, MaxAreNum, MaxWeb, MaxWebNum, MaxColor, MaxColorNum
    strGuide = "��վ�ۺ�ͳ����Ϣ"
    sql = "Select * From PE_StatInfoList"

    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn_Counter, 1, 1
    If Not rs.BOF And Not rs.EOF Then
        DayNum = rs("DayNum")
        AllNum = rs("TotalNum")
        TotalView = rs("TotalView")
        MonthMaxNum = rs("MonthMaxNum")
        MonthMaxDate = rs("MonthMaxDate")
        DayMaxNum = rs("DayMaxNum")
        DayMaxDate = rs("DayMaxDate")
        HourMaxNum = rs("HourMaxNum")
        HourMaxTime = rs("HourMaxTime")
        ChinaNum = rs("ChinaNum")
        OtherNum = rs("OtherNum")
        StartDate = rs("StartDate")
        StatDayNum = DateDiff("D", StartDate, Date) + 1
        If StatDayNum <= 0 Or IsNumeric(StatDayNum) = 0 Then
           AveDayNum = StatDayNum
        Else
           AveDayNum = CInt(AllNum / StatDayNum)
        End If
    End If
    rs.Close
    sql = "Select * From PE_StatVisit"
    rs.Open sql, Conn_Counter, 1, 1
    If Not rs.BOF And Not rs.EOF Then
        For i = 1 To 10
            CountNum = CountNum + rs("" & i & "")
        Next
    Else
      CountNum = 0
    End If
    rs.Close
    sql = "Select Top 1 * From PE_StatBrowser Order By TBrwNum DESC"
    rs.Open sql, Conn_Counter, 1, 1
    If Not rs.BOF And Not rs.EOF Then
        MaxBrw = rs("TBrowser")
        MaxBrwNum = rs("TBrwNum")
    End If
    rs.Close
    sql = "Select Top 1 * From PE_StatSystem Order By TSysNum DESC"
    rs.Open sql, Conn_Counter, 1, 1
    If Not rs.BOF And Not rs.EOF Then
        MaxSys = rs("TSystem")
        MaxSysNum = rs("TSysNum")
    End If
    rs.Close
    sql = "Select Top 1 * From PE_StatScreen Order By TScrNum DESC"
    rs.Open sql, Conn_Counter, 1, 1
    If Not rs.BOF And Not rs.EOF Then
        MaxScr = rs("TScreen")
        MaxScrNum = rs("TScrNum")
    End If
    rs.Close
    sql = "Select Top 1 * From PE_StatColor Order By TColNum DESC"
    rs.Open sql, Conn_Counter, 1, 1
    If Not rs.BOF And Not rs.EOF Then
        MaxColor = rs("TColor")
        MaxColorNum = rs("TColNum")
    End If
    rs.Close
    sql = "Select Top 1 * From PE_StatAddress Order By TAddNum DESC"
    rs.Open sql, Conn_Counter, 1, 1
    If Not rs.BOF And Not rs.EOF Then
        MaxAre = rs("TAddress")
        MaxAreNum = rs("TAddNum")
    End If
    rs.Close
    sql = "Select Top 1 * From PE_StatWeburl Order By TWebNum DESC"
    rs.Open sql, Conn_Counter, 1, 1
    If Not rs.BOF And Not rs.EOF Then
        MaxWeb = rs("TWeburl")
        MaxWebNum = rs("TWebNum")
    End If
    rs.Close
    TitleRight = "��ʼͳ�����ڣ�<font color=blue>" & StartDate & "</font>"

    Response.Write "<table width='100%'><tr><td align='left'>�����ڵ�λ�ã���վͳ�ƹ���&nbsp;&gt;&gt;&nbsp;" & Search & strGuide & "</td><td align='right'>" & TitleRight & "</td></tr></table>"
    Response.Write "<table border=0 cellpadding=2 cellspacing=1 width='100%' bgcolor='#FFFFFF' class='border'>"
    Response.Write "  <tr class='title' align='center'>"
    Response.Write "    <td align=center width='20%' height='22'>ͳ����</td>"
    Response.Write "    <td align=center width='30%'>ͳ������</td>"
    Response.Write "    <td width='20%'>ͳ����</td>"
    Response.Write "    <td align='center' width='30%'>ͳ������</td>"
    Response.Write "  </tr>"
    Response.Write "  <tbody>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align=center width='20%'>��ͳ������</td>"
    Response.Write "    <td align=center width='30%'>" & StatDayNum & "</td>"
    Response.Write "    <td align=center width='20%'>����·���</td>"
    Response.Write "    <td align=center width='30%'>" & MonthMaxNum & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class=tdbg>"
    Response.Write "    <td align=center width='20%'>�ܷ�����</td>"
    Response.Write "    <td align=center width='30%'>" & AllNum & "</td>"
    Response.Write "    <td align=center width='20%'>����·����·�</td>"
    Response.Write "    <td align=center width='30%'>" & MonthMaxDate & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align=center width='20%'>�ܷ�������</td>"
    Response.Write "    <td align=center width='30%'>" & CountNum & "</td>"
    Response.Write "    <td align=center width='20%'>����շ���</td>"
    Response.Write "    <td align=center width='30%'>" & DayMaxNum & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class=tdbg>"
    Response.Write "    <td align=center width='20%'>�������</td>"
    Response.Write "    <td align=center width='30%'>" & TotalView & "</td>"
    Response.Write "    <td align=center width='20%'>����շ�������</td>"
    Response.Write "    <td align=center width='30%'>" & DayMaxDate & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align=center width='20%'>ƽ���շ���</td>"
    Response.Write "    <td align=center width='30%'>" & AveDayNum & "</td>"
    Response.Write "    <td align=center width='20%'>���ʱ����</td>"
    Response.Write "    <td align=center width='30%'>" & HourMaxNum & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class=tdbg>"
    Response.Write "    <td align=center width='20%'>���շ�����</td>"
    Response.Write "    <td align=center width='30%'>" & DayNum & "</td>"
    Response.Write "    <td align=center width='20%'>���ʱ����ʱ��</td>"
    Response.Write "    <td align=center width='30%'>" & HourMaxTime & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class=tdbg>"
    Response.Write "    <td align=center width='20%'>Ԥ�ƽ��շ�����</td>"
    Response.Write "    <td align=center width='30%'>" & Int(DayNum * (24 * 60) / (Hour(Now) * 60 + Minute(Now))) & "</td>"
    Response.Write "    <td align=center width='20%'></td>"
    Response.Write "    <td align=center width='30%'></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr bgcolor='#39867B'>"
    Response.Write "    <td align=center width='20%' height='1'></td>"
    Response.Write "    <td align=center width='30%' height='1'></td>"
    Response.Write "    <td align=center width='20%' height='1'></td>"
    Response.Write "    <td align=center width='30%' height='1'></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class=tdbg>"
    Response.Write "    <td align=center width='20%'>���ڷ�������</td>"
    Response.Write "    <td align=center width='30%'>" & ChinaNum & "</td>"
    Response.Write "    <td align=center width='20%'>�����������</td>"
    Response.Write "    <td align=center width='30%'>" & OtherNum & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align=center width='20%'>���ò���ϵͳ</td>"
    If FoundInArr(RegCount_Fill, "FSystem", ",") = True Then
        Response.Write "    <td align=center width='30%'>" & MaxSys & " (" & MaxSysNum & ")</td>"
    Else
        Response.Write "    <td align=center width='30%'>�������</td>"
    End If
    Response.Write "    <td align=center width='20%'>���������</td>"
    If FoundInArr(RegCount_Fill, "FBrowser", ",") = True Then
        Response.Write "    <td align=center width='30%'>" & MaxBrw & " (" & MaxBrwNum & ")</td>"
    Else
        Response.Write "    <td align=center width='30%'>�������</td>"
    End If
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align=center width='20%'>�������ĵ�ַ</td>"
    If FoundInArr(RegCount_Fill, "FAddress", ",") = True Then
        Response.Write "    <td align=center width='30%'>" & MaxAre & " (" & MaxAreNum & ")</td>"
    Else
        Response.Write "    <td align=center width='30%'>�������</td>"
    End If
    Response.Write "    <td align=center width='20%'>����������վ</td>"
    If FoundInArr(RegCount_Fill, "FWeburl", ",") = True Then
        Response.Write "    <td align=center width='30%'>"
        If MaxWeb = "ֱ���������ǩ����" Then
            Response.Write "      " & Left(MaxWeb, 40) & " (" & MaxWebNum & ")"
        Else
            Response.Write "      <a href='" & MaxWeb & "' target='_blank'>" & Left(MaxWeb, 40) & "</a> (" & MaxWebNum & ")"
        End If
        Response.Write "    </td>"
    Else
        Response.Write "    <td align=center width='30%'>�������</td>"
    End If
    Response.Write "  </tr>"
    Response.Write "  <tr class=tdbg>"
    Response.Write "    <td align=center width='20%'>������Ļ�ֱ���</td>"
    If FoundInArr(RegCount_Fill, "FScreen", ",") = True Then
        Response.Write "    <td align=center width='30%'>" & MaxScr & " (" & MaxScrNum & ")</td>"
    Else
        Response.Write "    <td align=center width='30%'>�������</td>"
    End If
    Response.Write "    <td align=center width='20%'>������Ļ��ʾ��ɫ</td>"
    If FoundInArr(RegCount_Fill, "FColor", ",") = True Then
        Response.Write "    <td align=center width='30%'>" & MaxColor & " (" & MaxColorNum & ")</td>"
    Else
        Response.Write "    <td align=center width='30%'>�������</td>"
    End If
    Response.Write "  </tr>"
    Response.Write "  </tbody>"
    Response.Write "</table>"
    
    Set rs = Nothing
End Sub

Sub FVisitor()
    strGuide = "������ʼ�¼"
    sql = "Select * From PE_StatVisitor Order By VTime DESC"
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn_Counter, 1, 1
    If rs.BOF And rs.EOF Then
        Response.Write "<li>ϵͳ�������ݣ�"
    Else
        TotalPut = rs.RecordCount
        TitleRight = TitleRight & "�� <font color=red>" & TotalPut & "</font> �����ʼ�¼"
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > TotalPut Then
            If (TotalPut Mod MaxPerPage) = 0 Then
                CurrentPage = TotalPut \ MaxPerPage
            Else
                CurrentPage = TotalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < TotalPut Then
                rs.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        
        Dim VisitorNum
        VisitorNum = 0

        Response.Write "<table width='100%'><tr><td align='left'>�����ڵ�λ�ã���վͳ�ƹ���&nbsp;&gt;&gt;&nbsp;" & Search & strGuide & "</td><td align='right'>" & TitleRight & "</td></tr></table>"
        Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='0' class='border'>"
        Response.Write "  <tr class=title height='22'>"
        Response.Write "    <td align=center height='22'>����ʱ��(��������)</td>"
        Response.Write "    <td align=center height='22'>����ʱ��(�ͻ���)</td>"
        Response.Write "    <td align=center height='22'>������IP</td>"
        Response.Write "    <td align=center height='22'>��ַ</td>"
        Response.Write "    <td align=center height='22'>����ҳ��</td>"
        Response.Write "    <td align=center height='22'>����</td>"
        Response.Write "  </tr>"
        Do While Not rs.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td align=left width='120' height='22'>" & rs("VTime") & "</td>"
            Response.Write "    <td align=left width='120' height='22'>" & DateAdd("h", -rs("Timezone") - MasterTimeZone, rs("VTime")) & "</td>"
            Response.Write "    <td align=left width='80' height='22'>" & rs("IP") & "</td>"
            Response.Write "    <td align=left width='100' height='22'>" & rs("Address") & "</td>"
            Response.Write "    <td align=left height='22'><a href='" & rs("Referer") & "' title='" & rs("Referer") & "' target='_blank'>" & Left(rs("Referer"), 40) & "</a></td>"
            Response.Write "    <td align=left width='60' height='22'><a href='Admin_Counter.asp?Action=ClientDetail&id=" & rs("Id") & "'>�鿴��ϸ</a></td>"
            Response.Write "  </tr>"
            VisitorNum = VisitorNum + 1
            If VisitorNum >= MaxPerPage Then Exit Do
            rs.MoveNext
        Loop
        Response.Write "</table>"
        If TotalPut > 0 Then
            Response.Write ShowPage(strFileName, TotalPut, MaxPerPage, CurrentPage, True, True, "�����ʼ�¼", True)
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub FCounter()
    Item = Array("�״�", "����", "����", "�Ĵ�", "���", "����", "�ߴ�", "�˴�", "�Ŵ�", "ʮ������")
    ItemNum = 10
    strGuide = "���ʴ���ͳ�Ʒ���"
    StatItem = "��������"
    sql = "Select * From PE_StatVisit"
    Call Stable
End Sub

Sub StatYear()
    If Request("Type") = "" Then
       QYear = CStr(Year(Date))
    Else
       Search = "��ѯ�����"
    End If
    ItemNum = 12
    ReDim Item(11)
    For i = 0 To 11
      Item(i) = QYear & "��" & i + 1 & "��"
    Next
    strGuide = QYear & "�����ͳ�Ʒ���"
    StatItem = "�·�"
    sql = "Select * From PE_StatYear Where TYear='" & QYear & "'"

    Call Stable
End Sub

Sub StatAllYear()
    ItemNum = 12
    ReDim Item(ItemNum)
    For i = 0 To ItemNum - 1
      Item(i) = i + 1 & "��"
    Next

    strGuide = "ȫ�������ͳ�Ʒ���"
    StatItem = "�·�"
    sql = "Select * From PE_StatYear Where TYear='Total'"
    Call Stable
End Sub

Sub StatMonth()
    If Request("Type") = "" Then
       QMonth = CStr(Year(Date) & "-" & Month(Date))
    Else
       Search = "��ѯ�����"
    End If
    SYear = Mid(QMonth, 1, InStr(QMonth, "-") - 1)
    SMonth = Mid(QMonth, InStr(QMonth, "-") + 1)
    Select Case SMonth
    Case "2"
        If (SYear Mod 4) = 0 Then
           ItemNum = 29
        Else
           ItemNum = 28
        End If
    Case "4"
        ItemNum = 30
    Case "6"
        ItemNum = 30
    Case "9"
        ItemNum = 30
    Case "11"
        ItemNum = 30
    Case Else
        ItemNum = 31
    End Select
    ReDim Item(ItemNum - 1)
    For i = 0 To ItemNum - 1
      Item(i) = SYear & "��" & SMonth & "��" & i + 1 & "��"
    Next
    strGuide = QMonth & "�·���ͳ�Ʒ���"
    StatItem = "����"
    sql = "Select * From PE_StatMonth Where TMonth='" & QMonth & "'"
    Call Stable
End Sub

Sub StatAllMonth()
    ItemNum = 31
    ReDim Item(ItemNum)
    For i = 0 To ItemNum - 1
      Item(i) = i + 1 & "��"
    Next
    strGuide = "ȫ���·���ͳ�Ʒ���"
    StatItem = "����"
    sql = "Select * From PE_StatMonth Where TMonth='Total'"
    Call Stable
End Sub

Sub StatWeek()
    Item = Array("������", "����һ", "���ڶ�", "������", "������", "������", "������")
    ItemNum = 7
    strGuide = "���ܷ���ͳ�Ʒ���"
    StatItem = "����"
    sql = "Select * From PE_StatWeek Where Tweek='Current'"
    Call Stable
End Sub

Sub StatAllWeek()
    Item = Array("������", "����һ", "���ڶ�", "������", "������", "������", "������")
    ItemNum = 7
    strGuide = "ȫ���ܷ���ͳ�Ʒ���"
    StatItem = "����"
    sql = "Select * From PE_StatWeek Where Tweek='Total'"
    Call Stable
End Sub

Sub StatDay()
    If Request("Type") = "" Then
       QDay = CStr(Year(Date) & "-" & Month(Date) & "-" & Day(Date))
    Else
       Search = "��ѯ�����"
    End If
    ItemNum = 24
    ReDim Item(23)
    For i = 0 To ItemNum - 1
      Item(i) = Mid(i + 100, 2) & ":00-" & Mid(i + 101, 2) & ":00"
    Next
    strGuide = QDay & "�շ���ͳ�Ʒ���"
    StatItem = "Сʱ"
    sql = "Select * From PE_StatDay Where TDay='" & QDay & "'"
    Call Stable
End Sub

Sub StatAllDay()
    ItemNum = 24
    ReDim Item(ItemNum)
    For i = 0 To ItemNum - 1
      Item(i) = Mid(i + 100, 2) & ":00-" & Mid(i + 101, 2) & ":00"
    Next
    strGuide = "ȫ���շ���ͳ�Ʒ���"
    StatItem = "Сʱ"
    sql = "Select * From PE_StatDay Where TDay='Total'"
    Call Stable
End Sub

Sub FIP()
    sql = "Select * From PE_StatIp Order By TIpNum DESC"
    strGuide = "������IP��ַ����"
    StatItem = "IP��ַ"
    Call Ftable
End Sub

Sub FAddress()
    sql = "Select * From PE_StatAddress Order By TAddNum DESC"
    strGuide = "���������ڵ�ַ����"
    StatItem = "��ַ"
    Call Ftable
End Sub

Sub FTimezone()
    sql = "Select * From PE_StatTimezone Order By TtimNum DESC"
    strGuide = "����������ʱ������"
    StatItem = "ʱ��"
    Call Ftable
End Sub

Sub FWeburl()
    sql = "Select * From PE_StatWeburl Order By TWebNum DESC"
    strGuide = "������������վ����"
    StatItem = "������վ"
    Call Ftable
End Sub

Sub FKeyword()
    sql = "Select * From PE_StatKeyword Order By TKeywordNum DESC"
    strGuide = "�����������ؼ��ʷ���"
    StatItem = "�� �� ��"
    Call Ftable
End Sub

Sub FReferer()
    sql = "Select * From PE_StatRefer Order By TRefNum DESC"
    strGuide = "����������ҳ�����"
    StatItem = "����ҳ��"
    Call Ftable
End Sub

Sub FSystem()
    sql = "Select * From PE_StatSystem Order By TSysNum DESC"
    strGuide = "���������ò���ϵͳ����"
    StatItem = "����ϵͳ"
    Call Ftable
End Sub

Sub FBrowser()
    sql = "Select * From PE_StatBrowser Order By TBrwNum DESC"
    strGuide = "�������������������"
    StatItem = "�����"
    Call Ftable
End Sub

Sub FMozilla()
    sql = "Select * From PE_StatMozilla Order By TMozNum DESC"
    strGuide = "������HTTP_USER_AGENT�ַ�������"
    StatItem = "USER_AGENT"
    Call Ftable
End Sub

Sub FScreen()
    sql = "Select * From PE_StatScreen Order By TScrNum DESC"
    strGuide = "��������Ļ��С����"
    StatItem = "��Ļ��С"
    Call Ftable
End Sub

Sub FColor()
    sql = "Select * From PE_StatColor Order By TColNum DESC"
    strGuide = "��������Ļ��ʾ��ɫ����"
    StatItem = "��Ļ��ʾ��ɫ"
    Call Ftable
End Sub

Sub Stable()
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn_Counter, 1, 1
    If Not rs.BOF And Not rs.EOF Then
        Assay = rs.GetRows
        Rows = ItemNum - 1
    Else
        Rows = -1
    End If
    rs.Close
    Set rs = Nothing
    For i = 0 To Rows
        TotalNum = TotalNum + Assay(i, 0)
    Next
    If Rows >= 0 Then
        ReDim Percent(Rows)
        ReDim Barwidth(Rows)
    End If
    For i = 0 To Rows
        If TotalNum > 0 Then
            Percent(i) = FormatNumber(Int(Assay(i, 0) / TotalNum * 10000) / 100, 2, -1) & "%"
            Barwidth(i) = Assay(i, 0) / TotalNum * MaxWidth
        End If
    Next
    TitleRight = "��Чͳ�ƣ�<font color=red>" & TotalNum & "</font>"
    If Rows < 0 Then
        Response.Write "<li>ϵͳ�������ݣ�"
    Else
        Response.Write "<table width='100%'><tr><td align='left'>�����ڵ�λ�ã���վͳ�ƹ���&nbsp;&gt;&gt;&nbsp;" & Search & strGuide & "</td><td align='right'>" & TitleRight & "</td></tr></table>"
        Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
        Response.Write "  <tr class=title>"
        Response.Write "    <td align=left width='30%' nowrap height='22'>" & StatItem & "</td>"
        Response.Write "    <td align=left width='20%' nowrap>��������</td>"
        Response.Write "    <td align=left width='20%' nowrap>�ٷֱ�</td>"
        Response.Write "    <td align=left width='30%' nowrap>ͼʾ</td>"
        Response.Write "  </tr>"
        For i = 0 To Rows
            Response.Write "  <tr class='tdbg'>"
            Response.Write "    <td align=left>" & Item(i) & "</td>"
            Response.Write "    <td align=left>&nbsp;&nbsp;" & Assay(i, 0) & "</td>"
            Response.Write "    <td align=left>" & Percent(i) & "</td>"
            Response.Write "    <td align=left><img src='../Images/bar.gif' width='" & Barwidth(i) & "' height='10'></td>"
            Response.Write "  </tr>"
        Next
        Response.Write "</table>"
    End If
End Sub

Sub Ftable()
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn_Counter, 1, 1
    Do While Not rs.EOF
        TotalNum = TotalNum + rs(1)
        rs.MoveNext
    Loop
    rs.Close
    rs.Open sql, Conn_Counter, 1, 1
    If rs.BOF And rs.EOF Then
        Response.Write "<li>ϵͳ�������ݣ�"
    Else
        TotalPut = rs.RecordCount
        TitleRight = TitleRight & "��Чͳ�ƣ�<font color=red>" & TotalNum & "</font>"
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > TotalPut Then
            If (TotalPut Mod MaxPerPage) = 0 Then
                CurrentPage = TotalPut \ MaxPerPage
            Else
                CurrentPage = TotalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < TotalPut Then
                rs.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        
        Dim StatItemNum
        StatItemNum = 0
        Response.Write "<table width='100%'><tr><td align='left'>�����ڵ�λ�ã���վͳ�ƹ���&nbsp;&gt;&gt;&nbsp;" & Search & strGuide & "</td><td align='right'>" & TitleRight & "</td></tr></table>"
        Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
        Response.Write "  <tr class=title>"
        Response.Write "    <td align=left width='30%' nowrap height='22'>" & StatItem & "</td>"
        Response.Write "    <td align=left width='20%' nowrap>��������</td>"
        Response.Write "    <td align=left width='20%' nowrap>�ٷֱ�</td>"
        Response.Write "    <td align=left width='30%' nowrap>ͼʾ</td>"
        Response.Write "  </tr>"
        Do While Not rs.EOF
            Response.Write "  <tr class='tdbg'>"
            Response.Write "    <td align=left nowrap>"
            If (Action = "FWeburl" Or Action = "FReferer") And rs(0) <> "ֱ���������ǩ����" Then
                Response.Write "<a href='" & rs(0) & "' title='" & rs(0) & "' target='_blank'>" & Left(rs(0), 40) & "</a>"
            ElseIf Action = "FMozilla" Then
                Response.Write "<a title='" & rs(0) & "'>" & Left(rs(0), 40) & "</a>"
            Else
                Response.Write rs(0)
            End If
            Response.Write "    </td>"
            Response.Write "    <td align=left >&nbsp;&nbsp;" & rs(1) & "</td>"
            Response.Write "    <td align=left >" & FormatNumber(Int(rs(1) / TotalNum * 10000) / 100, 2, -1) & "%</td>"
            Response.Write "    <td align=left ><img src='../Images/bar.gif' width='" & rs(1) / TotalNum * MaxWidth & "' height='12'></td>"
            Response.Write "  </tr>"
            StatItemNum = StatItemNum + 1
            If StatItemNum >= MaxPerPage Then Exit Do
            rs.MoveNext
        Loop
        Response.Write "</table>"
        If TotalPut > 0 Then
            Response.Write ShowPage(strFileName, TotalPut, MaxPerPage, CurrentPage, True, True, "�����ʼ�¼", True)
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub FOnline()
    Dim OnNowTime
    OnNowTime = DateAdd("s", -OnlineTime, Now())
    strGuide = "��ǰ�����û�����"
    If CountDatabaseType = "SQL" Then
        sql = "select * from PE_StatOnline where LastTime>'" & OnNowTime & "' order by OnTime desc"
    Else
        sql = "select * from PE_StatOnline where LastTime>#" & OnNowTime & "# order by OnTime desc"
    End If

    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn_Counter, 1, 1
    If rs.BOF And rs.EOF Then
        Response.Write "<li>��ǰ�������ߣ�"
    Else
        TotalPut = rs.RecordCount
        TitleRight = TitleRight & "�� <font color=red>" & TotalPut & "</font> ���û�����"
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > TotalPut Then
            If (TotalPut Mod MaxPerPage) = 0 Then
                CurrentPage = TotalPut \ MaxPerPage
            Else
                CurrentPage = TotalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < TotalPut Then
                rs.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        
        Dim VisitorNum, LNowTime
        VisitorNum = 0

        Response.Write "<table width='100%'><tr><td align='left'>�����ڵ�λ�ã���վͳ�ƹ���&nbsp;&gt;&gt;&nbsp;" & Search & strGuide & "</td><td align='right'>" & TitleRight & "</td></tr></table>"
        Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
        Response.Write "  <tr class=title>"
        Response.Write "    <td align=center nowrap height='22'>���</td>"
        Response.Write "    <td align=center nowrap>������IP</td>"
        Response.Write "    <td align=center nowrap>��վʱ��</td>"
        Response.Write "    <td align=center nowrap>���ˢ��ʱ��</td>"
        Response.Write "    <td align=center nowrap>��ͣ��ʱ��</td>"
        Response.Write "    <td align=center nowrap>����ҳ�� �� �ͻ�����Ϣ</td>"
        Response.Write "  </tr>"
        
        Do While Not rs.EOF
            LNowTime = Cstrtime(CDate(Now() - rs("Ontime")))
            Response.Write "  <tr class='tdbg'>"
            Response.Write "    <td align=center width='8%' nowrap>" & VisitorNum & "</td>"
            Response.Write "    <td align=left width='15%' nowrap>" & rs("UserIP") & "</td>"
            Response.Write "    <td align=left width='17%' nowrap><a title=" & rs("OnTime") & ">" & TimeValue(rs("OnTime")) & "</a></td>"
            Response.Write "    <td align=left width='15%' nowrap>" & TimeValue(rs("LastTime")) & "</td>"
            Response.Write "    <td align=left width='15%' nowrap>" & LNowTime & "</td>"
            Response.Write "    <td align=left width='45%' nowrap title='����ҳ��: " & rs("UserPage") & vbCrLf & "�ͻ�����Ϣ: " & rs("UserAgent") & "'><a href=" & rs("UserPage") & " target=""_blank"">" & Left(Findpages(rs("UserPage")), 35) & "</a>"
            Response.Write "    </td>"
            Response.Write "  </tr>"
            VisitorNum = VisitorNum + 1
            If VisitorNum >= MaxPerPage Then Exit Do
            rs.MoveNext
        Loop
        Response.Write "</table>"
        If TotalPut > 0 Then
            Response.Write ShowPage(strFileName, TotalPut, MaxPerPage, CurrentPage, True, True, "�������û�", True)
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub

Function Findpages(furl)
    Dim Ffurl
    If furl <> "" Then
    Ffurl = Split(furl, "/")
    Findpages = Replace(furl, Ffurl(0) & "//" & Ffurl(2), "")
    If Findpages = "" Then Findpages = "/"
    Else
    Findpages = ""
    End If
End Function

Function Cstrtime(Lsttime)
    Dim Dminute, Dsecond
    Cstrtime = ""
    Dminute = 60 * Hour(Lsttime) + Minute(Lsttime)
    Dsecond = Second(Lsttime)
    If Dminute <> 0 Then Cstrtime = Dminute & "'"
    If Dsecond < 10 Then Cstrtime = Cstrtime & "0"
    Cstrtime = Cstrtime & Dsecond & """"
End Function

Sub HistoryList()
    Response.Write "<form name='form1' method='post' action='Admin_Counter.asp'>"
    Response.Write "  <table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='120'><strong>��վͳ�Ʋ�ѯ��</strong></td>"
    Response.Write "      <td>�������ͣ� "
    Response.Write "        <select name='type' size='1' class='Select' onChange=change_type()>"
    Response.Write "          <option value='1' selected>�ձ���</option>"
    Response.Write "          <option value='2'>�±���</option>"
    Response.Write "          <option value='3'>�걨��</option>"
    Response.Write "        </select>"
    Response.Write "        <select name='qyear' size='1' class='Select' onChange=change_it()>"
    For i = 2003 To 2010
        If i = Year(Date) Then
            Response.Write "<option value='" & i & "' selected>" & i & "</option>"
        Else
            Response.Write "<option value='" & i & "'>" & i & "</option>"
        End If
    Next
    Response.Write "        </select>"
    Response.Write "        ��"
    Response.Write "        <select name='qmonth' size='1' onChange=change_it()>"
    For i = 1 To 12
        If i = Month(Date) Then
            Response.Write "<option value='" & i & "' selected>" & i & "</option>"
        Else
            Response.Write "<option value='" & i & "'>" & i & "</option>"
        End If
    Next
    Response.Write "        </select>"
    Response.Write "        ��"
    Response.Write "        <select name='qday' size='1' >"
    Dim year29, monthdays
    year29 = Year(Date) Mod 4
    Select Case Month(Date)
    Case 2
        If year29 = 0 Then
            monthdays = 29
        Else
            monthdays = 28
        End If
    Case 4
        monthdays = 30
    Case 6
        monthdays = 30
    Case 9
        monthdays = 30
    Case 11
        monthdays = 30
    Case Else
        monthdays = 31
    End Select
    For i = 1 To monthdays
        If i = Day(Date) Then
            Response.Write "<option  value='" & i & "' selected>" & i & "</option>"
        Else
            Response.Write "<option  value='" & i & "'>" & i & "</option>"
        End If
    Next
    Response.Write "        </select>"
    Response.Write "        ��"
    Response.Write "        <input type='submit' name='Search' value='��ѯ'>"
    Response.Write "      </td>"
    Response.Write "      <td width='120' align='center'> </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub AutoAnalyse()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if(confirm('ȷʵҪ����ͳ�������Զ�������'))" & vbCrLf
    Response.Write "  {" & vbCrLf
    Response.Write "    return True;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else" & vbCrLf
    Response.Write "  {" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>" & vbCrLf
    Response.Write "  <tr class='title'>" & vbCrLf
    Response.Write "    <td height='22' align='center'><strong> ͳ �� �� �� �� �� �� �� </strong></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td height='150'>" & vbCrLf
    Response.Write "      <form name='myform' method='post' action='Admin_Counter.asp' onSubmit='return CheckForm();'>" & vbCrLf
    Response.Write "        <p align='center'>" & vbCrLf
    Response.Write "        <font color='#FF0000'><b>�����ô˹��ܣ���Ϊһ��������޷��ָ���</b></font>" & vbCrLf
    Response.Write "        <br>�˲�����������ݿ��н��ڷ��ʼ�¼������ݣ����ڽ��ڶ���վ�ķ���ͳ�����ݽ���ͳ�Ʒ���ʱʹ�á�" & vbCrLf
    Response.Write "        </p>" & vbCrLf
    Response.Write "        <p align='center'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='DoAutoAnalyse'>" & vbCrLf
    Response.Write "        <input type='submit' name='Submit' value=' ͳ�������Զ����� '>" & vbCrLf
    Response.Write "        </p>" & vbCrLf
    Response.Write "      </form>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub

Sub Init()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if(confirm('ȷʵҪ���г�ʼ����һ��������޷��ָ���'))" & vbCrLf
    Response.Write "  {" & vbCrLf
    Response.Write "    return True;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  else" & vbCrLf
    Response.Write "  {" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>" & vbCrLf
    Response.Write "  <tr class='title'>" & vbCrLf
    Response.Write "    <td height='22' align='center'><strong> �� �� �� ʼ �� </strong></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td height='150'>" & vbCrLf
    Response.Write "      <form name='myform' method='post' action='Admin_Counter.asp' onSubmit='return CheckForm();'>" & vbCrLf
    Response.Write "        <p align='center'>" & vbCrLf
    Response.Write "        <font color='#FF0000'><b>�����ô˹��ܣ���Ϊһ��������޷��ָ���</b></font>" & vbCrLf
    Response.Write "        <br>�˲�����������ݿ��е�����ͳ�����ݣ�����ϵͳ��ʼ��ʱ����Ҫ����վ�ķ���ͳ�����ݽ�������ͳ��ʱʹ�á�" & vbCrLf
    Response.Write "        </p>" & vbCrLf
    Response.Write "        <p align='center'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='DoInit'>" & vbCrLf
    Response.Write "        <input type='submit' name='Submit' value=' ͳ�����ݳ�ʼ�� '>" & vbCrLf
    Response.Write "        </p>" & vbCrLf
    Response.Write "      </form>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub

Sub ShowCompact()
    Response.Write "<form method='post' action='Admin_Counter.asp?action=CompactData'>"
    Response.Write "<table class='border' width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write " <tr class='title'>"
    Response.Write "     <td align='center' height='22' valign='middle'><b>ͳ�����ݿ�����ѹ��</b></td>"
    Response.Write " </tr>"
    Response.Write " <tr class='tdbg'>"
    Response.Write "     <td align='center' height='150' valign='middle'>"
    Response.Write "      <br>"
    Response.Write "      <br>"
    Response.Write "      ѹ��ǰ�������ȱ���ͳ�����ݿ⣬���ⷢ��������� <br>"
    Response.Write "      <br>"
    Response.Write "      <br>"
    Response.Write " <input name='submit' type=submit value=' ѹ��ͳ�����ݿ� '"
    If CountDatabaseType = "SQL" Then
        Response.Write " disabled"
    End If
    Response.Write "><br><br>"
    If ObjInstalled_FSO = False Then
        Response.Write "<b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
End Sub

Sub CompactData()
    Dim Engine, strDBPath, dbpath
    dbpath = Server.MapPath(db_counter)
    Call CloseConn_Counter
    strDBPath = Left(dbpath, InStrRev(dbpath, "\"))
    If fso.FileExists(dbpath) Then
        Set Engine = Server.CreateObject("JRO.JetEngine")
        Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
        fso.copyfile strDBPath & "temp.mdb", dbpath
        fso.DeleteFile (strDBPath & "temp.mdb")
        Set Engine = Nothing
        Call WriteSuccessMsg("ͳ�����ݿ�ѹ���ɹ���", ComeUrl)
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>ͳ�����ݿ�û���ҵ�!</li>"
    End If
End Sub

Sub DoInit()
    Conn_Counter.Execute ("delete from PE_StatAddress")

    Conn_Counter.Execute ("delete from PE_StatBrowser")
    Conn_Counter.Execute ("delete from PE_StatColor")
    Conn_Counter.Execute ("delete from PE_StatIp")
    Conn_Counter.Execute ("delete from PE_StatMozilla")
    Conn_Counter.Execute ("delete from PE_StatRefer")
    Conn_Counter.Execute ("delete from PE_StatScreen")
    Conn_Counter.Execute ("delete from PE_StatSystem")
    Conn_Counter.Execute ("delete from PE_StatTimezone")
    Conn_Counter.Execute ("delete from PE_StatVisit")
    Conn_Counter.Execute ("delete from PE_StatWeburl")
    Conn_Counter.Execute ("delete from PE_StatDay")
    Conn_Counter.Execute ("delete from PE_StatMonth")
    Conn_Counter.Execute ("delete from PE_StatWeek")
    Conn_Counter.Execute ("delete from PE_StatYear")
    Conn_Counter.Execute ("delete from PE_StatVisitor")
    Conn_Counter.Execute ("update PE_StatInfoList set StartDate='" & FormatDateTime(Date, 2) & "',OldDay='" & FormatDateTime(Date, 2) & "',TotalNum=0,TotalView=0,MonthNum=0,MonthMaxNum=0,OldMonth='',MonthMaxDate='',DayNum=0,DayMaxNum=0,DayMaxDate='',HourNum=0,HourMaxNum=0,OldHour='',HourMaxTime='',ChinaNum=0,OtherNum=0")
    Call WriteSuccessMsg("ͳ�����ݳ�ʼ���ɹ���", ComeUrl)
End Sub

Sub ClientDetail()
    Dim ClientNow
    Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from PE_StatVisitor where id=" & PE_CLng(Request("id"))
    rs.Open sql, Conn_Counter, 1, 1
    Response.Write "    <br><table width='100%' class='border' border='0' cellspacing='1' cellpadding='4' align='center'>" & vbCrLf
    Response.Write "      <tr class='title'> " & vbCrLf
    Response.Write "        <td colspan='2' class='title' align='center'><b>�� �� �� ¼ �� �� �� ʾ</b></td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'><b>����ʱ�䣨�Է�������ʱ���ǣ���</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("VTime") & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'><b>������IP��</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("IP") & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'><b>����������ʱ����</b></td>"
    Response.Write "        <td width='70%'>GMT" & rs("Timezone") & ":0</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'> " & vbCrLf
    Response.Write "          <b>���ڵ�ַ��</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("Address") & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'><b>����ʱ�䣨�Կͻ���ʱ���ǣ���</b></td>"
    Response.Write "        <td width='70%'>" & DateAdd("h", -rs("Timezone") - MasterTimeZone, rs("VTime")) & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'> " & vbCrLf
    Response.Write "          <b>����ҳ�棺</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("Referer") & "</td></tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'> " & vbCrLf
    Response.Write "          <b>����ϵͳ��</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("System") & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'> " & vbCrLf
    Response.Write "          <b>�������</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("Browser") & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'> " & vbCrLf
    Response.Write "          <b>��Ļ��С��</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("Screen") & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'> " & vbCrLf
    Response.Write "          <b>��Ļɫ�</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("Color") & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td colspan='2' nowrap> " & vbCrLf
    Response.Write "          <div align='center'>" & vbCrLf
    Response.Write "            <input type='button' name='Submit2' value='����' onClick=""window.location='Admin_Counter.asp?Action=FVisitor';"">" & vbCrLf
    Response.Write "          </div>" & vbCrLf
    Response.Write "        </td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    rs.Close
    Set rs = Nothing
End Sub

Sub IPAdd()
    Response.Write "    <form method='post' action='Admin_Counter.asp' name='myform'>" & vbCrLf
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' > " & vbCrLf
    Response.Write "         <tr class='title'>" & vbCrLf
    Response.Write "            <td height='22' colspan='2'> " & vbCrLf
    Response.Write "               <div align='center'><strong>ͳ��IP�����</strong></div>" & vbCrLf
    Response.Write "            </td>    " & vbCrLf
    Response.Write "         </tr>    " & vbCrLf
    Response.Write "               <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='350' class='tdbg5'><strong>��ʼ I P��</strong><br>ע�� ��ӵ�IP����������ݿ�����û�еļ�¼<br>��ֱ����ӣ���������ݿ����Ѿ����ڣ�����ʾ���Ƿ�����޸ġ�</td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='StartIP' type='text' id='StartIP' size='49' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "         </tr>    " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='350' class='tdbg5'><strong>��β I P��</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='EndIP' type='text' id='EndIP' size='49' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>  " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='350' class='tdbg5'><strong>��Դ��ϸ��ַ��</strong>" & vbCrLf
    Response.Write "               </td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='IPAddress' type='text' id='IPAddress' size='49' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>   " & vbCrLf
    Response.Write "        <tr class='tdbg'>     " & vbCrLf
    Response.Write "                     <td colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "                     " & vbCrLf
    Response.Write "                     <input name='Action' type='hidden' id='Action' value='SaveIPAdd'>        <input  type='submit' name='Submit' value=' �� �� '>&nbsp;&nbsp;        <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Counter.asp'"" style='cursor:hand;'>" & vbCrLf
    Response.Write "                     </td>    " & vbCrLf
    Response.Write "        </tr>  " & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "    </form>" & vbCrLf
    Call IPSearch
End Sub

Sub SaveIPAdd()
    Dim StartIP, EndIP, IPAddress
    If Request.Form("StartIP") = "" Or Not isIP(Request.Form("StartIP")) Then
        FoundErr = True
        ErrMsg = ErrMsg & "����д��ȷ��IP��ַ��"
        Exit Sub
    End If
    If Request.Form("EndIP") = "" Or Not isIP(Request.Form("StartIP")) Then
        FoundErr = True
        ErrMsg = ErrMsg & "����д��ȷ��IP��ַ��"
        Exit Sub
    End If
    If Request.Form("IPAddress") = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "����д��Դ��ϸ��ַ��"
        Exit Sub
    End If
    StartIP = EncodeIP(Trim(Request.Form("StartIP")))
    EndIP = EncodeIP(Trim(Request.Form("EndIP")))
    IPAddress = ReplaceBadChar(Trim(Request.Form("IPAddress")))
    Set rs = Server.CreateObject("adodb.recordset")
    sql = "select StartIP,EndIP,Address from PE_StatIpInfo where StartIP<=" & StartIP & " and EndIP>=" & EndIP & ""
    rs.Open sql, Conn_Counter, 1, 3
    If rs.EOF And rs.BOF Then
        rs.AddNew
        rs("StartIp") = StartIP
        rs("EndIP") = EndIP
        rs("Address") = IPAddress
        rs.Update
        Call WriteSuccessMsg("��վͳ��IP��ӳɹ���", ComeUrl)
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "���ʧ�ܣ������Ѵ��ڣ���������ip��ַ�������޸ġ�"
    End If
    rs.Close
End Sub


Sub IPManage()
    Response.Write "<table width='100%'><tr><td align='left'>�����ڵ�λ�ã�ͳ��IP����� &nbsp;&gt;&gt;&nbsp;IP ��ַ�����"
    Response.Write "</td><td align='right'>"
    Set rs = Server.CreateObject("adodb.recordset")

    Dim SearchIP, SearchAddress, Querysql,totalsql
    SearchAddress = ReplaceBadChar(Trim(Request("SearchAddress")))
    sql = "select top " & MaxPerPage & " StartIP,EndIP,Address from PE_StatIpInfo where 1=1"
    totalsql = totalsql& "select Count(*)  from PE_StatIpInfo where 1=1 "	
    If Request("SearchIP") <> "" Then
            SearchIP = EncodeIP(Trim(Request("SearchIP")))
            sql = sql & " And StartIp <=" & SearchIP & " and EndIp >=" & SearchIP
            totalsql = totalsql & " And StartIp <=" & SearchIP & " and EndIp >=" & SearchIP		
    Else
        If SearchAddress <> "" Then 
            sql = sql & " and Address like '%" & SearchAddress & "%'"
            totalsql = totalsql & " And StartIp <=" & SearchIP & " and EndIp >=" & SearchIP	
        End If	
    End If
    If CurrentPage > 1 Then
          Querysql = " and StartIP > (select max(StartIP) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " StartIP from PE_StatIpInfo where 1=1  "

            If Request("SearchIP") <> "" Then
                Querysql = Querysql & " And StartIp <=" & SearchIP & " and EndIp >=" & SearchIP		
            Else
                If SearchAddress <> "" Then 
                    Querysql = Querysql & " and Address like '%" & SearchAddress & "%'"
                End If			
            End If		  		  
            Querysql = Querysql & ") as Temp)"
    End If	
    totalPut = PE_CLng(Conn_Counter.Execute(totalsql)(0))
    sql = sql & Querysql
    rs.Open sql, Conn_Counter, 1, 1
    If rs.BOF And rs.EOF Then
        Response.Write "���ҵ� <font color=red>0</font> ��IP�μ�¼</td></tr></table>"
    Else
        'If SearchAddress="" and request("SearchIP")="" Then
            'Response.Write "����IP�μ�¼!</td></tr></table>"
        'Else
            Response.Write "�� <font color=red>" & TotalPut & "</font> ��IP�μ�¼</td></tr></table>"
        'End If
    End If
    Response.Write "    <table width='100%' class='border' border='0' cellspacing='1' cellpadding='0' align='center'>" & vbCrLf
    Response.Write "      <tr class='title'>" & vbCrLf
    Response.Write "        <td width='20%' height='22'> " & vbCrLf
    Response.Write "          <div align='center'><b>��ʼ IP</b></div>" & vbCrLf
    Response.Write "        </td>" & vbCrLf
    Response.Write "        <td width='20%' height='22'> " & vbCrLf
    Response.Write "          <div align='center'><b>��β IP</b></div>" & vbCrLf
    Response.Write "        </td>" & vbCrLf
    Response.Write "        <td width='42%' height='22'> " & vbCrLf
    Response.Write "          <div align='center'><b>��Դ��ϸ��ַ</b></div>" & vbCrLf
    Response.Write "        </td>" & vbCrLf
    Response.Write "        <td width='18%' nowrap> " & vbCrLf
    Response.Write "          <div align='center'><b>����</b></div>" & vbCrLf
    Response.Write "        </td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    If rs.RecordCount = 0 Then
        Response.Write "<tr class='tdbg'><td colspan='10' align='center'>û������������IP �μ�¼!</td><tr>" & vbCrLf
    Else
        Dim rsID, i, Sort
        i = 0
        Do While Not rs.EOF
            Response.Write "    <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            Response.Write "      <td width='20%' align='left' height='22'>" & DecodeIP(rs("StartIp")) & "</td>" & vbCrLf
            Response.Write "      <td width='20%' align='left' height='22'>" & DecodeIP(rs("EndIp")) & "</td>" & vbCrLf
            Response.Write "      <td width='42%' align='left' height='22'>" & rs("Address") & "</td>" & vbCrLf
            Response.Write "      <td width='18%' align='center' height='22'>" & vbCrLf
            Response.Write "        <a href='?action=editIP&StartIP=" & rs("StartIP") & "&EndIP=" & rs("EndIP") & "'>�༭</a>  |  <a href='?action=delIP&StartIP=" & rs("StartIP") & "&EndIP=" & rs("EndIP") & "'>ɾ��</a>" & vbCrLf
            Response.Write "    </td></tr>" & vbCrLf
            i = i + 1
            If i > MaxPerPage Then Exit Do
            rs.MoveNext
        Loop
    End If
    Response.Write "  </table>" & vbCrLf
    If TotalPut > 0 Then
        Response.Write ShowPage(strFileName, TotalPut, MaxPerPage, CurrentPage, True, True, "����¼", True)
    End If
    Call IPSearch
End Sub

Sub editIP()
    Dim StartIP, EndIP
    If Request("StartIP") = "" And Request("EndIP") = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "�����IP��ַ��"
        Exit Sub
    End If
    If Not (IsNumeric(Request("StartIP")) Or IsNumeric(Request("EndIP"))) Then
        FoundErr = True
        ErrMsg = ErrMsg & "����IP��ַ����"
        Exit Sub
    End If
    StartIP = Trim(Request("StartIP"))
    EndIP = Trim(Request("EndIP"))
    Set rs = Conn_Counter.Execute("select StartIP,EndIP,Address from PE_StatIpInfo where StartIP=" & StartIP & " and EndIP=" & EndIP & "")
    If rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "�������ݲ����ڸ�ip��ַ��"
    Else
        Response.Write "    <form method='post' action='Admin_Counter.asp' name='myform'>" & vbCrLf
        Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' > " & vbCrLf
        Response.Write "         <tr class='title'>" & vbCrLf
        Response.Write "            <td height='22' colspan='2'> " & vbCrLf
        Response.Write "               <div align='center'><strong>ͳ��IP���޸�</strong></div>" & vbCrLf
        Response.Write "            </td>    " & vbCrLf
        Response.Write "         </tr>    " & vbCrLf
        Response.Write "               <tr class='tdbg5'>      " & vbCrLf
        Response.Write "               <td width='350' class='tdbg5'><strong>��ʼ I P��</strong></td>      " & vbCrLf
        Response.Write "               <td class='tdbg'><input name='StartIP' type='text' id='StartIP' size='49' maxlength='30' value='" & DecodeIP(rs(0)) & "'>&nbsp;</td>    " & vbCrLf
        Response.Write "         </tr>    " & vbCrLf
        Response.Write "        <tr class='tdbg'>      " & vbCrLf
        Response.Write "               <td width='350' class='tdbg5'><strong>��β I P��</strong></td>      " & vbCrLf
        Response.Write "               <td class='tdbg'><input name='EndIP' type='text' id='EndIP' size='49' maxlength='30' value='" & DecodeIP(rs(1)) & "'>&nbsp;</td>    " & vbCrLf
        Response.Write "        </tr>  " & vbCrLf
        Response.Write "        <tr class='tdbg'>      " & vbCrLf
        Response.Write "               <td width='350' class='tdbg5'><strong>��Դ��ϸ��ַ��</strong></td>      " & vbCrLf
        Response.Write "               <td class='tdbg'><input name='IPAddress' type='text' id='IPAddress' size='49' maxlength='30' value='" & rs(2) & "'>&nbsp;</td>    " & vbCrLf
        Response.Write "        </tr>   " & vbCrLf
        Response.Write "        <tr class='tdbg'>     " & vbCrLf
        Response.Write "                     <td colspan='2' align='center' class='tdbg'>" & vbCrLf
        Response.Write "                     " & vbCrLf
        Response.Write "                     <input name='oldStartIP' type='hidden' id='oldStartIP' value='" & rs(0) & "'>" & vbCrLf
        Response.Write "                     <input name='oldEndIP' type='hidden' id='oldEndIP' value='" & rs(1) & "'>" & vbCrLf
        Response.Write "                     <input name='Action' type='hidden' id='Action' value='SaveIPedit'>        <input  type='submit' name='Submit' value=' �� �� '>&nbsp;&nbsp;        <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_Counter.asp'"" style='cursor:hand;'>" & vbCrLf
        Response.Write "                     </td>    " & vbCrLf
        Response.Write "        </tr>  " & vbCrLf
        Response.Write "      </table>" & vbCrLf
        Response.Write "    </form>" & vbCrLf
    End If
    Set rs = Nothing
    Call IPSearch
End Sub

Sub SaveIPedit()
    Dim StartIP, EndIP, IPAddress, oldStartIP, oldEndIP
    If Request.Form("StartIP") = "" Or Not isIP(Request.Form("StartIP")) Then
        FoundErr = True
        ErrMsg = ErrMsg & "����д��ȷ��ʼIP��ַ��"
        Exit Sub
    End If
    If Request.Form("EndIP") = "" Or Not isIP(Request.Form("EndIP")) Then
        FoundErr = True
        ErrMsg = ErrMsg & "����д��ȷ����IP��ַ��"
        Exit Sub
    End If
    If Request.Form("oldStartIP") = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "IP��ַ��ʧ��"
        Exit Sub
    End If
    If Request.Form("oldEndIP") = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "IP��ַ��ʧ��"
        Exit Sub
    End If
    If Request.Form("IPAddress") = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "����д��Դ��ϸ��ַ��"
        Exit Sub
    End If
    StartIP = EncodeIP(Trim(Request.Form("StartIP")))
    EndIP = EncodeIP(Trim(Request.Form("EndIP")))
    oldStartIP = Trim(Request.Form("oldStartIP"))
    oldEndIP = Trim(Request.Form("oldEndIP"))
    IPAddress = ReplaceBadChar(Trim(Request.Form("IPAddress")))
    Dim RowCount
    Conn_Counter.Execute ("update PE_StatIpInfo set StartIP=" & StartIP & ",EndIP=" & EndIP & ",Address='" & IPAddress & "' where StartIP=" & oldStartIP & " and EndIP=" & oldEndIP & ""), RowCount
    If RowCount = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "IP �޸�ʧ�ܣ���������ip��ַ���½����޸ġ�"
    Else
        Call WriteSuccessMsg("��վͳ��IP�޸ĳɹ���", "admin_counter.asp?Action=IPManage")
    End If
End Sub

Sub delIP()
    Dim StartIP, EndIP
    If Request("StartIP") = "" And Request("EndIP") = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "�����IP��ַ��"
        Exit Sub
    End If
    If Not (IsNumeric(Request("StartIP")) Or IsNumeric(Request("EndIP"))) Then
        FoundErr = True
        ErrMsg = ErrMsg & "����IP��ַ����"
        Exit Sub
    End If
    StartIP = Trim(Request("StartIP"))
    EndIP = Trim(Request("EndIP"))
    Dim RowCount
    Conn_Counter.Execute ("delete from PE_StatIpInfo where StartIP=" & StartIP & " and EndIP=" & EndIP & ""), RowCount
    If RowCount = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "IP ɾ��ʧ�ܣ���������ip��ַ����ɾ����"
    Else
        Call WriteSuccessMsg("��վͳ��IPɾ���ɹ���", ComeUrl)
    End If
End Sub

Sub IPSearch()
    Response.Write "    <form method='post' action='Admin_Counter.asp' name='myform'>" & vbCrLf
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' > " & vbCrLf

    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "            <td width='120'><strong>ͳ��I P��������</strong></td>"
    Response.Write "            <td>I P ��ַ��</td>" & vbCrLf
    Response.Write "            <td><input name='SearchIP' type='text' id='SearchIP' size='20' maxlength='20'>&nbsp;</td>" & vbCrLf
    Response.Write "            <td>��Դ��ϸ��ַ��</td>" & vbCrLf
    Response.Write "            <td><input name='SearchAddress' type='text' id='SearchAddress' size='20' maxlength='30'>&nbsp;</td>" & vbCrLf
    Response.Write "            <td>" & vbCrLf
    Response.Write "                     <input name='Action' type='hidden' id='Action' value='SearchIP'>        <input  type='submit' name='Submit' value=' �� �� '>" & vbCrLf
    Response.Write "            </td>    " & vbCrLf
    Response.Write "      <td width='110'> </td>"
    Response.Write "        </tr>  " & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "    </form>" & vbCrLf
End Sub

Sub Import()
    Response.Write "<form name='myform' method='post' action='Admin_Counter.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>IP���ݿ⵼��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;������Ҫ�����ģ�����ݿ���ļ����� "
    Response.Write "        <input name='TemplateMdb' type='text' id='TemplateMdb' value='../Count/IP.mdb' size='20' maxlength='50'>"
    Response.Write "        <input align=""center"" name='Submit' type='submit' id='Submit' value=' ��һ�� '>"
    Response.Write "        <br><font color=""#FF0000"">&nbsp;&nbsp;&nbsp;&nbsp;ע�⣺�����IP���ݻ�ֱ�Ӹ���ԭ����IP���ݣ������ñ��ݹ�����</Font><br>"		
    Response.Write "        <input name='Action' type='hidden' id='Action' value='DoImport'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub


Sub DoImport()
    On Error Resume Next
    Dim mdbname, tconn, trs, iCount
    
    '��õ���ģ�����ݿ�·��
    mdbname = Replace(Trim(Request.Form("Templatemdb")), "'", "")

    If mdbname = "" Then
        mdbname = Replace(Trim(Request.QueryString("Templatemdb")), "'", "")
    End If

    mdbname = Replace(mdbname, "��", "/") '��ֹ�ⲿ���Ӱ�ȫ����

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д����ģ�����ݿ���"
        Exit Sub
    End If

    '��������ģ�����ݿ�
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݿ����ʧ�ܣ����Ժ����ԣ�����ԭ��" & Err.Description
        Err.Clear
        Exit Sub
    End If 
	

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    tconn.Execute ("select * from PE_StatIpInfo")

    If Err Then
        Set trs = Nothing
        ErrMsg = ErrMsg & "<li>��Ҫ��������ݿ�,����ϵͳ�������ݿ�,��ʹ��ϵͳ�������ݿ⡣"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    		
    Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    Response.Write "  <tr align='center' class='title'><td height='22'><strong>���ݵ����У����Ժ�</strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='100' valign='top'><br><Div name=""ShowMess"" id=""ShowMess"" >���ڳ�ʼ�����ݿ�</Div></td></tr>" & vbCrLf
    Response.Write "  <tr align='center' class='tdbg'><td>"
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf	

    Conn_Counter.Execute ("delete from PE_StatIpInfo")
    Response.Flush()	
	
    Set trs = Server.CreateObject("ADODB.Recordset")
    trs.Open "select * from PE_StatIpInfo", tconn, 1, 1
	
    sql = "select * from PE_StatIpInfo"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, Conn_Counter, 1, 3
    Dim countIPNum
    countIPNum = 1
    Do While Not trs.EOF
        rs.addnew
        rs("StartIP") = trs("StartIP")
        rs("EndIP") = trs("EndIP")
        rs("Address") = trs("Address")
        rs.Update
        trs.MoveNext
        countIPNum = countIPNum + 1
		
        If (countIPNum mod 10000) = 0  then 
            Response.Write "<script>" & vbCrLf
            Response.Write "document.getElementById(""ShowMess"").innerHTML=""����ת����,���Ժ�<br>�Ѿ��ɹ�����"& countIPNum &"��IP����"";" & vbCrLf
            Response.Write "</script>" & vbCrLf			
            response.Flush()
        End If	
    Loop
    Response.Write "<script>" & vbCrLf
    Response.Write "document.getElementById(""ShowMess"").innerHTML=""�ɹ���������IP����"";" & vbCrLf
    Response.Write "</script>" & vbCrLf		
    Response.Flush()	
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
	
    tconn.Close
    Set tconn = Nothing
   ' Call WriteSuccessMsg("�Ѿ��ɹ���IP���ݿ⵼�룡", ComeUrl)
End Sub

Sub Export()
    Response.Write "<form name='myform' method='post' action='Admin_Counter.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>IP���ݿ⵼��</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;������Ҫ������ģ�����ݿ���ļ����� "
    Response.Write "        <input name='TemplateMdb' type='text' id='TemplateMdb' value='../Count/IP.mdb' size='20' maxlength='50'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' ��һ�� '>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='DoExport'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub


Sub DoExport()
    On Error Resume Next
    Dim mdbname, tconn, trs, iCount
    
    mdbname = Replace(Trim(Request.Form("Templatemdb")), "'", "")

    If mdbname = "" Then
        mdbname = Replace(Trim(Request.QueryString("Templatemdb")), "'", "")
    End If

    mdbname = Replace(mdbname, "��", "/") '��ֹ�ⲿ���Ӱ�ȫ����

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����д����ģ�����ݿ���"
        Exit Sub
    End If
    '��������IP���ݿ�
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>���ݿ����ʧ�ܣ����Ժ����ԣ�����ԭ��" & Err.Description
        Err.Clear
        Exit Sub
    End If 
	

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    tconn.Execute ("select * from PE_StatIpInfo")

    If Err Then
        Set trs = Nothing
        ErrMsg = ErrMsg & "<li>��Ҫ��������ݿ�,����ϵͳ�������ݿ�,��ʹ��ϵͳ�������ݿ⡣"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If
		
	
    tconn.Execute ("delete from PE_StatIpInfo")
    Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    Response.Write "  <tr align='center' class='title'><td height='22'><strong>���ݵ����У����Ժ����ݵ���ʱ����ˢ��ҳ��</strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='100' valign='top'><br><Div name=""ShowMess"" id=""ShowMess"" >���ڳ�ʼ��ip���ݿ⣬���Ժ�</Div></td></tr>" & vbCrLf
    Response.Write "  <tr align='center' class='tdbg'><td>"
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf		

    tconn.Execute ("delete from PE_StatIpInfo")
    Response.Flush()	
		
    Set trs = Server.CreateObject("ADODB.Recordset")
    trs.Open "select * from PE_StatIpInfo", tconn, 1, 3
	
    sql = "select * from PE_StatIpInfo"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, Conn_Counter, 1, 3
    Dim countIPNum
    countIPNum = 1
    Do While Not rs.EOF
        trs.addnew
        trs("StartIP") = rs("StartIP")
        trs("EndIP") = rs("EndIP")
        trs("Address") = rs("Address")
        trs.Update
        rs.MoveNext
        countIPNum = countIPNum + 1		
        If (countIPNum mod 10000) = 0  then 
            Response.Write "<script>" & vbCrLf
            Response.Write "document.getElementById(""ShowMess"").innerHTML=""����ת����,���Ժ�<br>�Ѿ��ɹ�����"& countIPNum &"��IP����"";" & vbCrLf
            Response.Write "</script>" & vbCrLf			
            response.Flush()
        End If			
    Loop
    Response.Write "<script>" & vbCrLf
    Response.Write "document.getElementById(""ShowMess"").innerHTML=""�ɹ���������IP����"";" & vbCrLf
    Response.Write "</script>" & vbCrLf		
    Response.Flush()
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
	
    tconn.Close
    Set tconn = Nothing
    'Call WriteSuccessMsg("�Ѿ��ɹ���IP���ݿ⵼����", ComeUrl)
End Sub



%>
