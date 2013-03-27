<!--#include file="Admin_Common.asp"-->
<!--#include file="../count/conn_counter.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = "Counter"   '其他权限

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

MaxWidth = 220      '放置统计条的表格的宽度
TotalNum = 0

'打开数据库
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
Response.Write "<title>网站统计管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "<script>" & vbCrLf
Response.Write "function change_type()" & vbCrLf
Response.Write "{ " & vbCrLf
Response.Write "    select_type=form1.type.options[form1.type.selectedIndex].text;" & vbCrLf
Response.Write "    switch(select_type)" & vbCrLf
Response.Write "    { " & vbCrLf
Response.Write "        case '日报表' :form1.qmonth.disabled=0;form1.qday.disabled=0;break;" & vbCrLf
Response.Write "        case '月报表' :form1.qmonth.disabled=0;form1.qday.disabled=1;break;" & vbCrLf
Response.Write "        case '年报表' :form1.qmonth.disabled=1;form1.qday.disabled=1;break;" & vbCrLf
Response.Write "    } " & vbCrLf
Response.Write "} " & vbCrLf
Response.Write "function change_it()" & vbCrLf
Response.Write "{ " & vbCrLf
Response.Write "    select_type=form1.type.options[form1.type.selectedIndex].text;" & vbCrLf
Response.Write "    if (select_type=='日报表')" & vbCrLf
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
    Call ShowPageTitle("网 站 统 计 配 置", 10025)
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='70' height='30'><strong>管理导航：</strong></td>" & vbCrLf
    Response.Write "    <td height='30'><a href='Admin_Counter.asp?Action=ShowConfig'>网站统计信息配置</a> | <a href='Admin_Counter.asp?Action=IPAdd'>统计IP库添加</a> | <a href='Admin_Counter.asp?Action=IPManage'>统计IP库管理</a> | <a href='Admin_Counter.asp?Action=Compact'>压缩统计数据库</a> | <a href='Admin_Counter.asp?Action=Init'>统计数据初始化</a>  | <a href='Admin_Counter.asp?Action=Export'>导出IP数据库</a> | <a href='Admin_Counter.asp?Action=Import'>导入IP数据库</a>    </td>" & vbCrLf
    Response.Write "  </tr>"
Else
    Call ShowPageTitle("网 站 统 计 管 理", 10025)
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td height='30'>"
    Response.Write "    <a href='Admin_Counter.asp?Action=Infolist'>综合统计</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=FVisitor'>访问记录</a>&nbsp;|"
    If FoundInArr(RegCount_Fill, "FVisit", ",") = True Then
    Response.Write "    <a href='Admin_Counter.asp?Action=FCounter'>访问次数</a>&nbsp;|"
    End If
    Response.Write "    <a href='Admin_Counter.asp?Action=StatYear'>年 报 表</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=StatAllYear'>全 部 年</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=StatMonth'>月 报 表</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=StatAllMonth'>全 部 月</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=StatWeek'>周 报 表</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=StatAllWeek'>全 部 周</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=StatDay'>日 报 表</a>&nbsp;|"
    Response.Write "    <a href='Admin_Counter.asp?Action=StatAllDay'>全 部 日</a>&nbsp;|<br>"
    If FoundInArr(RegCount_Fill, "IsCountOnline", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FOnline'>在线用户</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FIP", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FIp'>IP 地 址</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FAddress", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FAddress'>地址分析</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FTimezone", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FTimezone'>时区分析</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FKeyword", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FKeyword'>关 键 词</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FWeburl", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FWeburl'>来访网站</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FRefer", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FReferer'>链接页面</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FSystem", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FSystem'>操作系统</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FBrowser", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FBrowser'>浏 览 器</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FMozilla", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FMozilla'>字串分析</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FScreen", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FScreen'>屏幕大小</a>&nbsp;|"
    End If
    If FoundInArr(RegCount_Fill, "FColor", ",") = True Then
        Response.Write "    <a href='Admin_Counter.asp?Action=FColor'>屏幕色深</a>&nbsp;|"
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
        Response.Write "<li>网站统计配置数据丢失！"
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
    Response.Write "  if(confirm('强烈建议尽量选择少的统计功能项目，最好一个都不启用！！！'))" & vbCrLf
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
    Response.Write "<table width='100%'><tr><td align='left'>您现在的位置：网站统计信息配置</td></tr></table>"
    Response.Write "<form method='POST' action='Admin_Counter.asp?Action=SaveConfig' id='form1' name='form1' onsubmit='return ConfirmModify();'>" & vbCrLf

    Response.Write "<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr align='center'>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title6' onclick='ShowTabs(0)'>基本信息</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(1)'>初始化设置</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(2)'>功能项目</td>" & vbCrLf
    Response.Write "    <td id='TabTitle' class='title5' onclick='ShowTabs(3)'>调用代码</td>" & vbCrLf

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
    'Response.Write "      <td height='22' colspan='4'> <a name='SiteCountInfo'></a><strong>网站统计信息配置</strong></td>" & vbCrLf
    'Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong> 服务器所在时区：</strong></td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='MasterTimeZone' type='text' id='MasterTimeZone' value='" & MasterTimeZone & "' size='20' maxlength='50'>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong>在线用户的保留时间：</strong><br>" & vbCrLf
    Response.Write "      用户切换页面至其他网站或者关闭浏览器后，在线名单将在上述时间内删除该用户。这个间隔越小，网站统计的当前时刻在线名单越准确；这个间隔越大，网站统计的在线人数越多。" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='OnlineTime' type='text' id='OnlineTime' value='" & OnlineTime & "' size='20' maxlength='50'>" & vbCrLf
    Response.Write "      秒      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong>自动标记在线间隔：</strong><br>" & vbCrLf
    Response.Write "      客户端浏览器会每隔上述时间向服务器提交一次在线信息，同时服务器将其标记为在线，这个间隔越小，服务器需要处理的请求越多。</td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='Interval' type='text' id='Interval' value='60' size='20' maxlength='50' disabled>" & vbCrLf
    Response.Write "        秒" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong>自动标记在线间隔循环次数：</strong><br>" & vbCrLf
    Response.Write "      此是为了防止用户打开网页，但长时间无任何活动而设置。客户端浏览器向服务器提交在线信息次数超过此次数，立即停止提交。" & vbCrLf
    Response.Write "</td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='IntervalNum' type='text' id='IntervalNum' value='" & IntervalNum & "' size='20' maxlength='50'>" & vbCrLf
    Response.Write "        次" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong>保留访问记录数：</strong><br>" & vbCrLf
    Response.Write " 保存访问明细(最后访问)条目数。</td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='VisitRecord' type='text' id='VisitRecord' value='" & VisitRecord & "' size='20' maxlength='50'>" & vbCrLf
    Response.Write "        条" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong> 保留访问IP数(大于20小于800的数字)： </strong><br>" & vbCrLf
    Response.Write "      当不启用“在线人数统计”功能时，系统将以保留访问者IP的方式来防止刷新，即同一个IP访问多次或者在网站内切换页面，均只计算浏览量而不计算访问量。    </td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='KillRefresh' type='text' id='KillRefresh' value='" & KillRefresh & "' size='20' maxlength='50'>" & vbCrLf
    Response.Write "        个" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf

    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong>使用本系统前的访问量：</strong>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='OldTotalNum' type='text' id='OldTotalNum' value='" & OldTotalNum & "' size='20' maxlength='9'>" & vbCrLf
    Response.Write "        人次" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "        <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong> 使用本系统前的浏览量：</strong>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "        <input name='OldTotalView' type='text' id='OldTotalView' value='" & OldTotalView & "' size='20' maxlength='9'>" & vbCrLf
    Response.Write "        人次" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf

    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong>功能项目:</strong><br>" & vbCrLf
    Response.Write "      统计太多的项目会减慢访问速度，耗费太多网站资源，一段时间不想分析的功能项目建议不要起用！<br><font color='red'>强烈建议尽量选择少的功能项目，最好一个都不启用！！！</font><br>" & vbCrLf
    Response.Write "      </td><td>" & vbCrLf
    Response.Write "        <table width='100%'><tr>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "      <input name='RegFields_Fill' type='checkbox' value='IsCountOnline'" & vbCrLf
    If FoundInArr(RegCount_Fill, "IsCountOnline", ",") = True Then Response.Write " checked"
    Response.Write "      >启用“在线人数统计”功能</td><td><input name='RegFields_Fill' type='checkbox' value='FIP'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FIP", ",") = True Then Response.Write " checked"
    Response.Write "      >客户端IP地址分析</td><td><input name='RegFields_Fill' type='checkbox' value='FAddress'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FAddress", ",") = True Then Response.Write " checked"
    Response.Write "      >客户端地址分析 </td></tr><tr class='tdbg'><td><input name='RegFields_Fill' type='checkbox' value='FRefer'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FRefer", ",") = True Then Response.Write " checked"
    Response.Write "      >客户端链接页面分析 </td><td><input name='RegFields_Fill' type='checkbox' value='FTimezone'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FTimezone", ",") = True Then Response.Write " checked"
    Response.Write "      >客户端时区分析 </td><td><input name='RegFields_Fill' type='checkbox' value='FWeburl'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FWeburl", ",") = True Then Response.Write " checked"
    Response.Write "      >客户端来访网站分析  </td></tr><tr class='tdbg'><td><input name='RegFields_Fill' type='checkbox' value='FBrowser'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FBrowser", ",") = True Then Response.Write " checked"
    Response.Write "      >客户端浏览器分析 </td><td><input name='RegFields_Fill' type='checkbox' value='FMozilla'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FMozilla", ",") = True Then Response.Write " checked"
    Response.Write "      >客户端字串分析 </td><td><input name='RegFields_Fill' type='checkbox' value='FSystem'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FSystem", ",") = True Then Response.Write " checked"
    Response.Write "      >客户端操作系统分析 </td></tr> <tr class='tdbg'><td><input name='RegFields_Fill' type='checkbox' value='FScreen'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FScreen", ",") = True Then Response.Write " checked"
    Response.Write "      >客户端屏幕大小分析 </td><td><input name='RegFields_Fill' type='checkbox' value='FColor'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FColor", ",") = True Then Response.Write " checked"
    Response.Write "      >客户端屏幕色彩分析  </td><td><input name='RegFields_Fill' type='checkbox' value='FKeyword'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FKeyword", ",") = True Then Response.Write " checked"
    Response.Write "      >搜索关键词分析 </td></tr> <tr class='tdbg'><td><input name='RegFields_Fill' type='checkbox' value='FVisit'"
    If FoundInArr(RegCount_Fill, "FVisit", ",") = True Then Response.Write " checked"
    Response.Write "      >访问次数统计分析 </td><td><input name='RegFields_Fill' type='checkbox' value='FYesterDay'" & vbCrLf
    If FoundInArr(RegCount_Fill, "FYesterDay", ",") = True Then Response.Write " checked"
    Response.Write "      >启用昨日统计  </td><td>" & vbCrLf

    Response.Write "      </td></tr>" & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "   </td></tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong>统计计数代码类型：</strong><br>" & vbCrLf
    Response.Write "      [请先选择您想要的输出信息类型]</td>" & vbCrLf
    Response.Write "      <td colspan='3'> " & vbCrLf
    Response.Write "      <select name='select'  onChange='setFileFileds(this.value)'>" & vbCrLf
    Response.Write "        <option value='1' selected>显示简单样式信息</option>" & vbCrLf
    Response.Write "        <option value='2'>显示普通样式信息</option>" & vbCrLf
    Response.Write "        <option value='3'>显示复杂样式信息</option>" & vbCrLf
    Response.Write "        <option value='4'>统计但不显示信息</option>" & vbCrLf
    Response.Write "      </select>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong> 显示数据代码：</strong><br>" & vbCrLf
    Response.Write "      请将此代码拷贝到您需要做统计的页面，此代码不仅用于向放置了此代码的页面输出统计数据，而且还对该页面计数。<br></td>" & vbCrLf
    Response.Write "      <td colspan='3'><textarea name='selectKey' cols='50' rows='5' id='selectKey'><script src='{$InstallDir}Count/CounterLink.asp?style=simple'></script></textarea>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='300' height='25' class='tdbg5'><strong> 前台显示在线链接代码：</strong><br>" & vbCrLf
    Response.Write "      请将此代码拷贝到您需要显示在线列表链接的模板中，此代码仅用于向放置了此代码的页面显示在线列表链接，而不对该页面计数。<br></td>" & vbCrLf
    Response.Write "      <td colspan='3'><textarea name='LinkContent' cols='50' rows='5' id='LinkContent'><a href='{$InstallDir}Count/ShowOnline.asp' target='_blank'>网站在线情况详细列表</a></textarea>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "        </tbody>" & vbCrLf
    Response.Write "        </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    
    Response.Write "  <p align='center'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveConfig'>" & vbCrLf
    Response.Write "        <input name='cmdSave' type='submit' id='cmdSave' value=' 保存设置 '>" & vbCrLf
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
    Call WriteSuccessMsg("网站统计配置保存成功！", ComeUrl)
End Sub


Sub Infolist()
    Dim StartDate, StatDayNum, AllNum, TotalView, CountNum, AveDayNum, DayNum
    Dim MonthMaxNum, MonthMaxDate, DayMaxNum, DayMaxDate, HourMaxNum, HourMaxTime, ZoneNum, ChinaNum, OtherNum
    Dim MaxBrw, MaxBrwNum, MaxSys, MaxSysNum, MaxScr, MaxScrNum, MaxAre, MaxAreNum, MaxWeb, MaxWebNum, MaxColor, MaxColorNum
    strGuide = "网站综合统计信息"
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
    TitleRight = "开始统计日期：<font color=blue>" & StartDate & "</font>"

    Response.Write "<table width='100%'><tr><td align='left'>您现在的位置：网站统计管理&nbsp;&gt;&gt;&nbsp;" & Search & strGuide & "</td><td align='right'>" & TitleRight & "</td></tr></table>"
    Response.Write "<table border=0 cellpadding=2 cellspacing=1 width='100%' bgcolor='#FFFFFF' class='border'>"
    Response.Write "  <tr class='title' align='center'>"
    Response.Write "    <td align=center width='20%' height='22'>统计项</td>"
    Response.Write "    <td align=center width='30%'>统计数据</td>"
    Response.Write "    <td width='20%'>统计项</td>"
    Response.Write "    <td align='center' width='30%'>统计数据</td>"
    Response.Write "  </tr>"
    Response.Write "  <tbody>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align=center width='20%'>总统计天数</td>"
    Response.Write "    <td align=center width='30%'>" & StatDayNum & "</td>"
    Response.Write "    <td align=center width='20%'>最高月访量</td>"
    Response.Write "    <td align=center width='30%'>" & MonthMaxNum & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class=tdbg>"
    Response.Write "    <td align=center width='20%'>总访问量</td>"
    Response.Write "    <td align=center width='30%'>" & AllNum & "</td>"
    Response.Write "    <td align=center width='20%'>最高月访量月份</td>"
    Response.Write "    <td align=center width='30%'>" & MonthMaxDate & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align=center width='20%'>总访问人数</td>"
    Response.Write "    <td align=center width='30%'>" & CountNum & "</td>"
    Response.Write "    <td align=center width='20%'>最高日访量</td>"
    Response.Write "    <td align=center width='30%'>" & DayMaxNum & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class=tdbg>"
    Response.Write "    <td align=center width='20%'>总浏览量</td>"
    Response.Write "    <td align=center width='30%'>" & TotalView & "</td>"
    Response.Write "    <td align=center width='20%'>最高日访量日期</td>"
    Response.Write "    <td align=center width='30%'>" & DayMaxDate & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align=center width='20%'>平均日访量</td>"
    Response.Write "    <td align=center width='30%'>" & AveDayNum & "</td>"
    Response.Write "    <td align=center width='20%'>最高时访量</td>"
    Response.Write "    <td align=center width='30%'>" & HourMaxNum & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class=tdbg>"
    Response.Write "    <td align=center width='20%'>今日访问量</td>"
    Response.Write "    <td align=center width='30%'>" & DayNum & "</td>"
    Response.Write "    <td align=center width='20%'>最高时访量时间</td>"
    Response.Write "    <td align=center width='30%'>" & HourMaxTime & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class=tdbg>"
    Response.Write "    <td align=center width='20%'>预计今日访问量</td>"
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
    Response.Write "    <td align=center width='20%'>国内访问人数</td>"
    Response.Write "    <td align=center width='30%'>" & ChinaNum & "</td>"
    Response.Write "    <td align=center width='20%'>国外访问人数</td>"
    Response.Write "    <td align=center width='30%'>" & OtherNum & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align=center width='20%'>常用操作系统</td>"
    If FoundInArr(RegCount_Fill, "FSystem", ",") = True Then
        Response.Write "    <td align=center width='30%'>" & MaxSys & " (" & MaxSysNum & ")</td>"
    Else
        Response.Write "    <td align=center width='30%'>无须分析</td>"
    End If
    Response.Write "    <td align=center width='20%'>常用浏览器</td>"
    If FoundInArr(RegCount_Fill, "FBrowser", ",") = True Then
        Response.Write "    <td align=center width='30%'>" & MaxBrw & " (" & MaxBrwNum & ")</td>"
    Else
        Response.Write "    <td align=center width='30%'>无须分析</td>"
    End If
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align=center width='20%'>访问最多的地址</td>"
    If FoundInArr(RegCount_Fill, "FAddress", ",") = True Then
        Response.Write "    <td align=center width='30%'>" & MaxAre & " (" & MaxAreNum & ")</td>"
    Else
        Response.Write "    <td align=center width='30%'>无须分析</td>"
    End If
    Response.Write "    <td align=center width='20%'>访问最多的网站</td>"
    If FoundInArr(RegCount_Fill, "FWeburl", ",") = True Then
        Response.Write "    <td align=center width='30%'>"
        If MaxWeb = "直接输入或书签导入" Then
            Response.Write "      " & Left(MaxWeb, 40) & " (" & MaxWebNum & ")"
        Else
            Response.Write "      <a href='" & MaxWeb & "' target='_blank'>" & Left(MaxWeb, 40) & "</a> (" & MaxWebNum & ")"
        End If
        Response.Write "    </td>"
    Else
        Response.Write "    <td align=center width='30%'>无须分析</td>"
    End If
    Response.Write "  </tr>"
    Response.Write "  <tr class=tdbg>"
    Response.Write "    <td align=center width='20%'>常用屏幕分辨率</td>"
    If FoundInArr(RegCount_Fill, "FScreen", ",") = True Then
        Response.Write "    <td align=center width='30%'>" & MaxScr & " (" & MaxScrNum & ")</td>"
    Else
        Response.Write "    <td align=center width='30%'>无须分析</td>"
    End If
    Response.Write "    <td align=center width='20%'>常用屏幕显示颜色</td>"
    If FoundInArr(RegCount_Fill, "FColor", ",") = True Then
        Response.Write "    <td align=center width='30%'>" & MaxColor & " (" & MaxColorNum & ")</td>"
    Else
        Response.Write "    <td align=center width='30%'>无须分析</td>"
    End If
    Response.Write "  </tr>"
    Response.Write "  </tbody>"
    Response.Write "</table>"
    
    Set rs = Nothing
End Sub

Sub FVisitor()
    strGuide = "最近访问记录"
    sql = "Select * From PE_StatVisitor Order By VTime DESC"
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn_Counter, 1, 1
    If rs.BOF And rs.EOF Then
        Response.Write "<li>系统中无数据！"
    Else
        TotalPut = rs.RecordCount
        TitleRight = TitleRight & "共 <font color=red>" & TotalPut & "</font> 个访问记录"
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

        Response.Write "<table width='100%'><tr><td align='left'>您现在的位置：网站统计管理&nbsp;&gt;&gt;&nbsp;" & Search & strGuide & "</td><td align='right'>" & TitleRight & "</td></tr></table>"
        Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='0' class='border'>"
        Response.Write "  <tr class=title height='22'>"
        Response.Write "    <td align=center height='22'>访问时间(服务器端)</td>"
        Response.Write "    <td align=center height='22'>访问时间(客户端)</td>"
        Response.Write "    <td align=center height='22'>访问者IP</td>"
        Response.Write "    <td align=center height='22'>地址</td>"
        Response.Write "    <td align=center height='22'>链接页面</td>"
        Response.Write "    <td align=center height='22'>操作</td>"
        Response.Write "  </tr>"
        Do While Not rs.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td align=left width='120' height='22'>" & rs("VTime") & "</td>"
            Response.Write "    <td align=left width='120' height='22'>" & DateAdd("h", -rs("Timezone") - MasterTimeZone, rs("VTime")) & "</td>"
            Response.Write "    <td align=left width='80' height='22'>" & rs("IP") & "</td>"
            Response.Write "    <td align=left width='100' height='22'>" & rs("Address") & "</td>"
            Response.Write "    <td align=left height='22'><a href='" & rs("Referer") & "' title='" & rs("Referer") & "' target='_blank'>" & Left(rs("Referer"), 40) & "</a></td>"
            Response.Write "    <td align=left width='60' height='22'><a href='Admin_Counter.asp?Action=ClientDetail&id=" & rs("Id") & "'>查看明细</a></td>"
            Response.Write "  </tr>"
            VisitorNum = VisitorNum + 1
            If VisitorNum >= MaxPerPage Then Exit Do
            rs.MoveNext
        Loop
        Response.Write "</table>"
        If TotalPut > 0 Then
            Response.Write ShowPage(strFileName, TotalPut, MaxPerPage, CurrentPage, True, True, "个访问记录", True)
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub FCounter()
    Item = Array("首次", "二次", "三次", "四次", "五次", "六次", "七次", "八次", "九次", "十次以上")
    ItemNum = 10
    strGuide = "访问次数统计分析"
    StatItem = "次数分析"
    sql = "Select * From PE_StatVisit"
    Call Stable
End Sub

Sub StatYear()
    If Request("Type") = "" Then
       QYear = CStr(Year(Date))
    Else
       Search = "查询结果："
    End If
    ItemNum = 12
    ReDim Item(11)
    For i = 0 To 11
      Item(i) = QYear & "年" & i + 1 & "月"
    Next
    strGuide = QYear & "年访问统计分析"
    StatItem = "月份"
    sql = "Select * From PE_StatYear Where TYear='" & QYear & "'"

    Call Stable
End Sub

Sub StatAllYear()
    ItemNum = 12
    ReDim Item(ItemNum)
    For i = 0 To ItemNum - 1
      Item(i) = i + 1 & "月"
    Next

    strGuide = "全部年访问统计分析"
    StatItem = "月份"
    sql = "Select * From PE_StatYear Where TYear='Total'"
    Call Stable
End Sub

Sub StatMonth()
    If Request("Type") = "" Then
       QMonth = CStr(Year(Date) & "-" & Month(Date))
    Else
       Search = "查询结果："
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
      Item(i) = SYear & "年" & SMonth & "月" & i + 1 & "日"
    Next
    strGuide = QMonth & "月访问统计分析"
    StatItem = "日期"
    sql = "Select * From PE_StatMonth Where TMonth='" & QMonth & "'"
    Call Stable
End Sub

Sub StatAllMonth()
    ItemNum = 31
    ReDim Item(ItemNum)
    For i = 0 To ItemNum - 1
      Item(i) = i + 1 & "日"
    Next
    strGuide = "全部月访问统计分析"
    StatItem = "日期"
    sql = "Select * From PE_StatMonth Where TMonth='Total'"
    Call Stable
End Sub

Sub StatWeek()
    Item = Array("星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六")
    ItemNum = 7
    strGuide = "本周访问统计分析"
    StatItem = "星期"
    sql = "Select * From PE_StatWeek Where Tweek='Current'"
    Call Stable
End Sub

Sub StatAllWeek()
    Item = Array("星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六")
    ItemNum = 7
    strGuide = "全部周访问统计分析"
    StatItem = "星期"
    sql = "Select * From PE_StatWeek Where Tweek='Total'"
    Call Stable
End Sub

Sub StatDay()
    If Request("Type") = "" Then
       QDay = CStr(Year(Date) & "-" & Month(Date) & "-" & Day(Date))
    Else
       Search = "查询结果："
    End If
    ItemNum = 24
    ReDim Item(23)
    For i = 0 To ItemNum - 1
      Item(i) = Mid(i + 100, 2) & ":00-" & Mid(i + 101, 2) & ":00"
    Next
    strGuide = QDay & "日访问统计分析"
    StatItem = "小时"
    sql = "Select * From PE_StatDay Where TDay='" & QDay & "'"
    Call Stable
End Sub

Sub StatAllDay()
    ItemNum = 24
    ReDim Item(ItemNum)
    For i = 0 To ItemNum - 1
      Item(i) = Mid(i + 100, 2) & ":00-" & Mid(i + 101, 2) & ":00"
    Next
    strGuide = "全部日访问统计分析"
    StatItem = "小时"
    sql = "Select * From PE_StatDay Where TDay='Total'"
    Call Stable
End Sub

Sub FIP()
    sql = "Select * From PE_StatIp Order By TIpNum DESC"
    strGuide = "访问者IP地址分析"
    StatItem = "IP地址"
    Call Ftable
End Sub

Sub FAddress()
    sql = "Select * From PE_StatAddress Order By TAddNum DESC"
    strGuide = "访问者所在地址分析"
    StatItem = "地址"
    Call Ftable
End Sub

Sub FTimezone()
    sql = "Select * From PE_StatTimezone Order By TtimNum DESC"
    strGuide = "访问者所处时区分析"
    StatItem = "时区"
    Call Ftable
End Sub

Sub FWeburl()
    sql = "Select * From PE_StatWeburl Order By TWebNum DESC"
    strGuide = "访问者来访网站分析"
    StatItem = "来访网站"
    Call Ftable
End Sub

Sub FKeyword()
    sql = "Select * From PE_StatKeyword Order By TKeywordNum DESC"
    strGuide = "访问者搜索关键词分析"
    StatItem = "关 键 词"
    Call Ftable
End Sub

Sub FReferer()
    sql = "Select * From PE_StatRefer Order By TRefNum DESC"
    strGuide = "访问者链接页面分析"
    StatItem = "链接页面"
    Call Ftable
End Sub

Sub FSystem()
    sql = "Select * From PE_StatSystem Order By TSysNum DESC"
    strGuide = "访问者所用操作系统分析"
    StatItem = "操作系统"
    Call Ftable
End Sub

Sub FBrowser()
    sql = "Select * From PE_StatBrowser Order By TBrwNum DESC"
    strGuide = "访问者所用浏览器分析"
    StatItem = "浏览器"
    Call Ftable
End Sub

Sub FMozilla()
    sql = "Select * From PE_StatMozilla Order By TMozNum DESC"
    strGuide = "访问者HTTP_USER_AGENT字符串分析"
    StatItem = "USER_AGENT"
    Call Ftable
End Sub

Sub FScreen()
    sql = "Select * From PE_StatScreen Order By TScrNum DESC"
    strGuide = "访问者屏幕大小分析"
    StatItem = "屏幕大小"
    Call Ftable
End Sub

Sub FColor()
    sql = "Select * From PE_StatColor Order By TColNum DESC"
    strGuide = "访问者屏幕显示颜色分析"
    StatItem = "屏幕显示颜色"
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
    TitleRight = "有效统计：<font color=red>" & TotalNum & "</font>"
    If Rows < 0 Then
        Response.Write "<li>系统中无数据！"
    Else
        Response.Write "<table width='100%'><tr><td align='left'>您现在的位置：网站统计管理&nbsp;&gt;&gt;&nbsp;" & Search & strGuide & "</td><td align='right'>" & TitleRight & "</td></tr></table>"
        Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
        Response.Write "  <tr class=title>"
        Response.Write "    <td align=left width='30%' nowrap height='22'>" & StatItem & "</td>"
        Response.Write "    <td align=left width='20%' nowrap>访问人数</td>"
        Response.Write "    <td align=left width='20%' nowrap>百分比</td>"
        Response.Write "    <td align=left width='30%' nowrap>图示</td>"
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
        Response.Write "<li>系统中无数据！"
    Else
        TotalPut = rs.RecordCount
        TitleRight = TitleRight & "有效统计：<font color=red>" & TotalNum & "</font>"
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
        Response.Write "<table width='100%'><tr><td align='left'>您现在的位置：网站统计管理&nbsp;&gt;&gt;&nbsp;" & Search & strGuide & "</td><td align='right'>" & TitleRight & "</td></tr></table>"
        Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
        Response.Write "  <tr class=title>"
        Response.Write "    <td align=left width='30%' nowrap height='22'>" & StatItem & "</td>"
        Response.Write "    <td align=left width='20%' nowrap>访问人数</td>"
        Response.Write "    <td align=left width='20%' nowrap>百分比</td>"
        Response.Write "    <td align=left width='30%' nowrap>图示</td>"
        Response.Write "  </tr>"
        Do While Not rs.EOF
            Response.Write "  <tr class='tdbg'>"
            Response.Write "    <td align=left nowrap>"
            If (Action = "FWeburl" Or Action = "FReferer") And rs(0) <> "直接输入或书签导入" Then
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
            Response.Write ShowPage(strFileName, TotalPut, MaxPerPage, CurrentPage, True, True, "个访问记录", True)
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub FOnline()
    Dim OnNowTime
    OnNowTime = DateAdd("s", -OnlineTime, Now())
    strGuide = "当前在线用户分析"
    If CountDatabaseType = "SQL" Then
        sql = "select * from PE_StatOnline where LastTime>'" & OnNowTime & "' order by OnTime desc"
    Else
        sql = "select * from PE_StatOnline where LastTime>#" & OnNowTime & "# order by OnTime desc"
    End If

    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn_Counter, 1, 1
    If rs.BOF And rs.EOF Then
        Response.Write "<li>当前无人在线！"
    Else
        TotalPut = rs.RecordCount
        TitleRight = TitleRight & "共 <font color=red>" & TotalPut & "</font> 个用户在线"
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

        Response.Write "<table width='100%'><tr><td align='left'>您现在的位置：网站统计管理&nbsp;&gt;&gt;&nbsp;" & Search & strGuide & "</td><td align='right'>" & TitleRight & "</td></tr></table>"
        Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
        Response.Write "  <tr class=title>"
        Response.Write "    <td align=center nowrap height='22'>编号</td>"
        Response.Write "    <td align=center nowrap>访问者IP</td>"
        Response.Write "    <td align=center nowrap>上站时间</td>"
        Response.Write "    <td align=center nowrap>最后刷新时间</td>"
        Response.Write "    <td align=center nowrap>已停留时间</td>"
        Response.Write "    <td align=center nowrap>所在页面 及 客户端信息</td>"
        Response.Write "  </tr>"
        
        Do While Not rs.EOF
            LNowTime = Cstrtime(CDate(Now() - rs("Ontime")))
            Response.Write "  <tr class='tdbg'>"
            Response.Write "    <td align=center width='8%' nowrap>" & VisitorNum & "</td>"
            Response.Write "    <td align=left width='15%' nowrap>" & rs("UserIP") & "</td>"
            Response.Write "    <td align=left width='17%' nowrap><a title=" & rs("OnTime") & ">" & TimeValue(rs("OnTime")) & "</a></td>"
            Response.Write "    <td align=left width='15%' nowrap>" & TimeValue(rs("LastTime")) & "</td>"
            Response.Write "    <td align=left width='15%' nowrap>" & LNowTime & "</td>"
            Response.Write "    <td align=left width='45%' nowrap title='所在页面: " & rs("UserPage") & vbCrLf & "客户端信息: " & rs("UserAgent") & "'><a href=" & rs("UserPage") & " target=""_blank"">" & Left(Findpages(rs("UserPage")), 35) & "</a>"
            Response.Write "    </td>"
            Response.Write "  </tr>"
            VisitorNum = VisitorNum + 1
            If VisitorNum >= MaxPerPage Then Exit Do
            rs.MoveNext
        Loop
        Response.Write "</table>"
        If TotalPut > 0 Then
            Response.Write ShowPage(strFileName, TotalPut, MaxPerPage, CurrentPage, True, True, "个在线用户", True)
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
    Response.Write "      <td width='120'><strong>网站统计查询：</strong></td>"
    Response.Write "      <td>报表类型： "
    Response.Write "        <select name='type' size='1' class='Select' onChange=change_type()>"
    Response.Write "          <option value='1' selected>日报表</option>"
    Response.Write "          <option value='2'>月报表</option>"
    Response.Write "          <option value='3'>年报表</option>"
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
    Response.Write "        年"
    Response.Write "        <select name='qmonth' size='1' onChange=change_it()>"
    For i = 1 To 12
        If i = Month(Date) Then
            Response.Write "<option value='" & i & "' selected>" & i & "</option>"
        Else
            Response.Write "<option value='" & i & "'>" & i & "</option>"
        End If
    Next
    Response.Write "        </select>"
    Response.Write "        月"
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
    Response.Write "        日"
    Response.Write "        <input type='submit' name='Search' value='查询'>"
    Response.Write "      </td>"
    Response.Write "      <td width='120' align='center'> </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub AutoAnalyse()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if(confirm('确实要进行统计数据自动分析吗？'))" & vbCrLf
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
    Response.Write "    <td height='22' align='center'><strong> 统 计 数 据 自 动 分 析 </strong></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td height='150'>" & vbCrLf
    Response.Write "      <form name='myform' method='post' action='Admin_Counter.asp' onSubmit='return CheckForm();'>" & vbCrLf
    Response.Write "        <p align='center'>" & vbCrLf
    Response.Write "        <font color='#FF0000'><b>请慎用此功能，因为一旦清除将无法恢复！</b></font>" & vbCrLf
    Response.Write "        <br>此操作将清除数据库中近期访问记录表的数据，用于近期对网站的访问统计数据进行统计分析时使用。" & vbCrLf
    Response.Write "        </p>" & vbCrLf
    Response.Write "        <p align='center'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='DoAutoAnalyse'>" & vbCrLf
    Response.Write "        <input type='submit' name='Submit' value=' 统计数据自动分析 '>" & vbCrLf
    Response.Write "        </p>" & vbCrLf
    Response.Write "      </form>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub

Sub Init()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if(confirm('确实要进行初始化吗？一旦清除将无法恢复！'))" & vbCrLf
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
    Response.Write "    <td height='22' align='center'><strong> 数 据 初 始 化 </strong></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'>" & vbCrLf
    Response.Write "    <td height='150'>" & vbCrLf
    Response.Write "      <form name='myform' method='post' action='Admin_Counter.asp' onSubmit='return CheckForm();'>" & vbCrLf
    Response.Write "        <p align='center'>" & vbCrLf
    Response.Write "        <font color='#FF0000'><b>请慎用此功能，因为一旦清除将无法恢复！</b></font>" & vbCrLf
    Response.Write "        <br>此操作将清除数据库中的所有统计数据，用于系统初始化时及需要对网站的访问统计数据进行重新统计时使用。" & vbCrLf
    Response.Write "        </p>" & vbCrLf
    Response.Write "        <p align='center'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='DoInit'>" & vbCrLf
    Response.Write "        <input type='submit' name='Submit' value=' 统计数据初始化 '>" & vbCrLf
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
    Response.Write "     <td align='center' height='22' valign='middle'><b>统计数据库在线压缩</b></td>"
    Response.Write " </tr>"
    Response.Write " <tr class='tdbg'>"
    Response.Write "     <td align='center' height='150' valign='middle'>"
    Response.Write "      <br>"
    Response.Write "      <br>"
    Response.Write "      压缩前，建议先备份统计数据库，以免发生意外错误。 <br>"
    Response.Write "      <br>"
    Response.Write "      <br>"
    Response.Write " <input name='submit' type=submit value=' 压缩统计数据库 '"
    If CountDatabaseType = "SQL" Then
        Response.Write " disabled"
    End If
    Response.Write "><br><br>"
    If ObjInstalled_FSO = False Then
        Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
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
        Call WriteSuccessMsg("统计数据库压缩成功！", ComeUrl)
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>统计数据库没有找到!</li>"
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
    Call WriteSuccessMsg("统计数据初始化成功！", ComeUrl)
End Sub

Sub ClientDetail()
    Dim ClientNow
    Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from PE_StatVisitor where id=" & PE_CLng(Request("id"))
    rs.Open sql, Conn_Counter, 1, 1
    Response.Write "    <br><table width='100%' class='border' border='0' cellspacing='1' cellpadding='4' align='center'>" & vbCrLf
    Response.Write "      <tr class='title'> " & vbCrLf
    Response.Write "        <td colspan='2' class='title' align='center'><b>访 问 记 录 详 情 显 示</b></td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'><b>访问时间（以服务器端时区记）：</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("VTime") & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'><b>访问者IP：</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("IP") & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'><b>访问者所在时区：</b></td>"
    Response.Write "        <td width='70%'>GMT" & rs("Timezone") & ":0</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'> " & vbCrLf
    Response.Write "          <b>所在地址：</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("Address") & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'><b>访问时间（以客户端时区记）：</b></td>"
    Response.Write "        <td width='70%'>" & DateAdd("h", -rs("Timezone") - MasterTimeZone, rs("VTime")) & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'> " & vbCrLf
    Response.Write "          <b>链接页面：</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("Referer") & "</td></tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'> " & vbCrLf
    Response.Write "          <b>操作系统：</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("System") & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'> " & vbCrLf
    Response.Write "          <b>浏览器：</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("Browser") & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'> " & vbCrLf
    Response.Write "          <b>屏幕大小：</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("Screen") & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td width='30%' nowrap align='left'> " & vbCrLf
    Response.Write "          <b>屏幕色深：</b></td>" & vbCrLf
    Response.Write "        <td width='70%'>" & rs("Color") & "</td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "      <tr class='tdbg'> " & vbCrLf
    Response.Write "        <td colspan='2' nowrap> " & vbCrLf
    Response.Write "          <div align='center'>" & vbCrLf
    Response.Write "            <input type='button' name='Submit2' value='返回' onClick=""window.location='Admin_Counter.asp?Action=FVisitor';"">" & vbCrLf
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
    Response.Write "               <div align='center'><strong>统计IP库添加</strong></div>" & vbCrLf
    Response.Write "            </td>    " & vbCrLf
    Response.Write "         </tr>    " & vbCrLf
    Response.Write "               <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='350' class='tdbg5'><strong>起始 I P：</strong><br>注： 添加的IP，如果是数据库中尚没有的记录<br>将直接添加，如果在数据库中已经存在，将提示您是否进行修改。</td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='StartIP' type='text' id='StartIP' size='49' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "         </tr>    " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='350' class='tdbg5'><strong>结尾 I P：</strong></td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='EndIP' type='text' id='EndIP' size='49' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>  " & vbCrLf
    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "               <td width='350' class='tdbg5'><strong>来源详细地址：</strong>" & vbCrLf
    Response.Write "               </td>      " & vbCrLf
    Response.Write "               <td class='tdbg'><input name='IPAddress' type='text' id='IPAddress' size='49' maxlength='30'>&nbsp;</td>    " & vbCrLf
    Response.Write "        </tr>   " & vbCrLf
    Response.Write "        <tr class='tdbg'>     " & vbCrLf
    Response.Write "                     <td colspan='2' align='center' class='tdbg'>" & vbCrLf
    Response.Write "                     " & vbCrLf
    Response.Write "                     <input name='Action' type='hidden' id='Action' value='SaveIPAdd'>        <input  type='submit' name='Submit' value=' 添 加 '>&nbsp;&nbsp;        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Counter.asp'"" style='cursor:hand;'>" & vbCrLf
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
        ErrMsg = ErrMsg & "请填写正确的IP地址！"
        Exit Sub
    End If
    If Request.Form("EndIP") = "" Or Not isIP(Request.Form("StartIP")) Then
        FoundErr = True
        ErrMsg = ErrMsg & "请填写正确的IP地址！"
        Exit Sub
    End If
    If Request.Form("IPAddress") = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "请填写来源详细地址！"
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
        Call WriteSuccessMsg("网站统计IP添加成功！", ComeUrl)
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "添加失败，数据已存在，请搜索该ip地址并进行修改。"
    End If
    rs.Close
End Sub


Sub IPManage()
    Response.Write "<table width='100%'><tr><td align='left'>您现在的位置：统计IP库管理 &nbsp;&gt;&gt;&nbsp;IP 地址库管理"
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
        Response.Write "共找到 <font color=red>0</font> 条IP段记录</td></tr></table>"
    Else
        'If SearchAddress="" and request("SearchIP")="" Then
            'Response.Write "所有IP段记录!</td></tr></table>"
        'Else
            Response.Write "共 <font color=red>" & TotalPut & "</font> 条IP段记录</td></tr></table>"
        'End If
    End If
    Response.Write "    <table width='100%' class='border' border='0' cellspacing='1' cellpadding='0' align='center'>" & vbCrLf
    Response.Write "      <tr class='title'>" & vbCrLf
    Response.Write "        <td width='20%' height='22'> " & vbCrLf
    Response.Write "          <div align='center'><b>起始 IP</b></div>" & vbCrLf
    Response.Write "        </td>" & vbCrLf
    Response.Write "        <td width='20%' height='22'> " & vbCrLf
    Response.Write "          <div align='center'><b>结尾 IP</b></div>" & vbCrLf
    Response.Write "        </td>" & vbCrLf
    Response.Write "        <td width='42%' height='22'> " & vbCrLf
    Response.Write "          <div align='center'><b>来源详细地址</b></div>" & vbCrLf
    Response.Write "        </td>" & vbCrLf
    Response.Write "        <td width='18%' nowrap> " & vbCrLf
    Response.Write "          <div align='center'><b>操作</b></div>" & vbCrLf
    Response.Write "        </td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    If rs.RecordCount = 0 Then
        Response.Write "<tr class='tdbg'><td colspan='10' align='center'>没有满足条件的IP 段记录!</td><tr>" & vbCrLf
    Else
        Dim rsID, i, Sort
        i = 0
        Do While Not rs.EOF
            Response.Write "    <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            Response.Write "      <td width='20%' align='left' height='22'>" & DecodeIP(rs("StartIp")) & "</td>" & vbCrLf
            Response.Write "      <td width='20%' align='left' height='22'>" & DecodeIP(rs("EndIp")) & "</td>" & vbCrLf
            Response.Write "      <td width='42%' align='left' height='22'>" & rs("Address") & "</td>" & vbCrLf
            Response.Write "      <td width='18%' align='center' height='22'>" & vbCrLf
            Response.Write "        <a href='?action=editIP&StartIP=" & rs("StartIP") & "&EndIP=" & rs("EndIP") & "'>编辑</a>  |  <a href='?action=delIP&StartIP=" & rs("StartIP") & "&EndIP=" & rs("EndIP") & "'>删除</a>" & vbCrLf
            Response.Write "    </td></tr>" & vbCrLf
            i = i + 1
            If i > MaxPerPage Then Exit Do
            rs.MoveNext
        Loop
    End If
    Response.Write "  </table>" & vbCrLf
    If TotalPut > 0 Then
        Response.Write ShowPage(strFileName, TotalPut, MaxPerPage, CurrentPage, True, True, "条记录", True)
    End If
    Call IPSearch
End Sub

Sub editIP()
    Dim StartIP, EndIP
    If Request("StartIP") = "" And Request("EndIP") = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "请给定IP地址！"
        Exit Sub
    End If
    If Not (IsNumeric(Request("StartIP")) Or IsNumeric(Request("EndIP"))) Then
        FoundErr = True
        ErrMsg = ErrMsg & "给定IP地址有误！"
        Exit Sub
    End If
    StartIP = Trim(Request("StartIP"))
    EndIP = Trim(Request("EndIP"))
    Set rs = Conn_Counter.Execute("select StartIP,EndIP,Address from PE_StatIpInfo where StartIP=" & StartIP & " and EndIP=" & EndIP & "")
    If rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "错误，数据不存在该ip地址。"
    Else
        Response.Write "    <form method='post' action='Admin_Counter.asp' name='myform'>" & vbCrLf
        Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' > " & vbCrLf
        Response.Write "         <tr class='title'>" & vbCrLf
        Response.Write "            <td height='22' colspan='2'> " & vbCrLf
        Response.Write "               <div align='center'><strong>统计IP库修改</strong></div>" & vbCrLf
        Response.Write "            </td>    " & vbCrLf
        Response.Write "         </tr>    " & vbCrLf
        Response.Write "               <tr class='tdbg5'>      " & vbCrLf
        Response.Write "               <td width='350' class='tdbg5'><strong>起始 I P：</strong></td>      " & vbCrLf
        Response.Write "               <td class='tdbg'><input name='StartIP' type='text' id='StartIP' size='49' maxlength='30' value='" & DecodeIP(rs(0)) & "'>&nbsp;</td>    " & vbCrLf
        Response.Write "         </tr>    " & vbCrLf
        Response.Write "        <tr class='tdbg'>      " & vbCrLf
        Response.Write "               <td width='350' class='tdbg5'><strong>结尾 I P：</strong></td>      " & vbCrLf
        Response.Write "               <td class='tdbg'><input name='EndIP' type='text' id='EndIP' size='49' maxlength='30' value='" & DecodeIP(rs(1)) & "'>&nbsp;</td>    " & vbCrLf
        Response.Write "        </tr>  " & vbCrLf
        Response.Write "        <tr class='tdbg'>      " & vbCrLf
        Response.Write "               <td width='350' class='tdbg5'><strong>来源详细地址：</strong></td>      " & vbCrLf
        Response.Write "               <td class='tdbg'><input name='IPAddress' type='text' id='IPAddress' size='49' maxlength='30' value='" & rs(2) & "'>&nbsp;</td>    " & vbCrLf
        Response.Write "        </tr>   " & vbCrLf
        Response.Write "        <tr class='tdbg'>     " & vbCrLf
        Response.Write "                     <td colspan='2' align='center' class='tdbg'>" & vbCrLf
        Response.Write "                     " & vbCrLf
        Response.Write "                     <input name='oldStartIP' type='hidden' id='oldStartIP' value='" & rs(0) & "'>" & vbCrLf
        Response.Write "                     <input name='oldEndIP' type='hidden' id='oldEndIP' value='" & rs(1) & "'>" & vbCrLf
        Response.Write "                     <input name='Action' type='hidden' id='Action' value='SaveIPedit'>        <input  type='submit' name='Submit' value=' 修 改 '>&nbsp;&nbsp;        <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Counter.asp'"" style='cursor:hand;'>" & vbCrLf
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
        ErrMsg = ErrMsg & "请填写正确开始IP地址！"
        Exit Sub
    End If
    If Request.Form("EndIP") = "" Or Not isIP(Request.Form("EndIP")) Then
        FoundErr = True
        ErrMsg = ErrMsg & "请填写正确结束IP地址！"
        Exit Sub
    End If
    If Request.Form("oldStartIP") = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "IP地址丢失！"
        Exit Sub
    End If
    If Request.Form("oldEndIP") = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "IP地址丢失！"
        Exit Sub
    End If
    If Request.Form("IPAddress") = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "请填写来源详细地址！"
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
        ErrMsg = ErrMsg & "IP 修改失败，请搜索该ip地址重新进行修改。"
    Else
        Call WriteSuccessMsg("网站统计IP修改成功！", "admin_counter.asp?Action=IPManage")
    End If
End Sub

Sub delIP()
    Dim StartIP, EndIP
    If Request("StartIP") = "" And Request("EndIP") = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "请给定IP地址！"
        Exit Sub
    End If
    If Not (IsNumeric(Request("StartIP")) Or IsNumeric(Request("EndIP"))) Then
        FoundErr = True
        ErrMsg = ErrMsg & "给定IP地址有误！"
        Exit Sub
    End If
    StartIP = Trim(Request("StartIP"))
    EndIP = Trim(Request("EndIP"))
    Dim RowCount
    Conn_Counter.Execute ("delete from PE_StatIpInfo where StartIP=" & StartIP & " and EndIP=" & EndIP & ""), RowCount
    If RowCount = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "IP 删除失败，请搜索该ip地址重新删除。"
    Else
        Call WriteSuccessMsg("网站统计IP删除成功！", ComeUrl)
    End If
End Sub

Sub IPSearch()
    Response.Write "    <form method='post' action='Admin_Counter.asp' name='myform'>" & vbCrLf
    Response.Write "      <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' > " & vbCrLf

    Response.Write "        <tr class='tdbg'>      " & vbCrLf
    Response.Write "            <td width='120'><strong>统计I P库搜索：</strong></td>"
    Response.Write "            <td>I P 地址：</td>" & vbCrLf
    Response.Write "            <td><input name='SearchIP' type='text' id='SearchIP' size='20' maxlength='20'>&nbsp;</td>" & vbCrLf
    Response.Write "            <td>来源详细地址：</td>" & vbCrLf
    Response.Write "            <td><input name='SearchAddress' type='text' id='SearchAddress' size='20' maxlength='30'>&nbsp;</td>" & vbCrLf
    Response.Write "            <td>" & vbCrLf
    Response.Write "                     <input name='Action' type='hidden' id='Action' value='SearchIP'>        <input  type='submit' name='Submit' value=' 搜 索 '>" & vbCrLf
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
    Response.Write "      <td height='22' align='center'><strong>IP数据库导入</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;请输入要导入的模板数据库的文件名： "
    Response.Write "        <input name='TemplateMdb' type='text' id='TemplateMdb' value='../Count/IP.mdb' size='20' maxlength='50'>"
    Response.Write "        <input align=""center"" name='Submit' type='submit' id='Submit' value=' 下一步 '>"
    Response.Write "        <br><font color=""#FF0000"">&nbsp;&nbsp;&nbsp;&nbsp;注意：导入的IP数据会直接覆盖原来的IP数据，请做好备份工作。</Font><br>"		
    Response.Write "        <input name='Action' type='hidden' id='Action' value='DoImport'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub


Sub DoImport()
    On Error Resume Next
    Dim mdbname, tconn, trs, iCount
    
    '获得导入模板数据库路径
    mdbname = Replace(Trim(Request.Form("Templatemdb")), "'", "")

    If mdbname = "" Then
        mdbname = Replace(Trim(Request.QueryString("Templatemdb")), "'", "")
    End If

    mdbname = Replace(mdbname, "＄", "/") '防止外部链接安全问题

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请填写导入模板数据库名"
        Exit Sub
    End If

    '建立导入模板数据库
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>数据库操作失败，请以后再试，错误原因：" & Err.Description
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
        ErrMsg = ErrMsg & "<li>您要导入的数据库,不是系统方案数据库,请使用系统方案数据库。"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    		
    Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    Response.Write "  <tr align='center' class='title'><td height='22'><strong>数据导入中，请稍候</strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='100' valign='top'><br><Div name=""ShowMess"" id=""ShowMess"" >正在初始化数据库</Div></td></tr>" & vbCrLf
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
            Response.Write "document.getElementById(""ShowMess"").innerHTML=""数据转换中,请稍候<br>已经成功导入"& countIPNum &"条IP数据"";" & vbCrLf
            Response.Write "</script>" & vbCrLf			
            response.Flush()
        End If	
    Loop
    Response.Write "<script>" & vbCrLf
    Response.Write "document.getElementById(""ShowMess"").innerHTML=""成功导入所有IP数据"";" & vbCrLf
    Response.Write "</script>" & vbCrLf		
    Response.Flush()	
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
	
    tconn.Close
    Set tconn = Nothing
   ' Call WriteSuccessMsg("已经成功将IP数据库导入！", ComeUrl)
End Sub

Sub Export()
    Response.Write "<form name='myform' method='post' action='Admin_Counter.asp'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'> "
    Response.Write "      <td height='22' align='center'><strong>IP数据库导出</strong></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;请输入要导出的模板数据库的文件名： "
    Response.Write "        <input name='TemplateMdb' type='text' id='TemplateMdb' value='../Count/IP.mdb' size='20' maxlength='50'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 下一步 '>"
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

    mdbname = Replace(mdbname, "＄", "/") '防止外部链接安全问题

    If mdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请填写导入模板数据库名"
        Exit Sub
    End If
    '建立导出IP数据库
    Set tconn = Server.CreateObject("ADODB.Connection")
    tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)

    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>数据库操作失败，请以后再试，错误原因：" & Err.Description
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
        ErrMsg = ErrMsg & "<li>您要导入的数据库,不是系统方案数据库,请使用系统方案数据库。"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If
		
	
    tconn.Execute ("delete from PE_StatIpInfo")
    Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    Response.Write "  <tr align='center' class='title'><td height='22'><strong>数据导出中，请稍候，数据导出时请勿刷新页面</strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='100' valign='top'><br><Div name=""ShowMess"" id=""ShowMess"" >正在初始化ip数据库，请稍候！</Div></td></tr>" & vbCrLf
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
            Response.Write "document.getElementById(""ShowMess"").innerHTML=""数据转换中,请稍候<br>已经成功导出"& countIPNum &"条IP数据"";" & vbCrLf
            Response.Write "</script>" & vbCrLf			
            response.Flush()
        End If			
    Loop
    Response.Write "<script>" & vbCrLf
    Response.Write "document.getElementById(""ShowMess"").innerHTML=""成功导出所有IP数据"";" & vbCrLf
    Response.Write "</script>" & vbCrLf		
    Response.Flush()
    trs.Close
    Set trs = Nothing
    rs.Close
    Set rs = Nothing
	
    tconn.Close
    Set tconn = Nothing
    'Call WriteSuccessMsg("已经成功将IP数据库导出！", ComeUrl)
End Sub



%>
