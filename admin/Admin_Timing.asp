<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = False   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = "Timing"   '其他权限

strFileName = "Admin_Timing.asp?Action=" & Action

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>定 时 系 统 项 目 管 理</title>" & vbCrLf
Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""Admin_Style.css"">" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">" & vbCrLf

If Action <> "DoTiming2" Then
    Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""border"">" & vbCrLf
    Call ShowPageTitle(" 定 时 系 统 项 目 管 理 ", 10055)
    Response.Write "</table>" & vbCrLf
End If

Select Case Action
    Case "DoMainTiming"
        Call DoMainTiming
    Case "DoTiming"
        Call DoTiming
    Case "DoTiming2"
        Call DoTiming2
    Case "SaveTiming", "SaveModify"
        Call SaveTiming
    Case Else
        Call main
End Select

If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

'=================================================
'过程名：main()
'作  用：定时采集管理
'=================================================
Sub main()
    Dim rs, rsTiming, sql, iCount
    Dim Timing_CollectionItemID, Timing_SetDate, Timing_SetWeekday, Timing_SetDay, Timing_Time, Timing_Passed, Timing_Renovate, Timing_Date, Timing_AreaCollection
    Dim arrChannelID, i, CreateItemType, CreateItemTopNewNum, CreateItemDate, CreateClass, CreateSpecial, CreateChannel
    Dim arrTimingCreateSetting
    '得到定时属性
    sql = "select Timing_CollectionItemID,Timing_SetDate,Timing_SetWeekday,Timing_SetDay,Timing_Time,Timing_Date,Timing_AreaCollection from PE_Config"
    Set rsTiming = Server.CreateObject("adodb.recordset")
    rsTiming.Open sql, Conn, 1, 1

    If rsTiming.EOF Then   '没有找到该项目
    Else
        Timing_CollectionItemID = rsTiming("Timing_CollectionItemID")
        Timing_SetDate = rsTiming("Timing_SetDate")
        Timing_Time = rsTiming("Timing_Time")
        Timing_SetWeekday = rsTiming("Timing_SetWeekday")
        Timing_SetDay = rsTiming("Timing_SetDay")
        Timing_Date = rsTiming("Timing_Date")
        Timing_AreaCollection = rsTiming("Timing_AreaCollection")
    End If

    rsTiming.Close
    Set rsTiming = Nothing

    Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"">" & vbCrLf
    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""70"" height=""30""><strong>管理导航：</strong></td>" & vbCrLf
    Response.Write "    <td height=""30""><a href=Admin_Timing.asp?Action=Main>管理首页</a> | <a href=""Admin_Timing.asp?Action=DoMainTiming"" target='_blank'>启动定时项目</a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>"
    Response.Write "<br>" & vbCrLf
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
    Response.Write "        function ShowTimingType(num){ " & vbCrLf
    Response.Write "                switch(num){" & vbCrLf
    Response.Write "                case ""0"":     " & vbCrLf
    Response.Write "                        document.myform.Timing_SetWeekday.style.display='none';" & vbCrLf
    Response.Write "                        document.myform.Timing_SetDay.style.display='none';" & vbCrLf
    Response.Write "                        break;" & vbCrLf
    Response.Write "                case ""1"":     " & vbCrLf
    Response.Write "                        document.myform.Timing_SetWeekday.style.display='';" & vbCrLf
    Response.Write "                        document.myform.Timing_SetDay.style.display='none';" & vbCrLf
    Response.Write "                        break;" & vbCrLf
    Response.Write "                case ""2"":" & vbCrLf
    Response.Write "                        document.myform.Timing_SetWeekday.style.display='none';" & vbCrLf
    Response.Write "                        document.myform.Timing_SetDay.style.display='';" & vbCrLf
    Response.Write "                        break;" & vbCrLf
    Response.Write "                default:" & vbCrLf
    Response.Write "                        alert(""错误参数调用！"");" & vbCrLf
    Response.Write "                        break;" & vbCrLf
    Response.Write "                }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<form name='myform' method='post' action='Admin_Timing.asp?action=SaveTiming'>" & vbCrLf
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr align='center' height='24'>"
    Response.Write "<td id='TabTitle' class='title6' onclick='ShowTabs(0)'>定时采集</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(1)'>定时生成</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(2)'>定时设置</td>" & vbCrLf
    Response.Write "<td>&nbsp;</td></tr></table>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
   
    Response.Write " <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write " <tr class='tdbg'>"
    Response.Write "   <td align='center'>"
    Response.Write "     <table border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "      <tr>"
    Response.Write "       <td>请选择定时采集的项目：<br>"
    '得到采集属性
    sql = "select * from PE_Item where Flag=" & PE_True
    Set rs = Conn.Execute(sql)
    Response.Write "            <select name='Timing_CollectionItemID' size='2' multiple style='height:300px;width:450px;'>"

    If rs.BOF And rs.EOF Then
        Response.Write "         <option value=''>还没有采集项目！</option>"
        '关闭提交按钮
        iCount = 0
    Else
        iCount = rs.RecordCount

        Do While Not rs.EOF
            Response.Write "     <option value='" & rs("ItemID") & "' "

            If FoundInArr(Timing_CollectionItemID, rs("ItemID"), ",") = True Then
                Response.Write " selected"
            End If

            Response.Write " >" & rs("ItemName") & "</option>"
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing

    Response.Write "         </select>"
    Response.Write "       </td>"
    Response.Write "       <td align='left'>&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='Submit' value=' 选定所有 ' onclick='SelectAll()'>"
    Response.Write "       <br><br>&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='Submit' value=' 取消选定 ' onclick='UnSelectAll()'><br><br><br><b>&nbsp;提示：按住“Ctrl”或“Shift”键可以多选</b></td>"
    Response.Write "      </tr>"
    Response.Write "      <tr class='tdbg'><td><Input type='checkbox' Name='Timing_AreaCollection' value='1' " & IsRadioChecked(Timing_AreaCollection, "1") & "> 是否区域采集</td><td></td></tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    Response.Write "    </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr align='left'>" & vbCrLf
    Response.Write "      <td class='tdbg' valign='top' width='20%'>" & vbCrLf
    Response.Write "        <table width='100%' border='0' cellpadding='2' cellspacing='1'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>" & vbCrLf
    Response.Write "            <td>" & vbCrLf
    Response.Write "                           请选择定时生成的频道：<br>" & vbCrLf
    Dim SqlI, RsI
    SqlI = "select ChannelID,ChannelName,ModuleType from PE_Channel where ModuleType<>0 and ModuleType<>4 and Disabled=" & PE_False & " and UseCreateHTML > 0 order by ChannelID desc"
    Set RsI = Server.CreateObject("adodb.recordset")
    RsI.Open SqlI, Conn, 1, 1
    Response.Write "<table border='0' cellpadding='0' cellspacing='0' width='100%' height='100%' align='center'>"

    If RsI.EOF And RsI.BOF Then
        Response.Write "<center>您的频道还没有设置生成功能,或您指定的生成频道已经禁用。</center>"
    Else
        i = 0

        Do While Not RsI.EOF
            i = i + 1
            Response.Write "<tr id='tdcolor" & i & "'"

            If i = 1 Then Response.Write " bgcolor='#ffffff' "
            Response.Write " onCLICK='change_item(" & i & ")'>"
            Response.Write "   <td valign='top'> " & RsI("ChannelName") & "</td><td >属性设置 >>></td>"
            Response.Write "  </tr>"
            RsI.MoveNext
        Loop

        Response.Write "<script language='javascript'>" & vbCrLf
        Response.Write "function change_item(num){" & vbCrLf
        Response.Write "    for (td_i=1;td_i<=" & i & ";td_i++){" & vbCrLf
        Response.Write "        if (td_i==num){" & vbCrLf
        Response.Write "                    eval(""td_""+td_i+"".style.display='';"");" & vbCrLf
        Response.Write "                        eval(""tdcolor""+td_i+"".style.backgroundColor='#ffffff';"");  " & vbCrLf
        Response.Write "                }" & vbCrLf
        Response.Write "                else{" & vbCrLf
        Response.Write "                        eval(""td_""+td_i+"".style.display=\""none\"";"");" & vbCrLf
        Response.Write "                        eval(""tdcolor""+td_i+"".style.backgroundColor='#F0F0F0'"");  " & vbCrLf
        Response.Write "                }" & vbCrLf
        Response.Write "   }" & vbCrLf
        Response.Write "}" & vbCrLf
        Response.Write "</script>" & vbCrLf
    End If

    RsI.Close
    Set RsI = Nothing
    Response.Write "      </table> "
    Response.Write "     </td>" & vbCrLf
    Response.Write "     <td valign='top' width='80%'>" & vbCrLf

    sql = "select ChannelID,ChannelName,ModuleType,TimingCreateSetting from PE_Channel where ModuleType<>0 and ModuleType<>4 and Disabled=" & PE_False & " and UseCreateHTML > 0 order by ChannelID desc"
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 1

    If rs.EOF And rs.BOF Then
        Response.Write "<center>您的频道还没有设置生成功能,或您指定的生成频道已经禁用。</center>"
    Else
        i = 0

        Do While Not rs.EOF
            i = i + 1

            If InStr(rs("TimingCreateSetting"), ",") > 0 Then
                arrTimingCreateSetting = Split(rs("TimingCreateSetting"), ",")
                CreateItemType = PE_CLng(arrTimingCreateSetting(2))
                CreateItemTopNewNum = PE_CLng(arrTimingCreateSetting(3))
                CreateItemDate = arrTimingCreateSetting(4)
                CreateClass = arrTimingCreateSetting(5)
                CreateSpecial = arrTimingCreateSetting(6)
                CreateChannel = arrTimingCreateSetting(7)
            End If

            Response.Write "      <table width='100%' border='0' id='td_" & i & "' cellpadding='2' cellspacing='1' bgcolor='#ffffff' "

            If i > 1 Then Response.Write " style='display:none'"
            Response.Write " >" & vbCrLf
            Response.Write "       <tr class='tdbg'>" & vbCrLf
            Response.Write "          <td height='22' colspan='2' bgcolor='#E0EEF5'>&nbsp;&nbsp;项目内容页生成操作 </td>" & vbCrLf
            Response.Write "       </tr>" & vbCrLf
            Response.Write "       <tr class='tdbg'>" & vbCrLf
            Response.Write "          <td width='30' align='center'><input type='radio' name='CreateItemType" & i & "'  " & IsRadioChecked(CreateItemType, 0) & " value='0'></td>"
            Response.Write "          <td height='30'> 不生成项目内容页" & vbCrLf
            Response.Write "          </td>" & vbCrLf
            Response.Write "       </tr>" & vbCrLf
            Response.Write "       <tr class='tdbg'>" & vbCrLf
            Response.Write "           <td width='30' align='center'><input type='radio' name='CreateItemType" & i & "'  " & IsRadioChecked(CreateItemType, 1) & "  value='1'></td>"
            Response.Write "           <td height='30'> 生成最新" & vbCrLf
            Response.Write "             <Input name='CreateItemTopNewNum" & i & "' value='" & CreateItemTopNewNum & "' size=8 maxlength='10'> 篇项目" & vbCrLf
            Response.Write "           </td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
            Response.Write "          <tr class='tdbg'>" & vbCrLf
            Response.Write "           <td width='30' align='center'><input type='radio' name='CreateItemType" & i & "'  " & IsRadioChecked(CreateItemType, 2) & "  value='2'></td>"
            Response.Write "           <td height='30'> 生成最近" & vbCrLf
            Response.Write "            <Input name='CreateItemDate" & i & "' type='text' value='" & CreateItemDate & "' size=8 maxlength='10'> 天内的项目内容页" & vbCrLf
            Response.Write "           </td>" & vbCrLf
            Response.Write "         </tr>" & vbCrLf
            Response.Write "         <tr class='tdbg'>" & vbCrLf
            Response.Write "          <td width='30' align='center'><input type='radio' name='CreateItemType" & i & "'   " & IsRadioChecked(CreateItemType, 3) & "  value='3'></td>"
            Response.Write "          <td height='30'> 生成所有未生成的项目</td>" & vbCrLf
            Response.Write "         </tr>" & vbCrLf
            Response.Write "         <tr class='tdbg'>" & vbCrLf
            Response.Write "          <td width='30' align='center'><input type='radio' name='CreateItemType" & i & "'   " & IsRadioChecked(CreateItemType, 4) & "  value='4'></td>"
            Response.Write "          <td height='30'> 生成所有项目</td>" & vbCrLf
            Response.Write "         </tr>" & vbCrLf
            Response.Write "     </td>" & vbCrLf
            Response.Write "     </tr>" & vbCrLf
            Response.Write "       <tr class='tdbg'>" & vbCrLf
            Response.Write "         <td height='22' colspan='2' bgcolor='#E0EEF5'>&nbsp;&nbsp;栏目列表页生成操作 </td>" & vbCrLf
            Response.Write "       </tr>" & vbCrLf
            Response.Write "      <tr class='tdbg'>" & vbCrLf
            Response.Write "        <td width='30' align='center'><input type='checkbox' name='CreateClass" & i & "'  " & IsRadioChecked(CreateClass, "True") & " value='True'></td>"
            Response.Write "        <td height='30'>" & vbCrLf
            Response.Write "          生成所有栏目列表页" & vbCrLf
            Response.Write "        </td>" & vbCrLf
            Response.Write "      </tr>" & vbCrLf
            Response.Write "      <tr class='tdbg'>" & vbCrLf
            Response.Write "         <td height='22' colspan='2' bgcolor='#E0EEF5'>&nbsp;&nbsp;专题列表页生成操作 </td>" & vbCrLf
            Response.Write "      </tr>" & vbCrLf
            Response.Write "      <tr class='tdbg'>" & vbCrLf
            Response.Write "        <td width='30' align='center'><input type='checkbox' name='CreateSpecial" & i & "'  " & IsRadioChecked(CreateSpecial, "True") & " value='True'></td>"
            Response.Write "        <td height='30'>" & vbCrLf
            Response.Write "          生成所有专题列表页" & vbCrLf
            Response.Write "        </td>" & vbCrLf
            Response.Write "      </tr>" & vbCrLf
            Response.Write "      <tr class='tdbg'>" & vbCrLf
            Response.Write "        <td height='22' colspan='2' bgcolor='#E0EEF5'>&nbsp;&nbsp;频道页生成操作 </td>" & vbCrLf
            Response.Write "      </tr>" & vbCrLf
            Response.Write "      <tr class='tdbg'>" & vbCrLf
            Response.Write "        <td width='30' align='center'><input type='checkbox' name='CreateChannel" & i & "'   " & IsRadioChecked(CreateChannel, "True") & " value='True'></td>"
            Response.Write "        <td height='30'>" & vbCrLf
            Response.Write "          生成所选频道首页" & vbCrLf
            Response.Write "        </td>" & vbCrLf
            Response.Write "      </tr>" & vbCrLf
            Response.Write "      </table>" & vbCrLf
            Response.Write "      <INPUT TYPE='hidden' name='ChannelID" & i & "' value='" & rs("ChannelID") & "'>" & vbCrLf
            Response.Write "      <INPUT TYPE='hidden' name='ModuleType" & i & "' value='" & rs("ModuleType") & "'>" & vbCrLf
            rs.MoveNext
        Loop

    End If

    rs.Close
    Set rs = Nothing

    Response.Write "     </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "   </table>" & vbCrLf
    Response.Write "  </td>" & vbCrLf
    Response.Write " </tr>" & vbCrLf
    Response.Write " </tbody>" & vbCrLf
    Response.Write "    <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "      <tr height='30' class='tdbg'>"
    Response.Write "        <td class='tdbg5' align='right' width='100' >定时启动时间：&nbsp;&nbsp;</td>"
    Response.Write "        <td class='tdbg' >"
    Response.Write "         <select name='Timing_SetDate' id='SetTiming_Date' onchange=""javascript:ShowTimingType(this.options[this.selectedIndex].value)"">" & vbCrLf
    Response.Write "           <option value='0' " & IsOptionSelected(Timing_SetDate, 0) & ">每天</option>" & vbCrLf
    Response.Write "           <option value='1' " & IsOptionSelected(Timing_SetDate, 1) & ">每周</option>" & vbCrLf
    Response.Write "           <option value='2' " & IsOptionSelected(Timing_SetDate, 2) & ">每月</option>" & vbCrLf
    Response.Write "         </select>" & vbCrLf
    Response.Write "         <select name='Timing_SetWeekday' id='Timing_SetWeekday' " & IsStyleDisplay(Timing_SetDate, 1) & ">" & vbCrLf
    Response.Write "           <option value='1' " & IsOptionSelected(Timing_SetWeekday, 1) & ">星期日</option>" & vbCrLf
    Response.Write "           <option value='2' " & IsOptionSelected(Timing_SetWeekday, 2) & ">星期一</option>" & vbCrLf
    Response.Write "           <option value='3' " & IsOptionSelected(Timing_SetWeekday, 3) & ">星期二</option>" & vbCrLf
    Response.Write "           <option value='4' " & IsOptionSelected(Timing_SetWeekday, 4) & ">星期三</option>" & vbCrLf
    Response.Write "           <option value='5' " & IsOptionSelected(Timing_SetWeekday, 5) & ">星期四</option>" & vbCrLf
    Response.Write "           <option value='6' " & IsOptionSelected(Timing_SetWeekday, 6) & ">星期五</option>" & vbCrLf
    Response.Write "           <option value='7' " & IsOptionSelected(Timing_SetWeekday, 7) & ">星期六</option>" & vbCrLf
    Response.Write "         </select>" & vbCrLf
    Response.Write "         <select name='Timing_SetDay' " & IsStyleDisplay(Timing_SetDate, 2) & ">" & GetNumber_Option(1, 31, Timing_SetDay) & "</select>" & vbCrLf
    Response.Write "         <select name=""Timing_Time"" id=""Timing_Time"">" & vbCrLf
    Response.Write "           <option value=""00:00:00"" " & IsOptionSelected("00:00:00", Timing_Time) & ">00:00:00</option>" & vbCrLf
    Response.Write "           <option value=""00:30:00"" " & IsOptionSelected("00:30:00", Timing_Time) & ">00:30:00</option>" & vbCrLf
    Response.Write "           <option value=""01:00:00"" " & IsOptionSelected("01:00:00", Timing_Time) & ">01:00:00</option>" & vbCrLf
    Response.Write "           <option value=""01:30:00"" " & IsOptionSelected("01:30:00", Timing_Time) & ">01:30:00</option>" & vbCrLf
    Response.Write "           <option value=""02:00:00"" " & IsOptionSelected("02:00:00", Timing_Time) & ">02:00:00</option>" & vbCrLf
    Response.Write "           <option value=""02:30:00"" " & IsOptionSelected("02:30:00", Timing_Time) & ">02:30:00</option>" & vbCrLf
    Response.Write "           <option value=""03:00:00"" " & IsOptionSelected("03:00:00", Timing_Time) & ">03:00:00</option>" & vbCrLf
    Response.Write "           <option value=""03:30:00"" " & IsOptionSelected("03:30:00", Timing_Time) & ">03:30:00</option>" & vbCrLf
    Response.Write "           <option value=""04:00:00"" " & IsOptionSelected("04:00:00", Timing_Time) & ">04:00:00</option>" & vbCrLf
    Response.Write "           <option value=""04:30:00"" " & IsOptionSelected("04:30:00", Timing_Time) & ">04:30:00</option>" & vbCrLf
    Response.Write "           <option value=""05:00:00"" " & IsOptionSelected("05:00:00", Timing_Time) & ">05:00:00</option>" & vbCrLf
    Response.Write "           <option value=""05:30:00"" " & IsOptionSelected("05:30:00", Timing_Time) & ">05:30:00</option>" & vbCrLf
    Response.Write "           <option value=""06:00:00"" " & IsOptionSelected("06:00:00", Timing_Time) & ">06:00:00</option>" & vbCrLf
    Response.Write "           <option value=""06:30:00"" " & IsOptionSelected("06:30:00", Timing_Time) & ">06:30:00</option>" & vbCrLf
    Response.Write "           <option value=""07:00:00"" " & IsOptionSelected("07:00:00", Timing_Time) & ">07:00:00</option>" & vbCrLf
    Response.Write "           <option value=""07:30:00"" " & IsOptionSelected("07:30:00", Timing_Time) & ">07:30:00</option>" & vbCrLf
    Response.Write "           <option value=""08:00:00"" " & IsOptionSelected("08:00:00", Timing_Time) & ">08:00:00</option>" & vbCrLf
    Response.Write "           <option value=""08:30:00"" " & IsOptionSelected("08:30:00", Timing_Time) & ">08:30:00</option>" & vbCrLf
    Response.Write "           <option value=""09:00:00"" " & IsOptionSelected("09:00:00", Timing_Time) & ">09:00:00</option>" & vbCrLf
    Response.Write "           <option value=""09:30:00"" " & IsOptionSelected("09:30:00", Timing_Time) & ">09:30:00</option>" & vbCrLf
    Response.Write "           <option value=""10:00:00"" " & IsOptionSelected("10:00:00", Timing_Time) & ">10:00:00</option>" & vbCrLf
    Response.Write "           <option value=""10:30:00"" " & IsOptionSelected("10:30:00", Timing_Time) & ">10:30:00</option>" & vbCrLf
    Response.Write "           <option value=""11:00:00"" " & IsOptionSelected("11:00:00", Timing_Time) & ">11:00:00</option>" & vbCrLf
    Response.Write "           <option value=""11:30:00"" " & IsOptionSelected("11:30:00", Timing_Time) & ">11:30:00</option>" & vbCrLf
    Response.Write "           <option value=""12:00:00"" " & IsOptionSelected("12:00:00", Timing_Time) & ">12:00:00</option>" & vbCrLf
    Response.Write "           <option value=""12:30:00"" " & IsOptionSelected("12:30:00", Timing_Time) & ">12:30:00</option>" & vbCrLf
    Response.Write "           <option value=""13:00:00"" " & IsOptionSelected("13:00:00", Timing_Time) & ">13:00:00</option>" & vbCrLf
    Response.Write "           <option value=""13:30:00"" " & IsOptionSelected("13:30:00", Timing_Time) & ">13:30:00</option>" & vbCrLf
    Response.Write "           <option value=""14:00:00"" " & IsOptionSelected("14:00:00", Timing_Time) & ">14:00:00</option>" & vbCrLf
    Response.Write "           <option value=""14:30:00"" " & IsOptionSelected("14:30:00", Timing_Time) & ">14:30:00</option>" & vbCrLf
    Response.Write "           <option value=""15:00:00"" " & IsOptionSelected("15:00:00", Timing_Time) & ">15:00:00</option>" & vbCrLf
    Response.Write "           <option value=""15:30:00"" " & IsOptionSelected("15:30:00", Timing_Time) & ">15:30:00</option>" & vbCrLf
    Response.Write "           <option value=""16:00:00"" " & IsOptionSelected("16:00:00", Timing_Time) & ">16:00:00</option>" & vbCrLf
    Response.Write "           <option value=""16:30:00"" " & IsOptionSelected("16:30:00", Timing_Time) & ">16:30:00</option>" & vbCrLf
    Response.Write "           <option value=""17:00:00"" " & IsOptionSelected("17:00:00", Timing_Time) & ">17:00:00</option>" & vbCrLf
    Response.Write "           <option value=""17:30:00"" " & IsOptionSelected("17:30:00", Timing_Time) & ">17:30:00</option>" & vbCrLf
    Response.Write "           <option value=""18:00:00"" " & IsOptionSelected("18:00:00", Timing_Time) & ">18:00:00</option>" & vbCrLf
    Response.Write "           <option value=""18:30:00"" " & IsOptionSelected("18:30:00", Timing_Time) & ">18:30:00</option>" & vbCrLf
    Response.Write "           <option value=""19:00:00"" " & IsOptionSelected("19:00:00", Timing_Time) & ">19:00:00</option>" & vbCrLf
    Response.Write "           <option value=""19:30:00"" " & IsOptionSelected("19:30:00", Timing_Time) & ">19:30:00</option>" & vbCrLf
    Response.Write "           <option value=""20:00:00"" " & IsOptionSelected("20:00:00", Timing_Time) & ">20:00:00</option>" & vbCrLf
    Response.Write "           <option value=""20:30:00"" " & IsOptionSelected("20:30:00", Timing_Time) & ">20:30:00</option>" & vbCrLf
    Response.Write "           <option value=""21:00:00"" " & IsOptionSelected("21:00:00", Timing_Time) & ">21:00:00</option>" & vbCrLf
    Response.Write "           <option value=""21:30:00"" " & IsOptionSelected("21:30:00", Timing_Time) & ">21:30:00</option>" & vbCrLf
    Response.Write "           <option value=""22:00:00"" " & IsOptionSelected("22:00:00", Timing_Time) & ">22:00:00</option>" & vbCrLf
    Response.Write "           <option value=""22:30:00"" " & IsOptionSelected("22:30:00", Timing_Time) & ">22:30:00</option>" & vbCrLf
    Response.Write "           <option value=""23:00:00"" " & IsOptionSelected("23:00:00", Timing_Time) & ">23:00:00</option>" & vbCrLf
    Response.Write "           <option value=""23:30:00"" " & IsOptionSelected("23:30:00", Timing_Time) & ">23:30:00</option>" & vbCrLf
    Response.Write "         </select>" & vbCrLf
    Response.Write "        </td>" & vbCrLf
    Response.Write "      </tr>" & vbCrLf
    Response.Write "     </tbody></table>" & vbCrLf
    Response.Write "   </td>" & vbCrLf
    Response.Write " </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<table width='100%' border='0'>" & vbCrLf
    Response.Write "  <tr height='50'>" & vbCrLf
    Response.Write "     <td align='center'>" & vbCrLf
    Response.Write "       <input name='iChannelID' type='hidden' id='Action' value='" & i & "'>" & vbCrLf
    Response.Write "       <input type='submit' name='Submit' value='保存定时设置' onClick=""document.myform.Action.value='SaveTiming';"">" & vbCrLf
    Response.Write "       <input name='Action' type='hidden' id='Action' value='main'>" & vbCrLf
    Response.Write "     </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.Timing_CollectionItemID.length;i++){" & vbCrLf
    Response.Write "    document.myform.Timing_CollectionItemID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.Timing_CollectionItemID.length;i++){" & vbCrLf
    Response.Write "    document.myform.Timing_CollectionItemID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

'=================================================
'过程名：SaveTiming
'作  用：保存定时采集数据
'=================================================
Sub SaveTiming()
    
    Dim rs, sql
    Dim iChannelID, i, CreateItemType, CreateItemTopNewNum, CreateItemDate, CreateClass, CreateSpecial, CreateChannel
    Dim Timing_CollectionItemID, Timing_Time, Timing_Renovate, Timing_Passed, Timing_SetDate, Timing_SetWeekday, Timing_SetDay
    Dim TimingCreateSetting, Timing_ClsDate, Timing_AreaCollection

    '定时属性保存
    Timing_CollectionItemID = Trim(Request("Timing_CollectionItemID"))
    Timing_Time = Trim(Request("Timing_Time"))
    Timing_Renovate = PE_CLng(Trim(Request("Timing_Renovate")))
    Timing_Passed = Trim(Request("Timing_Passed"))
    Timing_ClsDate = Trim(Request("Timing_ClsDate"))
    Timing_AreaCollection = Trim(Request("Timing_AreaCollection"))

    Timing_SetDate = PE_CLng(Trim(Request("Timing_SetDate")))
    Timing_SetWeekday = PE_CLng(Trim(Request("Timing_SetWeekday")))
    Timing_SetDay = PE_CLng(Trim(Request("Timing_SetDay")))

    If IsNull(Timing_CollectionItemID) = True Or Timing_CollectionItemID = "" Then
        Timing_CollectionItemID = "0"
    Else
        Timing_CollectionItemID = ReplaceBadChar(Timing_CollectionItemID)
    End If

    If Timing_Passed = "yes" Then
        Timing_Passed = True
    Else
        Timing_Passed = False
    End If

    If Timing_ClsDate = "yes" Then
        If SystemDatabaseType = "SQL" Then
            Conn.Execute "update PE_Config set Timing_Date='" & Date & "'"
        Else
            Conn.Execute "update PE_Config set Timing_Date=#" & Date & "#"
        End If
    End If

    Call PE_Cache.DelAllCache

    Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from PE_Config"
    rs.Open sql, Conn, 1, 3

    If rs.EOF Or rs.BOF Then
        Response.Write "您还没有添入系统设置表,请运行系统初始化。"
        Response.End
    Else
        rs("Timing_AreaCollection") = Timing_AreaCollection
        rs("Timing_CollectionItemID") = Timing_CollectionItemID
        rs("Timing_Time") = Timing_Time
        rs("Timing_SetDate") = Timing_SetDate
        rs("Timing_SetWeekday") = Timing_SetWeekday
        rs("Timing_SetDay") = Timing_SetDay
        rs.Update
    End If

    rs.Close
    Set rs = Nothing
        
    '定时生成保存
    iChannelID = PE_CLng(Trim(Request("iChannelID")))

    For i = 1 To iChannelID
        ChannelID = PE_CLng(Trim(Request("ChannelID" & i)))
        ModuleType = PE_CLng(Trim(Request("ModuleType" & i)))
        CreateItemType = PE_CLng(Trim(Request("CreateItemType" & i)))
        CreateItemTopNewNum = PE_CLng(Trim(Request("CreateItemTopNewNum" & i)))
        CreateItemDate = PE_CLng(Trim(Request("CreateItemDate" & i)))
        CreateClass = Trim(Request("CreateClass" & i))
        CreateSpecial = Trim(Request("CreateSpecial" & i))
        CreateChannel = Trim(Request("CreateChannel" & i))

        If CreateItemType = 1 Then
            If CreateItemTopNewNum = 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>请选择您要最新生成的数量</li>"
            End If

        ElseIf CreateItemType = 2 Then

            If CreateItemDate = 0 Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>请选择您最近生成的天数</li>"
            End If
        End If

        If FoundErr = True Then
            Call WriteErrMsg(ErrMsg, ComeUrl)
            Exit Sub
        End If

        Set rs = Server.CreateObject("adodb.recordset")
        sql = "select top 1 * from PE_Channel where ChannelID=" & ChannelID
        rs.Open sql, Conn, 1, 3
        rs("TimingCreateSetting") = ChannelID & "," & ModuleType & "," & CreateItemType & "," & CreateItemTopNewNum & "," & CreateItemDate & "," & CreateClass & "," & CreateSpecial & "," & CreateChannel
        rs.Update
        rs.MoveNext
        rs.Close
        Set rs = Nothing
    Next

    Call WriteSuccessMsg("定时管理设置成功！", ComeUrl)
End Sub
'=================================================
'过程名：DoMainTiming
'作  用：定时采集主框架
'=================================================
Sub DoMainTiming()

    Response.Write "        <script language=""JavaScript"">" & vbCrLf
    Response.Write "        <!--" & vbCrLf
    Response.Write "        function Timing_Time(Timing_AreaCollection,CollectionItemID,TimingCreate){" & vbCrLf
    Response.Write "                objFiles.innerHTML= ""<iframe marginwidth=0 marginheight=0 frameborder=0 name='libin' width='100%' height='100%' src='Admin_Collection.asp?Action=Start&ItemID=""+CollectionItemID+""&ItemNum=1&ListNum=1&Arr_i=0&CollecNewsA=0&CollecNewsi=0&ItemSucceedNum=0&ItemSucceedNum2=0&CollecNewsj=0&ImagesNumAll=0&ItemIDtemp=0&CollecType=1&Content_object=1&Timing_AreaCollection=""+Timing_AreaCollection+""&TimingCreate=""+TimingCreate+""'></iframe>"";" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        //-->" & vbCrLf
    Response.Write "        </script>" & vbCrLf
    Response.Write "        <table border='0' cellpadding='0' cellspacing='0' width='100%' height='100%' align='center'>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "           <td valign='top' height='30%'>"
    Response.Write "            <iframe marginwidth=0 marginheight=0 frameborder=0 name=""libin"" width=""100%"" height=""100%"" src=""Admin_Timing.asp?Action=DoTiming2""></iframe>"
    Response.Write "           </td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "           <td valign='top' id='objFiles' width='70%'></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
End Sub
'=================================================
'过程名：DoTiming2
'作  用：定时启动项目
'=================================================
Sub DoTiming2()

    Dim rs, sql
    Dim rnd_temp
    Dim Timing_AreaCollection, Timing_CollectionItemID, CollectionItemName, Timing_SetDate, Timing_SetWeekday, Timing_SetDay, Timing_Time, Timing_Passed, Timing_Date, Timing_Renovate
    Dim arrChannelID, i, CreateItemType, CreateItemTopNewNum, CreateItemDate, CreateClass, CreateSpecial, CreateChannel
    Dim TimingCreate, TimingCreateNum, CreateChannelName
    Dim Timing_Startup

    Timing_Startup = False
    rnd_temp = Trim(Request("rnd_temp"))

    If PE_Cache.CacheIsEmpty("CollectionItemName" & rnd_temp) Then
        '清除缓存
        Call PE_Cache.DelAllCache
        '产生5位随机数
        rnd_temp = CStr(rnd_num(5))

        '加载定时记录
        sql = "select Timing_AreaCollection,Timing_CollectionItemID,Timing_SetDate,Timing_SetWeekday,Timing_SetDay,Timing_Time,Timing_Date from PE_Config"
        Set rs = Server.CreateObject("adodb.recordset")
        rs.Open sql, Conn, 1, 1

        If Not rs.EOF Then
            Timing_AreaCollection = rs("Timing_AreaCollection")
            Timing_CollectionItemID = rs("Timing_CollectionItemID")
            Timing_SetDate = rs("Timing_SetDate")
            Timing_SetWeekday = rs("Timing_SetWeekday")
            Timing_SetDay = rs("Timing_SetDay")
            Timing_Time = rs("Timing_Time")
            Timing_Date = rs("Timing_Date")
        End If

        rs.Close
        Set rs = Nothing

        If IsNull(Timing_CollectionItemID) = True Or Timing_CollectionItemID = "" Or IsValidID(Timing_CollectionItemID) = False Then
            Timing_CollectionItemID = "0"
        End If

        If Timing_CollectionItemID = "0" Then
        Else
            '获得采集项目名称
            sql = "select ItemName from PE_Item where ItemID"

            If InStr(Timing_CollectionItemID, ",") > 0 Then
                sql = sql & " in (" & Timing_CollectionItemID & ")"
            Else
                sql = sql & " =" & Timing_CollectionItemID
            End If

            sql = sql & " and Flag=" & PE_True
            Set rs = Server.CreateObject("adodb.recordset")
            rs.Open sql, Conn, 1, 1

            If rs.EOF And rs.BOF Then
            Else

                Do While Not rs.EOF

                    If CollectionItemName = "" Then
                        CollectionItemName = rs("ItemName")
                    Else
                        CollectionItemName = CollectionItemName & "," & rs("ItemName")
                    End If

                    rs.MoveNext
                Loop

            End If

            rs.Close
            Set rs = Nothing
        End If
        '获得频道项目名称
        sql = "select ChannelName,Disabled,UseCreateHTML,TimingCreateSetting from PE_Channel  where  ModuleType<>0 and ModuleType<>4 and Disabled=" & PE_False & " and UseCreateHTML > 0  order by ChannelID asc"
        Set rs = Server.CreateObject("adodb.recordset")
        rs.Open sql, Conn, 1, 1

        If rs.EOF And rs.BOF Then
        Else

            Do While Not rs.EOF
                TimingCreateNum = TimingCreateNum + 1

                If TimingCreate = "" Then
                    TimingCreate = rs("TimingCreateSetting")
                Else
                    TimingCreate = TimingCreate & "$" & rs("TimingCreateSetting")
                End If

                If CreateChannelName = "" Then
                    CreateChannelName = rs("ChannelName")
                Else
                    CreateChannelName = CreateChannelName & "," & rs("ChannelName")
                End If

                rs.MoveNext
            Loop

        End If

        rs.Close
        Set rs = Nothing

        '加载缓存
        
        PE_Cache.SetValue "Timing_AreaCollection" & rnd_temp, Timing_AreaCollection
        PE_Cache.SetValue "Timing_CollectionItemID" & rnd_temp, Timing_CollectionItemID
        PE_Cache.SetValue "CollectionItemName" & rnd_temp, CollectionItemName
        PE_Cache.SetValue "Timing_Date" & rnd_temp, Timing_Date
        PE_Cache.SetValue "Timing_SetDate" & rnd_temp, Timing_SetDate
        PE_Cache.SetValue "Timing_SetWeekday" & rnd_temp, Timing_SetWeekday
        PE_Cache.SetValue "Timing_SetDay" & rnd_temp, Timing_SetDay
        PE_Cache.SetValue "Timing_Time" & rnd_temp, Timing_Time
        PE_Cache.SetValue "TimingCreate" & rnd_temp, TimingCreate
        PE_Cache.SetValue "CreateChannelName" & rnd_temp, CreateChannelName
        
    End If

    Timing_AreaCollection = PE_Cache.GetValue("Timing_AreaCollection" & rnd_temp)
    Timing_CollectionItemID = PE_Cache.GetValue("Timing_CollectionItemID" & rnd_temp)
    CollectionItemName = PE_Cache.GetValue("CollectionItemName" & rnd_temp)
    Timing_Date = PE_Cache.GetValue("Timing_Date" & rnd_temp)
    Timing_SetDate = PE_Cache.GetValue("Timing_SetDate" & rnd_temp)
    Timing_SetWeekday = PE_Cache.GetValue("Timing_SetWeekday" & rnd_temp)
    Timing_SetDay = PE_Cache.GetValue("Timing_SetDay" & rnd_temp)
    Timing_Time = PE_Cache.GetValue("Timing_Time" & rnd_temp)
    TimingCreate = PE_Cache.GetValue("TimingCreate" & rnd_temp)
    CreateChannelName = PE_Cache.GetValue("CreateChannelName" & rnd_temp)

    If IsDate(Timing_Time) = False Then
        Call WriteErrMsg("<li>请先设置定时项目的时间，再运行定时。", ComeUrl)
        Exit Sub
    End If

    If CollectionItemName = "" Then
        CollectionItemName = "还没有选择要定时的采集项目!"
    End If

    If CreateChannelName = "" Then
        CreateChannelName = "还没有选择要定时的生成频道!"
    End If

    If IsDate(Timing_Date) = False Or Timing_Date > Date + 1 Then
        Timing_Date = Date - 1
        PE_Cache.SetValue "Timing_Date" & rnd_temp, Timing_Date
    End If

    If Timing_SetDate = 1 Then
        If Timing_SetWeekday = Weekday(Date) Then
            Timing_Startup = True
        End If

    ElseIf Timing_SetDate = 2 Then

        If Timing_SetDay = Day(Date) Then
            Timing_Startup = True
        End If

    ElseIf Timing_SetDate = 0 Then
        Timing_Startup = True
    End If

    '系统当前时间 > 设置启动时间 And 系统当前时间 < 设置启动时间+30分 And 系统当前日期 > 记录日期　And　时间类型为真 Then
    If CDate(Hour(Time) & ":" & Minute(Time) & ":" & Second(Time)) > CDate(Timing_Time) And CDate(Hour(Time) & ":" & Minute(Time) & ":" & Second(Time)) < CDate(Hour(Timing_Time) & ":" & Minute(Timing_Time) + 29 & ":" & Second(Timing_Time) + 59) And Date > Timing_Date And Timing_Startup = True Then

        '加载定时记录
        If SystemDatabaseType = "SQL" Then
            Conn.Execute "update PE_Config set Timing_Date='" & Date & "'"
        Else
            Conn.Execute "update PE_Config set Timing_Date=#" & Date & "#"
        End If
        
        PE_Cache.DelCache "Timing_AreaCollection" & rnd_temp
        PE_Cache.DelCache "Timing_CollectionItemID" & rnd_temp
        PE_Cache.DelCache "CollectionItemName" & rnd_temp
        PE_Cache.DelCache "Timing_Date" & rnd_temp
        PE_Cache.DelCache "Timing_SetDate" & rnd_temp
        PE_Cache.DelCache "Timing_SetWeekday" & rnd_temp
        PE_Cache.DelCache "Timing_SetDay" & rnd_temp
        PE_Cache.DelCache "Timing_Time" & rnd_temp
        PE_Cache.DelCache "TimingCreate" & rnd_temp
        PE_Cache.DelCache "CreateChannelName" & rnd_temp

        rnd_temp = ""
        Response.Write "<script language=""JavaScript"">" & vbCrLf
        Response.Write "<!--" & vbCrLf
        Response.Write "    parent.Timing_Time('" & Timing_AreaCollection & "','" & Timing_CollectionItemID & "','" & TimingCreate & "');" & vbCrLf
        Response.Write "//-->" & vbCrLf
        Response.Write "</script>" & vbCrLf
        Response.Write "<center><FONT style='font-size:12px' color='red'>请稍等，系统正在执行定时项目。</FONT></center>"
        Call Refresh("Admin_Timing.asp?Action=DoTiming2&rnd_temp=" & rnd_temp,10)
        'Response.Write "<meta http-equiv=""refresh"" content=10;url=""Admin_Timing.asp?Action=DoTiming2&rnd_temp=" & rnd_temp & """>"
    Else
        Response.Write "<br>"
        Response.Write "&nbsp;&nbsp;您要定时采集项目是：<FONT style='font-size:12px' color='red'>" & CollectionItemName & "</FONT><br>"
        If Timing_AreaCollection = "1" Then
            Response.Write "&nbsp;&nbsp;定时区域采集：<FONT style='font-size:12px' color='red'>启动</font><br>"
        End If
        Response.Write "&nbsp;&nbsp;您要定时生成项目是：<FONT style='font-size:12px' color='red'>" & CreateChannelName & "</FONT><br>"
        Response.Write "&nbsp;&nbsp;当前的日期是：<FONT style='font-size:12px' color='red'>" & Date & " </FONT><br>"
        Response.Write "&nbsp;&nbsp;您指定的时间是"

        If Timing_SetDate = 1 Then
            Response.Write "每周的<FONT style='font-size:12px' color='red'>星期"

            Select Case Timing_SetWeekday

                Case 1
                    Response.Write "日"

                Case 2
                    Response.Write "一"

                Case 3
                    Response.Write "二"

                Case 4
                    Response.Write "三"

                Case 5
                    Response.Write "四"

                Case 6
                    Response.Write "五"

                Case 7
                    Response.Write "六"
            End Select

            Response.Write "</FONT> "
        ElseIf Timing_SetDate = 2 Then
            Response.Write "每月的<FONT style='font-size:12px' color='red'>" & Timing_SetDay & "</FONT>日 "
        ElseIf Timing_SetDate = 0 Then
            Response.Write "每天的"
        End If

        Response.Write "<FONT style='font-size:12px' color='red'>" & Timing_Time & " </FONT><br>"
        Response.Write "&nbsp;&nbsp;页面每 10 秒刷新一次"
        Response.Write "&nbsp;&nbsp;<center><FONT style='font-size:12px' color='red'>定时项目已经开启,您可以切换浏览器去做其他工作,记得在您临走时为了安全记得开启 windows 安全认证。</FONT></center><br>"
        Call Refresh("Admin_Timing.asp?Action=DoTiming2&rnd_temp=" & rnd_temp,10)		
        'Response.Write "<meta http-equiv=""refresh"" content=10;url=""Admin_Timing.asp?Action=DoTiming2&rnd_temp=" & rnd_temp & """>"
    End If

End Sub
'=================================================
'过程名：DoTiming
'作  用：定时生成HTML
'=================================================
Sub DoTiming()
    Dim TimingCreate, arrTimingCreate, TimingCreateNum, CreateChannelItem, TimingCreateUrl
    Dim CreateChannelType, CreateActionType, CreateType
    
    TimingCreate = Trim(Request("TimingCreate"))
    TimingCreateNum = PE_CLng(Trim(Request("TimingCreateNum")))

    If TimingCreate = "" Then
        Exit Sub
    End If

    arrTimingCreate = Split(TimingCreate, "$")

    If TimingCreateNum > UBound(arrTimingCreate) Then
        Response.Write "<html><head><title>成功信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
        Response.Write "<link href='images/Style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbCrLf
        Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
        Response.Write "  <tr align='center' class='title'><td height='22'><strong>恭喜您！</strong></td></tr>" & vbCrLf
        Response.Write "  <tr class='tdbg'><td height='100' valign='top' align='center'><br><font color=red>" & Date & "</font>&nbsp;定时生成执行完成!</td></tr>" & vbCrLf
        Response.Write "  <tr align='center' class='tdbg'><td>"
        Response.Write "</td></tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
        Exit Sub
    End If

    CreateChannelItem = Split(arrTimingCreate(TimingCreateNum), ",")

    '1--文章  2--下载  3--图片  5--商城
    If CreateChannelItem(1) = 1 Then
        CreateChannelType = "Article"
        CreateActionType = "CreateArticle"
    ElseIf CreateChannelItem(1) = 2 Then
        CreateChannelType = "Soft"
        CreateActionType = "CreateSoft"
    ElseIf CreateChannelItem(1) = 3 Then
        CreateChannelType = "Photo"
        CreateActionType = "CreatePhoto"
    ElseIf CreateChannelItem(1) = 5 Then
        CreateChannelType = "Product"
        CreateActionType = "CreateProduct"
    End If

    If PE_CLng(CreateChannelItem(2)) = 0 Then
        CreateActionType = "CreateOther"
    End If

    TimingCreateUrl = "Admin_Create" & CreateChannelType & ".asp?Action=" & CreateActionType & "&CreateType=8&CreateItemType=" & CreateChannelItem(2) & "&ChannelProperty=" & arrTimingCreate(TimingCreateNum) & "&TimingCreateNum=" & TimingCreateNum & "&ChannelID=" & CreateChannelItem(0) & "&ClassID=1&TotalCreate=20&ShowBack=No&TimingCreate=" & TimingCreate

    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "function aaa(){window.location.href='" & TimingCreateUrl & "';}" & vbCrLf
    Response.Write "    setTimeout('aaa()',5000);" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub
'*************************  类模块主区域结束  *******************************
'*************************  类模块函数通用开始  *****************************
'=================================================
'过程名：rnd_num
'作  用：产生指定位置的随机数
'参  数：产生的随机数  ----内容
'=================================================
Function rnd_num(rLen)
    Dim ri, rmax, rmin
    rmax = 1
    rmin = 1
    For ri = 1 To rLen + 1
        rmax = rmax * 10
    Next
    rmax = rmax - 1
    For ri = 1 To rLen
        rmin = rmin * 10
    Next
    Randomize
    rnd_num = Int((rnd_num - rmin + 1) * Rnd) + rmin
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
