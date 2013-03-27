<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<%
Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 1      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

Dim dbpath, Barwidth

dbpath = Server.MapPath(DBFileName)

Barwidth = 500
Response.Write "<html><head><title>数据库管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("数 据 库 管 理", 10009)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30' ><strong>管理导航：</strong></td><td>"
Response.Write "<a href='Admin_Database.asp?Action=Backup'>备份数据库</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Database.asp?Action=Restore'>恢复数据库</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Database.asp?Action=Compact'>压缩数据库</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Database.asp?Action=Init'>系统初始化</a>&nbsp;|&nbsp;"
Response.Write "<a href='Admin_Database.asp?Action=SpaceSize'>系统空间占用情况</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case Action
Case "Backup"
    Call ShowBackup
Case "BackupData"
    Call BackupData
Case "Compact"
    Call ShowCompact
Case "CompactData"
    Call CompactData
Case "Restore"
    Call ShowRestore
Case "RestoreData"
    Call RestoreData
Case "Init"
    Call ShowInit
Case "Clear"
    Call ShowInit
Case "SpaceSize"
    Call SpaceSize
Case Else
    FoundErr = True
    ErrMsg = ErrMsg & "<li>错误参数！</li>"
End Select
If FoundErr = True Then
    Call WriteEntry(2, AdminName, "数据管理操作失败，失败原因：" & ErrMsg)
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub ShowBackup()
    Response.Write "<form method='post' action='Admin_Database.asp?action=BackupData'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "      <td align='center' height='22' valign='middle'><b>备 份 数 据 库</b></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='150' align='center' valign='middle'>"
    Response.Write "<table cellpadding='3' cellspacing='1' border='0' width='100%'>"
    Response.Write "  <tr>"
    Response.Write " <td width='200' height='33' align='right'>备份目录：</td>"
    Response.Write " <td><input type=text size=20 name=bkfolder value=Databackup></td>"
    Response.Write " <td>相对路径目录，如目录不存在，将自动创建</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write " <td width='200' height='34' align='right'>备份名称：</td>"
    Response.Write " <td height='34'><input type=text size=20 name=bkDBname value='" & Date & "'></td>"
    Response.Write " <td height='34'>不用输入文件名后缀（默认为“.asa”）。如有同名文件，将覆盖</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr align='center'>"
    Response.Write " <td height='40' colspan='3'><input name='submit' type=submit value=' 开始备份 '"
    If SystemDatabaseType = "SQL" Or ObjInstalled_FSO = False Then
        Response.Write " disabled"
    End If
    Response.Write "></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    If ObjInstalled_FSO = False Then
        Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    If SystemDatabaseType = "SQL" Then
        Response.Write "<br><b>说明：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;您使用的是SQL版，请直接使用SQL2000提供的数据库备份功能进行备份！<br><br>"
    End If
End Sub

Sub ShowCompact()
    Response.Write "<form method='post' action='Admin_Database.asp?action=CompactData'>"
    Response.Write "<table class='border' width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write " <tr class='title'>"
    Response.Write "     <td align='center' height='22' valign='middle'><b>数据库在线压缩</b></td>"
    Response.Write " </tr>"
    Response.Write " <tr class='tdbg'>"
    Response.Write "     <td align='center' height='150' valign='middle'>"
    Response.Write "      <br>"
    Response.Write "      <br>"
    Response.Write "      压缩前，建议先备份数据库，以免发生意外错误。 <br>"
    Response.Write "      <br>"
    Response.Write "      <br>"
    Response.Write " <input name='submit' type=submit value=' 压缩数据库 '"
    If SystemDatabaseType = "SQL" Then
        Response.Write " disabled"
    End If
    Response.Write "><br><br>"
    If ObjInstalled_FSO = False Or ObjInstalled_FSO = False Then
        Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    If SystemDatabaseType = "SQL" Then
        Response.Write "<br><b>说明：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;您使用的是SQL版，无需进行压缩操作！<br><br>"
    End If
End Sub

Sub ShowRestore()
    Response.Write "<form method='post' action='Admin_Database.asp?action=RestoreData'>"
    Response.Write "<table width='100%' class='border' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td align='center' height='22' valign='middle'><b>数据库恢复</b></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td align='center' height='150' valign='middle'>"
    Response.Write "      <table width='100%' border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "        <tr>"
    Response.Write "          <td width='200' height='30' align='right'>原备份数据库路径（相对）：</td>"
    Response.Write "          <td height='30'><input name=backpath type=text id='backpath' value='Databackup\PowerEasy.asa' size=50 maxlength='200'></td>"
    Response.Write "        </tr>"
    Response.Write "        <tr align='center'>"
    Response.Write "          <td height='40' colspan='2'><input name='submit' type=submit value=' 恢复数据 '"
    If SystemDatabaseType = "SQL" Or ObjInstalled_FSO = False Then
        Response.Write " disabled"
    End If
    Response.Write ">"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    If ObjInstalled_FSO = False Then
        Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    If SystemDatabaseType = "SQL" Then
        Response.Write "<br><b>说明：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;您使用的是SQL版，请直接使用SQL2000提供的数据库恢复功能进行恢复！<br><br>"
    Else
        Response.Write "<br><b>说明：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;原备份数据库的扩展名必须为：asa或者asp<br><br>"
    End If
End Sub

Sub ShowInit()
    Dim ChannelTable, rsChannel, sqlChannel
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if(confirm('确实要清除选定的表吗？一旦清除将无法恢复！'))" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "      if (document.myform.PE_User.checked==true)" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "           if(confirm('您选择了清除会员数据，如果本系统的会员和其他系统共用数据库，则一旦清除将无法恢复！'))" & vbCrLf
    Response.Write "             return true;" & vbCrLf
    Response.Write "           else" & vbCrLf
    Response.Write "             return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "      else" & vbCrLf
    Response.Write "         return true;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  else" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnCheckChannel(){" & vbCrLf
    Response.Write "  if(document.myform.chkChannel.checked){" & vbCrLf
    Response.Write "    document.myform.chkChannel.checked = document.myform.chkChannel.checked&0;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckChannel(form){" & vbCrLf
    Response.Write "  for (var i=0;i<form.elements.length;i++){" & vbCrLf
    Response.Write "    var e = form.elements[i];" & vbCrLf
    Response.Write "    if (e.name){" & vbCrLf
    Response.Write "      if (e.name.substr(0,5) == 'C_PE_' && e.disabled==false)" & vbCrLf
    Response.Write "         e.checked = form.chkChannel.checked;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnCheckShop(){" & vbCrLf
    Response.Write "  if(document.myform.chkShop.checked){" & vbCrLf
    Response.Write "    document.myform.chkShop.checked = document.myform.chkShop.checked&0;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckShop(form){" & vbCrLf
    Response.Write "  for (var i=0;i<form.elements.length;i++){" & vbCrLf
    Response.Write "    var e = form.elements[i];" & vbCrLf
    Response.Write "    if (e.name){" & vbCrLf
    Response.Write "      if (e.name.substr(0,5) == 'S_PE_' && e.disabled==false)" & vbCrLf
    Response.Write "         e.checked = form.chkShop.checked;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function UnCheckJob(){" & vbCrLf
    Response.Write "  if(document.myform.chkJob.checked){" & vbCrLf
    Response.Write "    document.myform.chkJob.checked = document.myform.chkJob.checked&0;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckJob(form){" & vbCrLf
    Response.Write "  for (var i=0;i<form.elements.length;i++){" & vbCrLf
    Response.Write "    var e = form.elements[i];" & vbCrLf
    Response.Write "    if (e.name){" & vbCrLf
    Response.Write "      if (e.name.substr(0,2) == 'J_' && e.disabled==false)" & vbCrLf
    Response.Write "         e.checked = form.chkJob.checked;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnCheckHouse(){" & vbCrLf
    Response.Write "  if(document.myform.chkHouse.checked){" & vbCrLf
    Response.Write "    document.myform.chkHouse.checked = document.myform.chkHouse.checked&0;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckHouse(form){" & vbCrLf
    Response.Write "  for (var i=0;i<form.elements.length;i++){" & vbCrLf
    Response.Write "    var e = form.elements[i];" & vbCrLf
    Response.Write "    if (e.name){" & vbCrLf
    Response.Write "      if (e.name.substr(0,2) == 'H_' && e.disabled==false)" & vbCrLf
    Response.Write "         e.checked = form.chkHouse.checked;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnCheckOther(){" & vbCrLf
    Response.Write "  if(document.myform.chkOther.checked){" & vbCrLf
    Response.Write "    document.myform.chkOther.checked = document.myform.chkOther.checked&0;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function CheckOther(form){" & vbCrLf
    Response.Write "  for (var i=0;i<form.elements.length;i++){" & vbCrLf
    Response.Write "    var e = form.elements[i];" & vbCrLf
    Response.Write "    if (e.name){" & vbCrLf
    Response.Write "      if (e.name.substr(0,3) == 'PE_' && e.disabled==false)" & vbCrLf
    Response.Write "         e.checked = form.chkOther.checked;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<form action='Admin_Database.asp' method='post' name='myform' id='myform' onSubmit='return CheckForm();'>"
    Response.Write "<table class='border' width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td align='center' height='22' valign='middle'><b>系 统 初 始 化</b></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='100%' height='150' align=center valign='middle'>"
    If Action = "Clear" Then
        Response.Write "      <div align='left'>"
        Call ClearData
        Response.Write "      </div>"
    Else
        Response.Write "      <table border='0' cellspacing='0' cellpadding='5'>"
        Response.Write "        <tr>"
        Response.Write "          <td>"
        Response.Write "            <fieldset name='ChannelData'><legend>频道数据 <input name='chkChannel' type='checkbox' id='chkChannel' onclick='CheckChannel(this.form)' value=''></legend><table width='600' border='0' cellpadding='0' cellspacing='5'>"

        sqlChannel = "select * from PE_Channel where ChannelType<=1 and ChannelID<>4 order by OrderID"
        Set rsChannel = Conn.Execute(sqlChannel)
        Do While Not rsChannel.EOF
            Select Case rsChannel("ModuleType")
            Case 1
                ChannelTable = "PE_Article"
            Case 2
                ChannelTable = "PE_Soft"
            Case 3
                ChannelTable = "PE_Photo"
            Case 5
                ChannelTable = "PE_Product"
            End Select
            Response.Write "              <tr>"
            Response.Write "                <td width='20%'><input name='C_PE_Class_" & rsChannel("ChannelID") & "' type='checkbox' id='C_PE_Class_" & rsChannel("ChannelID") & "' onclick='UnCheckChannel()' value='yes'> " & rsChannel("ChannelName") & "栏目</td>"
            Response.Write "                <td width='20%'><input name='C_PE_Special_" & rsChannel("ChannelID") & "' type='checkbox' id='C_PE_Special_" & rsChannel("ChannelID") & "' onclick='UnCheckChannel()' value='yes'> " & rsChannel("ChannelName") & "专题</td>"
            Response.Write "                <td width='20%'><input name='C_" & ChannelTable & "_" & rsChannel("ChannelID") & "' type='checkbox' id='C_" & ChannelTable & "_" & rsChannel("ChannelID") & "' onclick='UnCheckChannel()' value='yes'> " & rsChannel("ChannelName") & "数据</td>"
            Response.Write "                <td width='20%'><input name='C_PE_Comment_" & rsChannel("ChannelID") & "' type='checkbox' id='C_PE_Comment_" & rsChannel("ChannelID") & "' onclick='UnCheckChannel()' value='yes'> " & rsChannel("ChannelName") & "评论</td>"
            Response.Write "                <td width='20%'><input name='C_PE_JsFile_" & rsChannel("ChannelID") & "' type='checkbox' id='C_PE_JsFile_" & rsChannel("ChannelID") & "' onclick='UnCheckChannel()' value='yes'> " & rsChannel("ChannelName") & "JS数据</td>"
            Response.Write "              </tr>"
            rsChannel.MoveNext
        Loop
        rsChannel.Close
        Set rsChannel = Nothing
        Response.Write "            </table></fieldset>"
        Response.Write "          </td>"
        Response.Write "        </tr>"
        Response.Write "        <tr>"
        Response.Write "          <td>"
        Response.Write "            <fieldset name='ShoprData'><legend>商城数据 <input name='chkShop' type='checkbox' id='chkShop' onclick='CheckShop(this.form)' value=''></legend><table width='600' border='0' cellpadding='0' cellspacing='5'>"
        Response.Write "              <tr>"
        Response.Write "                <td width='20%'><input name='S_PE_OrderForm' type='checkbox' id='S_PE_OrderForm' onclick='UnCheckShop()' value='yes'> 订单数据</td>"
        Response.Write "                <td width='20%'><input name='S_PE_Bank' type='checkbox' id='S_PE_Bank' onclick='UnCheckShop()' value='yes'> 银行帐户</td>"
        Response.Write "                <td width='20%'><input name='S_PE_BankrollItem' type='checkbox' id='S_PE_BankrollItem' onclick='UnCheckShop()' value='yes'> 资金记录</td>"
        Response.Write "                <td width='20%'><input name='S_PE_DeliverItem' type='checkbox' id='S_PE_DeliverItem' onclick='UnCheckShop()' value='yes'> 发退货记录</td>"
        Response.Write "                <td width='20%'><input name='S_PE_Payment' type='checkbox' id='S_PE_Payment' onclick='UnCheckShop()' value='yes'> 在线支付记录</td>"
        Response.Write "              </tr>"
        Response.Write "              <tr>"
        Response.Write "                <td width='20%'><input name='S_PE_DeliverType' type='checkbox' id='S_PE_DeliverType' onclick='UnCheckShop()' value='yes'> 送货方式</td>"
        Response.Write "                <td width='20%'><input name='S_PE_PaymentType' type='checkbox' id='S_PE_PaymentType' onclick='UnCheckShop()' value='yes'> 付款方式</td>"
        Response.Write "                <td width='20%'><input name='S_PE_PresentProject' type='checkbox' id='S_PE_PresentProject' onclick='UnCheckShop()' value='yes'> 促销方案</td>"
        Response.Write "                <td width='20%'><input name='S_PE_Producer' type='checkbox' id='S_PE_Producer' onclick='UnCheckShop()' value='yes'> 生 产 商</td>"
        Response.Write "                <td width='20%'><input name='S_PE_Trademark' type='checkbox' id='S_PE_Trademark' onclick='UnCheckShop()' value='yes'> 商品品牌</td>"
        Response.Write "              </tr>"
        Response.Write "              <tr>"
        Response.Write "                <td width='20%'><input name='S_PE_Client' type='checkbox' id='S_PE_Client' onclick='UnCheckShop()' value='yes'> 客户信息</td>"
        Response.Write "                <td width='20%'><input name='S_PE_Company' type='checkbox' id='S_PE_Company' onclick='UnCheckShop()' value='yes'> 企业信息</td>"
        Response.Write "                <td width='20%'><input name='S_PE_Contacter' type='checkbox' id='S_PE_Contacter' onclick='UnCheckShop()' value='yes'> 联系人信息</td>"
        Response.Write "                <td width='20%'><input name='S_PE_ServiceItem' type='checkbox' id='S_PE_ServiceItem' onclick='UnCheckShop()' value='yes'> 服务记录</td>"
        Response.Write "                <td width='20%'><input name='S_PE_ComplainItem' type='checkbox' id='S_PE_ComplainItem' onclick='UnCheckShop()' value='yes'> 投诉记录</td>"
        Response.Write "              </tr>"
        Response.Write "            </table></fieldset>"
        Response.Write "          </td>"
        Response.Write "        </tr>"

        Response.Write "        <tr>"
        Response.Write "          <td>"
        Response.Write "            <fieldset name='ShoprData'><legend>招聘数据 <input name='chkJob' type='checkbox' id='chkJob' onclick='CheckJob(this.form)' value=''></legend><table width='600' border='0' cellpadding='0' cellspacing='5'>"
        Response.Write "              <tr>"
        Response.Write "                <td width='20%'><input name='J_PE_JobCategory' type='checkbox' id='J_PE_JobCategory' onclick='UnCheckJob()' value='yes'> 工作类别</td>"
        Response.Write "                <td width='20%'><input name='J_PE_Position' type='checkbox' id='J_PE_Position' onclick='UnCheckJob()' value='yes'> 职位信息</td>"
        Response.Write "                <td width='20%'><input name='J_PE_PositionSupplyInfo' type='checkbox' id='J_PE_PositionSupplyInfo' onclick='UnCheckJob()' value='yes'> 申请职位记录</td>"
        Response.Write "                <td width='20%'><input name='J_PE_SubCompany' type='checkbox' id='J_PE_SubCompany' onclick='UnCheckJob()' value='yes'> 分公司信息</td>"
        Response.Write "                <td width='20%'><input name='J_PE_WorkPlace' type='checkbox' id='J_PE_WorkPlace' onclick='UnCheckJob()' value='yes'> 工作地点</td>"
        Response.Write "              </tr>"
        Response.Write "              <tr>"
        Response.Write "                <td width='20%'><input name='J_PE_Resume' type='checkbox' id='J_PE_Resume' onclick='UnCheckJob()' value='yes'> 个人简历</td>"
        Response.Write "                <td width='20%'></td>"
        Response.Write "                <td width='20%'></td>"
        Response.Write "                <td width='20%'></td>"
        Response.Write "                <td width='20%'></td>"
        Response.Write "              </tr>"
        Response.Write "            </table></fieldset>"
        Response.Write "          </td>"
        Response.Write "        </tr>"
        Response.Write "        <tr>"
        Response.Write "          <td>"
        Response.Write "            <fieldset name='ShoprData'><legend>房产数据 <input name='chkHouse' type='checkbox' id='chkHouse' onclick='CheckHouse(this.form)' value=''></legend><table width='600' border='0' cellpadding='0' cellspacing='5'>"
        Response.Write "              <tr>"
        Response.Write "                <td width='20%'><input name='H_PE_HouseConfig' type='checkbox' id='H_PE_HouseConfig' onclick='UnCheckHouse()' value='yes'> 房产信息配置</td>"
        Response.Write "                <td width='20%'><input name='H_PE_HouseCZ' type='checkbox' id='H_PE_HouseCZ' onclick='UnCheckHouse()' value='yes'> 出租信息</td>"
        Response.Write "                <td width='20%'><input name='H_PE_HouseCS' type='checkbox' id='H_PE_HouseCS' onclick='UnCheckHouse()' value='yes'> 出售信息</td>"
        Response.Write "                <td width='20%'><input name='H_PE_HouseQG' type='checkbox' id='H_PE_HouseQG' onclick='UnCheckHouse()' value='yes'> 求购信息</td>"
        Response.Write "                <td width='20%'><input name='H_PE_HouseQZ' type='checkbox' id='H_PE_HouseQZ' onclick='UnCheckHouse()' value='yes'> 求租信息</td>"
        Response.Write "              </tr>"
        Response.Write "              <tr>"
        Response.Write "                <td width='20%'><input name='H_PE_HouseHZ' type='checkbox' id='H_PE_HouseHZ' onclick='UnCheckHouse()' value='yes'> 合租信息</td>"
        Response.Write "                <td width='20%'></td>"
        Response.Write "                <td width='20%'></td>"
        Response.Write "                <td width='20%'></td>"
        Response.Write "                <td width='20%'></td>"
        Response.Write "              </tr>"
        Response.Write "            </table></fieldset>"
        Response.Write "          </td>"
        Response.Write "        </tr>"
        Response.Write "        <tr>"
        Response.Write "          <td>"
        Response.Write "            <fieldset name='OtherData'><legend>其他数据 <input name='chkOther' type='checkbox' id='chkOther' onclick='CheckOther(this.form)' value=''></legend><table width='600' border='0' cellpadding='0' cellspacing='5'>"
        Response.Write "              <tr>"
        Response.Write "                <td width='20%'><input name='PE_Announce' type='checkbox' id='PE_Announce' onclick='UnCheckOther()' value='yes'> 网站公告</td>"
        Response.Write "                <td width='20%'><input name='PE_Advertisement' type='checkbox' id='PE_Advertisement' onclick='UnCheckOther()' value='yes'> 网站广告</td>"
        Response.Write "                <td width='20%'><input name='PE_Vote' type='checkbox' id='PE_Vote' onclick='UnCheckOther()' value='yes'> 网站调查</td>"
        Response.Write "                <td width='20%'><input name='PE_FriendSite' type='checkbox' id='PE_FriendSite' onclick='UnCheckOther()' value='yes'> 友情链接</td>"
        Response.Write "                <td width='20%'><input name='PE_Log' type='checkbox' id='PE_Log' onclick='UnCheckOther()' value='yes'> 网站日志</td>"
        Response.Write "              </tr>"
        Response.Write "              <tr>"
        Response.Write "                <td width='20%'><input name='PE_GuestBook' type='checkbox' id='PE_GuestBook' onclick='UnCheckOther()' value='yes'> 所有留言</td>"
        Response.Write "                <td width='20%'><input name='PE_Author' type='checkbox' id='PE_Author' onclick='UnCheckOther()' value='yes'> 作者数据</td>"
        Response.Write "                <td width='20%'><input name='PE_CopyFrom' type='checkbox' id='PE_CopyFrom' onclick='UnCheckOther()' value='yes'> 来源数据</td>"
        Response.Write "                <td width='20%'><input name='PE_NewKeys' type='checkbox' id='PE_NewKeys' onclick='UnCheckOther()' value='yes'> 关 键 字</td>"
        Response.Write "                <td width='20%'><input name='PE_KeyLink' type='checkbox' id='PE_KeyLink' onclick='UnCheckOther()' value='yes'> 站内链接</td>"
        Response.Write "              </tr>"
        Response.Write "              <tr>"
        Response.Write "                <td width='20%'><input name='PE_User' type='checkbox' id='PE_User' onclick='UnCheckOther()' value='yes'> 注册会员</td>"
        Response.Write "                <td width='20%'><input name='PE_UserGroup' type='checkbox' id='PE_UserGroup' onclick='UnCheckOther()' value='yes'> 自定义会员组</td>"
        Response.Write "                <td width='20%'><input name='PE_ConsumeLog' type='checkbox' id='PE_ConsumeLog' onclick='UnCheckOther()' value='yes'> 消费明细</td>"
        Response.Write "                <td width='20%'><input name='PE_Favorite' type='checkbox' id='PE_Favorite' onclick='UnCheckOther()' value='yes'> 收藏记录</td>"
        Response.Write "                <td width='20%'><input name='PE_Card' type='checkbox' id='PE_Card' onclick='UnCheckOther()' value='yes'> 充 值 卡</td>"
        Response.Write "              </tr>"
        Response.Write "              <tr>"
        Response.Write "                <td width='20%'><input name='PE_Field' type='checkbox' id='PE_Field' onclick='UnCheckOther()' value='yes'> 自定义字段</td>"
        Response.Write "                <td width='20%'><input name='PE_Label' type='checkbox' id='PE_Label' onclick='UnCheckOther()' value='yes'> 自定义标签</td>"
        Response.Write "                <td width='20%'><input name='PE_Item' type='checkbox' id='PE_Item' onclick='UnCheckOther()' value='yes'> 采集数据</td>"
        Response.Write "                <td width='20%'><input name='PE_Equipment' type='checkbox' id='PE_Equipment' onclick='UnCheckOther()' value='yes'> 室场设备</td>"
        Response.Write "                <td width='20%'><input name='PE_Message' type='checkbox' id='PE_Message' onclick='UnCheckOther()' value='yes'> 所有短消息</td>"
        Response.Write "              </tr>"
        Response.Write "              <tr>"
        Response.Write "                <td width='20%'><input name='PE_Supply' type='checkbox' id='PE_Supply' onclick='UnCheckShop()' value='yes'> 供求信息</td>"
        Response.Write "                <td width='20%'></td>"
        Response.Write "                <td width='20%'></td>"
        Response.Write "                <td width='20%'></td>"
        Response.Write "                <td width='20%'></td>"
        Response.Write "              </tr>"
        Response.Write "            </table></fieldset>"
        Response.Write "          </td>"
        Response.Write "        </tr>"
        Response.Write "        <tr>"
        Response.Write "          <td align='center'><input name='Action' type='hidden' id='Action' value='Clear'>"
        Response.Write "            <input type='submit' name='Submit' value='清除所选数据库内容'>"
        Response.Write "          </td>"
        Response.Write "        </tr>"
        Response.Write "      </table>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    If Action <> "Clear" Then
        Response.Write "<b>说明：</b>&nbsp;&nbsp;<font color='#FF0000'>请慎用此功能，因为一旦清除将无法恢复！</font><br>"
    End If
End Sub

Sub SpaceSize()
    'On Error Resume Next
    Response.Write "<br><table class='border' width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td align='center' height='22' valign='middle'><b>系统空间占用情况</b></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='100%' height='150' valign='middle'>"
    Response.Write "    <blockquote>"
    Response.Write "      <br><b>系统文件占用空间情况：</b><br>"
    Response.Write "      基本系统占用空间：" & ShowSpace("SiteRoot|AD|Admin|Editor|FriendSite|Language|Inc|JS|Reg|Sdms|SiteMap|User|xml")
    Response.Write "      <br>"
    Response.Write "      系统图片占用空间：" & ShowSpace("Images|Skin|AuthorPic|BlogPic|CopyFromPic")
    Response.Write "      <br>"
    Response.Write "      统计系统占用空间：" & ShowSpace("Count")
    Response.Write "      <br>"
    Response.Write "      数 据 库占用空间：" & ShowSpace("Database")
    Response.Write "      <br>"
    Response.Write "      导出数据占用空间：" & ShowSpace("Temp")
    Response.Write "      <br><br>"
    Response.Write "      <b>频道文件占用空间情况：</b><br>"
    Dim rsChannel, sqlChannel
    sqlChannel = "select * from PE_Channel where ChannelType<=1 order by OrderID"
    Set rsChannel = Conn.Execute(sqlChannel)
    Do While Not rsChannel.EOF
        Response.Write "      <font color='#0000ff'>" & rsChannel("ChannelName") & "</font>占用空间：" & ShowSpace(rsChannel("ChannelDir"))
        Response.Write "      <br>"
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    Response.Write "      <br>未知目录占用空间：" & ShowSpace(GetOtherFolder())
    Response.Write "      <br>网站占用空间总计：" & ShowSpace(" ")
    Response.Write "    </blockquote>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub

Sub BackupData()
    Dim bkfolder, bkdbname
    bkfolder = Trim(Request("bkfolder"))
    bkdbname = Trim(Request("bkdbname"))
    If bkfolder = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定备份目录！</li>"
    End If
    If bkdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定备份文件名</li>"
    End If
    If FoundErr = True Then Exit Sub
    bkfolder = Server.MapPath(bkfolder)
    If fso.FileExists(dbpath) Then
        If fso.FolderExists(bkfolder) = False Then
            fso.CreateFolder (bkfolder)
        End If
        fso.copyfile dbpath, bkfolder & "\" & bkdbname & ".asa"
        Call WriteSuccessMsg("备份数据库成功，备份的数据库为：<br>" & bkfolder & "\" & bkdbname & ".asa", ComeUrl)
        Call WriteEntry(1, AdminName, "备份数据库")
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到源数据库文件，请检查Conn.asp中的配置。</li>"
    End If
End Sub

Sub CompactData()
    'On Error Resume Next

    Dim Engine, strDBPath
    Call CloseConn

    strDBPath = Left(dbpath, InStrRev(dbpath, "\"))
    If fso.FileExists(dbpath) Then
        Set Engine = Server.CreateObject("JRO.JetEngine")
        Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
        fso.copyfile strDBPath & "temp.mdb", dbpath
        fso.DeleteFile (strDBPath & "temp.mdb")
        Set Engine = Nothing
        Call WriteSuccessMsg("数据库压缩成功!", ComeUrl)
        Call OpenConn
        Call WriteEntry(1, AdminName, "压缩数据库")
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>数据库没有找到!</li>"
    End If
    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & Err.Description
        Err.Clear
        Exit Sub
    End If
End Sub

Sub RestoreData()
    Dim backpath
    backpath = Trim(Request.Form("backpath"))
    If backpath = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定原备份的数据库文件名！</li>"
        Exit Sub
    End If
    If GetFileExt(backpath) <> "asa" And GetFileExt(backpath) <> "asp" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>原备份数据库文件的扩展名必须为asa或asp！</li>"
        Exit Sub
    End If
    backpath = Server.MapPath(backpath)
    If fso.FileExists(backpath) Then
        fso.copyfile backpath, dbpath
        Call WriteSuccessMsg("成功恢复数据！", ComeUrl)
        Call WriteEntry(1, AdminName, "恢复数据库")
        Call ClearSiteCache(0)
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的备份文件！</li>"
    End If
End Sub

Sub ClearData()
    Dim strClear, z
    z = 0

    strClear = strClear & "<b>清空内容列表：</b>"
    strClear = strClear & "<br><br><b>频道数据：</b>"
    Dim ChannelTable, rsChannel, sqlChannel
    sqlChannel = "select * from PE_Channel where ChannelType<=1 and ChannelID<>4 order by OrderID"
    Set rsChannel = Conn.Execute(sqlChannel)
    Do While Not rsChannel.EOF
        If Request("C_PE_Class_" & rsChannel("ChannelID")) = "yes" Then
            Conn.Execute ("delete from PE_Class where ChannelID=" & rsChannel("ChannelID"))
            strClear = strClear & rsChannel("ChannelName") & "栏目&nbsp;&nbsp;"
            z = z + 1
        End If
        If Request("C_PE_Special_" & rsChannel("ChannelID")) = "yes" Then
            Conn.Execute ("delete from PE_Special where ChannelID=" & rsChannel("ChannelID"))
            strClear = strClear & rsChannel("ChannelName") & "专题&nbsp;&nbsp;"
            z = z + 1
        End If
        Select Case rsChannel("ModuleType")
        Case 1
            ChannelTable = "PE_Article"
        Case 2
            ChannelTable = "PE_Soft"
        Case 3
            ChannelTable = "PE_Photo"
        Case 5
            ChannelTable = "PE_Product"
        End Select
        If Request("C_" & ChannelTable & "_" & rsChannel("ChannelID")) = "yes" Then
            If ChannelTable = "PE_Product" Then
                Conn.Execute ("delete from PE_Product")
            Else
                Conn.Execute ("delete from " & ChannelTable & " where ChannelID=" & rsChannel("ChannelID"))
            End If
            strClear = strClear & rsChannel("ChannelName") & "数据&nbsp;&nbsp;"
            z = z + 1
        End If
        If Request("C_PE_Comment_" & rsChannel("ChannelID")) = "yes" Then
            
            '删除所属频道评论
            Dim Infotable
            If rsChannel("ModuleType") = 1 Then
                Infotable = "Article"
            ElseIf rsChannel("ModuleType") = 2 Then
                Infotable = "Photo"
            ElseIf rsChannel("ModuleType") = 3 Then
                Infotable = "Soft"
            ElseIf rsChannel("ModuleType") = 5 Then
                Infotable = "Product"
            End If
            
            Dim rsComment
            Set rsComment = Conn.Execute("Select I." & Infotable & "ID,I.ChannelID,C.InfoID from PE_" & Infotable & " I inner join PE_Comment C on I." & Infotable & "ID=C.InfoID where  I.ChannelID=" & rsChannel("ChannelID") & "")
            If rsComment.BOF And rsComment.EOF Then
            Else
                Do While Not rsComment.EOF
                    Conn.Execute "delete from PE_Comment where InfoID=" & rsComment("InfoID")
                    rsComment.MoveNext
                    z = z + 1
                Loop
            End If
            Set rsComment = Nothing
            'Conn.Execute ("delete from PE_Comment")
            strClear = strClear & rsChannel("ChannelName") & "评论&nbsp;&nbsp;"
        End If
        If Request("C_PE_JsFile_" & rsChannel("ChannelID")) = "yes" Then
            Conn.Execute ("delete from PE_JsFile where ChannelID=" & rsChannel("ChannelID"))
            strClear = strClear & rsChannel("ChannelName") & "JS数据&nbsp;&nbsp;"
            z = z + 1
        End If
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing

    strClear = strClear & "<br><br><b>商城数据：</b>"
    If Request("S_PE_OrderForm") = "yes" Then
        Conn.Execute ("delete from PE_OrderForm")
        Conn.Execute ("delete from PE_OrderFormItem")
        Conn.Execute ("delete from PE_TransferItem")
        strClear = strClear & "订单数据&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("S_PE_Bank") = "yes" Then
        Conn.Execute ("delete from PE_Bank")
        strClear = strClear & "银行帐户&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("S_PE_BankrollItem") = "yes" Then
        Conn.Execute ("delete from PE_BankrollItem")
        Conn.Execute ("delete from PE_InvoiceItem")
        strClear = strClear & "资金记录&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("S_PE_DeliverItem") = "yes" Then
        Conn.Execute ("delete from PE_DeliverItem")
        strClear = strClear & "发退货记录&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("S_PE_DeliverType") = "yes" Then
        Conn.Execute ("delete from PE_DeliverType")
        Conn.Execute ("delete from PE_DeliverCharge")
        strClear = strClear & "送货方式&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("S_PE_Payment") = "yes" Then
        Conn.Execute ("delete from PE_Payment")
        strClear = strClear & "在线支付记录&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("S_PE_PaymentType") = "yes" Then
        Conn.Execute ("delete from PE_PaymentType")
        strClear = strClear & "付款方式&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("S_PE_PresentProject") = "yes" Then
        Conn.Execute ("delete from PE_PresentProject")
        strClear = strClear & "促销方案&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("S_PE_Producer") = "yes" Then
        Conn.Execute ("delete from PE_Producer")
        strClear = strClear & "生 产 商&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("S_PE_Trademark") = "yes" Then
        Conn.Execute ("delete from PE_Trademark")
        strClear = strClear & "商品品牌&nbsp;&nbsp;"
        z = z + 1
    End If

    strClear = strClear & "<br><br><b>其他数据：</b>"
    If Request("PE_Announce") = "yes" Then
        Conn.Execute ("delete from PE_Announce")
        strClear = strClear & "网站公告&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_Advertisement") = "yes" Then
        Conn.Execute ("delete from PE_AdZone")
        Conn.Execute ("delete from PE_Advertisement")
        strClear = strClear & "网站广告&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_Vote") = "yes" Then
        Conn.Execute ("delete from PE_Vote")
        strClear = strClear & "网站调查&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_FriendSite") = "yes" Then
        Conn.Execute ("delete from PE_FsKind")
        Conn.Execute ("delete from PE_FriendSite")
        strClear = strClear & "友情链接&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_Log") = "yes" Then
        Conn.Execute ("delete from PE_Log")
        strClear = strClear & "网站日志&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_GuestBook") = "yes" Then
        Conn.Execute ("delete from PE_GuestBook")
        Conn.Execute ("delete from PE_GuestKind")
        strClear = strClear & "所有留言&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_Author") = "yes" Then
        Conn.Execute ("delete from PE_Author")
        strClear = strClear & "作者数据&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_CopyFrom") = "yes" Then
        Conn.Execute ("delete from PE_CopyFrom")
        strClear = strClear & "来源数据&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_NewKeys") = "yes" Then
        Conn.Execute ("delete from PE_NewKeys")
        strClear = strClear & "关键字&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_KeyLink") = "yes" Then
        Conn.Execute ("delete from PE_KeyLink")
        strClear = strClear & "站内链接&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_User") = "yes" Then
        Conn.Execute ("delete from PE_User")
        strClear = strClear & "注册会员&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_UserGroup") = "yes" Then
        Conn.Execute ("delete from PE_UserGroup where GroupType>2")
        strClear = strClear & "自定义会员组&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_ConsumeLog") = "yes" Then
        Conn.Execute ("delete from PE_ConsumeLog")
        Conn.Execute ("delete from PE_RechargeLog")
        strClear = strClear & "消费明细&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_Favorite") = "yes" Then
        Conn.Execute ("delete from PE_Favorite")
        strClear = strClear & "收藏记录&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_Card") = "yes" Then
        Conn.Execute ("delete from PE_Card")
        strClear = strClear & "充值卡&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_Field") = "yes" Then
        Conn.Execute ("delete from PE_Field")
        strClear = strClear & "自定义字段&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_Label") = "yes" Then
        Conn.Execute ("delete from PE_Label")
        strClear = strClear & "自定义标签&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_Item") = "yes" Then
        Conn.Execute ("delete from PE_Item")
        Conn.Execute ("delete from PE_Filters")
        Conn.Execute ("delete from PE_HistrolyNews")
        Conn.Execute ("delete from PE_AreaCollection")
        strClear = strClear & "采集数据&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_Equipment") = "yes" Then
        Conn.Execute ("delete from PE_Classroom")
        Conn.Execute ("delete from PE_Equipment")
        Conn.Execute ("delete from PE_UsedDetail")
        strClear = strClear & "室场设备&nbsp;&nbsp;"
        z = z + 1
    End If
    If Request("PE_Message") = "yes" Then
        Conn.Execute ("delete from PE_Message")
        Conn.Execute ("update PE_User set UnreadMsg=0")
        strClear = strClear & "所有短消息&nbsp;&nbsp;"
        z = z + 1
    End If

    If Request("PE_Supply") = "yes" Then
        Conn.Execute ("delete from PE_Supply")
        strClear = strClear & "供求信息&nbsp;&nbsp;"
        z = z + 1
    End If
    strClear = strClear & "<br><br><b>房产数据：</b>"
    '----------------------------------
    If Request("H_PE_HouseConfig") = "yes" Then
        Conn.Execute ("delete from PE_HouseConfig")
        strClear = strClear & "房产配置&nbsp;&nbsp;"
        z = z + 1
    End If

    If Request("H_PE_HouseCZ") = "yes" Then
        Conn.Execute ("delete from PE_HouseCZ")
        strClear = strClear & "所有出租信息&nbsp;&nbsp;"
        z = z + 1
    End If

    If Request("H_PE_HouseCS") = "yes" Then
        Conn.Execute ("delete from PE_HouseCS")
        strClear = strClear & "所有出售信息&nbsp;&nbsp;"
        z = z + 1
    End If

    If Request("H_PE_HouseQG") = "yes" Then
        Conn.Execute ("delete from PE_HouseQG")
        strClear = strClear & "所有求购信息&nbsp;&nbsp;"
        z = z + 1
    End If

    If Request("H_PE_HouseQZ") = "yes" Then
        Conn.Execute ("delete from PE_HouseQZ")
        strClear = strClear & "所有求租信息&nbsp;&nbsp;"
        z = z + 1
    End If

    If Request("H_PE_HouseHZ") = "yes" Then
        Conn.Execute ("delete from PE_HouseHZ")
        strClear = strClear & "所有合租信息&nbsp;&nbsp;"
        z = z + 1
    End If
    strClear = strClear & "<br><br><b>招聘数据：</b>"
    '----------------------------------
    If Request("J_PE_JobCategory") = "yes" Then
        Conn.Execute ("delete from PE_JobCategory")
        strClear = strClear & "工作类别&nbsp;&nbsp;"
        z = z + 1
    End If

    If Request("J_PE_Position") = "yes" Then
        Conn.Execute ("delete from PE_Position")
        strClear = strClear & "所有职位信息&nbsp;&nbsp;"
        z = z + 1
    End If

    If Request("J_PE_PositionSupplyInfo") = "yes" Then
        Conn.Execute ("delete from PE_PositionSupplyInfo")
        strClear = strClear & "所有申请职位数据&nbsp;&nbsp;"
        z = z + 1
    End If

    If Request("J_PE_Resume") = "yes" Then
        Conn.Execute ("delete from PE_Resume")
        strClear = strClear & "所有个人简历&nbsp;&nbsp;"
        z = z + 1
    End If

    If Request("J_PE_SubCompany") = "yes" Then
        Conn.Execute ("delete from PE_SubCompany")
        strClear = strClear & "所有分公司信息&nbsp;&nbsp;"
        z = z + 1
    End If

    If Request("J_PE_WorkPlace") = "yes" Then
        Conn.Execute ("delete from PE_WorkPlace")
        strClear = strClear & "所有工作地点&nbsp;&nbsp;"
        z = z + 1
    End If


    If z > 0 Then
        Response.Write "系统初始化成功！共计 <font color='red'>" & CStr(z) & "</font> 项内容被清空。<br><br>"
        Response.Write strClear
        Call WriteEntry(1, AdminName, "系统初始化")
        Call ClearSiteCache(0)
    Else
        Response.Write "您没有选择任何数据库内容，请重新选择后再进行操作。"
    End If
End Sub

Function ShowSpace(FolderPath)
    Dim ft, fd, fs, TotalSize, SpaceSize, FolderBarWidth, arrPath, strSize, i
    Set ft = fso.GetFolder(Server.MapPath(InstallDir))
    TotalSize = ft.size
    If TotalSize = 0 Then TotalSize = 1

    SpaceSize = 0
    arrPath = Split(FolderPath, "|")
    For i = 0 To UBound(arrPath)
        If arrPath(i) = "SiteRoot" Then
            Set fd = fso.GetFolder(Server.MapPath(InstallDir))
            For Each fs In fd.Files
                SpaceSize = SpaceSize + fs.size
            Next
        Else
            If fso.FolderExists(Server.MapPath(InstallDir & arrPath(i))) Then
                Set fd = fso.GetFolder(Server.MapPath(InstallDir & arrPath(i)))
                SpaceSize = SpaceSize + fd.size
            End If
        End If
    Next
    FolderBarWidth = PE_CLng((SpaceSize / TotalSize) * Barwidth)

    strSize = SpaceSize & "&nbsp;Byte"
    If SpaceSize > 1024 Then
       SpaceSize = (SpaceSize / 1024)
       strSize = FormatNumber(SpaceSize, 2, vbTrue, vbFalse, vbTrue) & "&nbsp;KB"
    End If
    If SpaceSize > 1024 Then
       SpaceSize = (SpaceSize / 1024)
       strSize = FormatNumber(SpaceSize, 2, vbTrue, vbFalse, vbTrue) & "&nbsp;MB"
    End If
    If SpaceSize > 1024 Then
       SpaceSize = (SpaceSize / 1024)
       strSize = FormatNumber(SpaceSize, 2, vbTrue, vbFalse, vbTrue) & "&nbsp;GB"
    End If
    strSize = "<font face=verdana>" & strSize & "</font>"
    ShowSpace = "&nbsp;<img src='../images/bar.gif' width='" & FolderBarWidth & "' height='10' title='" & FolderPath & "'>&nbsp;" & strSize
End Function

Function GetOtherFolder()
    Dim ft, fd, strOther, strSystem, arrPath
    strSystem = "AD|Admin|AuthorPic|BlogPic|CopyFromPic|Count|Database|Editor|FriendSite|Images|Inc|JS|Language|Reg|Sdms|SiteMap|Skin|Temp|User|xml"
    Dim rsChannel, sqlChannel
    sqlChannel = "select * from PE_Channel where ChannelType<=1 order by OrderID"
    Set rsChannel = Conn.Execute(sqlChannel)
    Do While Not rsChannel.EOF
        strSystem = strSystem & "|" & rsChannel("ChannelDir")
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing

    Set ft = fso.GetFolder(Server.MapPath(InstallDir))
    For Each fd In ft.SubFolders
        If InStr("|" & strSystem & "|", "|" & fd.name & "|") = 0 Then
            If strOther = "" Then
                strOther = fd.name
            Else
                strOther = strOther & "|" & fd.name
            End If
        End If
    Next
    GetOtherFolder = strOther
End Function
%>
