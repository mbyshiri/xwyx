<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 0      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

strFileName = "Admin_ShowPurview.asp"

Response.Write "<html><head><title>查看管理权限</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Response.Write "  <tr class='topbg'>" & vbCrLf
Response.Write "    <td height='22' colspan='2' align='center'><strong>查 看 管 理 权 限</strong></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>管理导航：</strong></td>" & vbCrLf
Response.Write "    <td> <a href='Admin_ShowPurview.asp'>管理权限首页</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
Response.Write "  <tr class='title'>"
Response.Write "    <td height='22'>" & GetChannelList() & "</td>"
Response.Write "  </tr>"
Response.Write "</table><br>"
Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
Response.Write "  <tr>"
Response.Write "    <td height='22'>" & GetManagePath() & "</td>"
Response.Write "  </tr>"
Response.Write "</table>"


If ChannelID = 0 Then
    Call ShowAllPurview
ElseIf ChannelID = 4 Then
    Call ShowGuestBookPurview
Else
    Call ShowChannelPurview
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub ShowAllPurview()

    Dim rsChannel, sqlChannel, rsAdmin, Channel_Purview
    sqlChannel = "select * from PE_Channel where ChannelType<=1 and ChannelID<>4 and Disabled=" & PE_False & " order by OrderID"
    Set rsChannel = Server.CreateObject("adodb.recordset")
    rsChannel.Open sqlChannel, Conn, 1, 1
    Do While Not rsChannel.EOF
        Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Set rsAdmin = Conn.Execute("select AdminPurview_" & rsChannel("ChannelDir") & " from PE_Admin where AdminName='" & AdminName & "'")
        If Not (rsAdmin.BOF And rsAdmin.EOF) Then
            Channel_Purview = rsAdmin(0)
        End If
        rsAdmin.Close
        Set rsAdmin = Nothing
        Response.Write "  <tr class='title' height='22'>"
        Response.Write "    <td colspan='4'><strong>" & rsChannel("ChannelName") & "</strong> "
        If Channel_Purview = 1 Then Response.Write "（频道管理员）"
        If Channel_Purview = 2 Then Response.Write "（栏目总编）"
        If Channel_Purview = 3 Then Response.Write "（栏目管理员）"
        If Channel_Purview = 4 Then Response.Write "（无权限）"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='30%'>各栏目" & rsChannel("ChannelShortName") & "录入、审核、管理权限</td>"
        Response.Write "    <td align='center' width='20%'>"
        If Channel_Purview <= 2 Then
            Response.Write "<font color=blue>全部权限</font>"
        ElseIf Channel_Purview = 3 Then
            Response.Write "<a href='Admin_ShowPurview.asp?iChannelID=" & rsChannel("ChannelID") & "'><font color=blue>部分权限</font></a>"
        Else
            Response.Write "<font color=red>无权限</font>"
        End If
        Response.Write "    </td>"
        Response.Write "    <td width='30%'>专题" & rsChannel("ChannelShortName") & "管理权限</td>"
        Response.Write "    <td align='center' width='20%'>"
        If AdminPurview = 1 Or Channel_Purview <= 2 Then
            Response.Write "<font color=blue>√</font>"
        Else
            Response.Write "<font color=red>×</font>"
        End If
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='30%'>栏目管理、专题管理、生成管理权限</td>"
        Response.Write "    <td align='center' width='20%'>"
        If AdminPurview = 1 Or Channel_Purview = 1 Then
            Response.Write "<font color=blue>√</font>"
        Else
            Response.Write "<font color=red>×</font>"
        End If
        Response.Write "    </td>"
        Response.Write "    <td width='30%'>" & rsChannel("ChannelShortName") & "评论、回收站及其它管理权限</td>"
        Response.Write "    <td align='center' width='20%'>"
        If AdminPurview = 1 Or Channel_Purview = 1 Then
            Response.Write "<font color=blue>√</font>"
        Else
            Response.Write "<font color=red>×</font>"
        End If
        Response.Write "    </td>"
        Response.Write "  </tr>"

        Response.Write "  <tr>"
        Response.Write "    <td class='tdbg' colspan='4'>"
        Response.Write "<b>更多权限：</b><br>"
        Response.Write "模板管理&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Template_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        Response.Write "JS文件管理&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "JsFile_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        Response.Write "顶部菜单&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Keyword_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        Response.Write "关键字管理&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Template_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        If rsChannel("ModuleType") = 5 Then
            Response.Write "厂商管理&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Producer_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
            Response.Write "品牌管理&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Trademark_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        Else
            Response.Write "作者管理&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Author_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
            Response.Write "来源管理&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Copyfrom_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        End If
        Response.Write "更新XML&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "XML_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        Response.Write "自定义字段&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "Field_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        Response.Write "广告管理&nbsp;" & ShowChannelOtherPurview(Channel_Purview, "AD_" & rsChannel("ChannelDir")) & "&nbsp;&nbsp;"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "</table>"
        Response.Write "<br>"
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing

    Dim rsGuestBook, sqlGuestBook, rsAdminGuest, GuestBook_Purview
    sqlGuestBook = "select * from PE_Channel where ChannelType<=1 and ChannelID=4 and Disabled=" & PE_False & ""
    Set rsGuestBook = Server.CreateObject("adodb.recordset")
    rsGuestBook.Open sqlGuestBook, Conn, 1, 1
    If Not (rsGuestBook.EOF And rsGuestBook.BOF) Then
        Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        Set rsAdminGuest = Conn.Execute("select AdminPurview_GuestBook from PE_Admin where AdminName='" & AdminName & "'")
        If Not (rsAdminGuest.BOF And rsAdminGuest.EOF) Then
            GuestBook_Purview = rsAdminGuest(0)
        End If
        rsAdminGuest.Close
        Set rsAdminGuest = Nothing
        Response.Write "  <tr class='title' height='22'>"
        Response.Write "    <td colspan='4'><strong>" & rsGuestBook("ChannelName") & "</strong> "
        If GuestBook_Purview = 1 Then Response.Write "（频道管理员）"
        If GuestBook_Purview = 2 Then Response.Write "（栏目总编）"
        If GuestBook_Purview = 3 Then Response.Write "（栏目管理员）"
        If GuestBook_Purview = 4 Then Response.Write "（无权限）"
        Response.Write "    </td>"
        Response.Write "  </tr>"

        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='30%'>各栏目" & rsGuestBook("ChannelShortName") & "修改、删除、移动、审核、精华、固顶、回复权限</td>"
        Response.Write "    <td align='center' width='20%'>"
        If GuestBook_Purview <= 2 Then
            Response.Write "<font color=blue>全部权限</font>"
        ElseIf GuestBook_Purview = 3 Then
            Response.Write "<a href='Admin_ShowPurview.asp?iChannelID=" & rsGuestBook("ChannelID") & "'><font color=blue>部分权限</font></a>"
        Else
            Response.Write "<font color=red>无权限</font>"
        End If
        Response.Write "    </td>"
        Response.Write "    <td width='30%'>可以管理" & rsGuestBook("ChannelShortName") & "类别</td>"
        Response.Write "    <td align='center' width='20%'>"
        If AdminPurview = 1 Or GuestBook_Purview <= 2 Then
            Response.Write "<font color=blue>√</font>"
        Else
            Response.Write "<font color=red>×</font>"
        End If
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='30%'>栏目管理，可以执行首页嵌入代码生成</td>"
        Response.Write "    <td align='center' width='20%'>"
        If AdminPurview = 1 Or GuestBook_Purview = 1 Then
            Response.Write "<font color=blue>√</font>"
        Else
            Response.Write "<font color=red>×</font>"
        End If
        Response.Write "    </td>"
        Response.Write "    <td width='30%'>" & rsGuestBook("ChannelShortName") & "其它管理权限</td>"
        Response.Write "    <td align='center' width='20%'>"
        If AdminPurview = 1 Or GuestBook_Purview = 1 Then
            Response.Write "<font color=blue>√</font>"
        Else
            Response.Write "<font color=red>×</font>"
        End If
        Response.Write "    </td>"
        Response.Write "  </tr>"

        Response.Write "  <tr>"
        Response.Write "    <td class='tdbg' colspan='4'>"
        Response.Write "<b>更多权限：</b><br>"
        Response.Write "广告管理&nbsp;" & ShowChannelOtherPurview(GuestBook_Purview, "AD_" & rsGuestBook("ChannelDir")) & "&nbsp;&nbsp;"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "</table>"
        Response.Write "<br>"
    End If
    rsGuestBook.Close
    Set rsGuestBook = Nothing
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td colspan='6' height='22'><strong>其他网站管理权限</strong><strong> </strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>修改自己密码 " & ShowPurview("ModifyPwd") & "</td>"
    Response.Write "    <td width='16%'>网站频道管理 " & ShowPurview("Channel") & "</td>"
    Response.Write "    <td width='16%'>采集管理 " & ShowPurview("Collection") & "</td>"
    Response.Write "    <td width='16%'>短消息管理 " & ShowPurview("Message") & "</td>"
    Response.Write "    <td width='16%'>邮件列表管理 " & ShowPurview("MailList") & "</td>"
    Response.Write "    <td width='16%'>网站广告管理 " & ShowPurview("AD") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>友情链接管理 " & ShowPurview("FriendSite") & "</td>"
    Response.Write "    <td width='16%'>网站公告管理 " & ShowPurview("Announce") & "</td>"
    Response.Write "    <td width='16%'>网站调查管理 " & ShowPurview("Vote") & "</td>"
    Response.Write "    <td width='16%'>网站统计管理 " & ShowPurview("Counter") & "</td>"
    Response.Write "    <td width='16%'>网站风格管理 " & ShowPurview("Skin") & "</td>"
    Response.Write "    <td width='16%'>通用模板管理 " & ShowPurview("Template") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>自定义标签管理 " & ShowPurview("Label") & "</td>"
    Response.Write "    <td width='16%'>网站缓存管理 " & ShowPurview("Cache") & "</td>"
    Response.Write "    <td width='16%'>站内链接管理 " & ShowPurview("KeyLink") & "</td>"
    Response.Write "    <td width='16%'>字符过滤管理 " & ShowPurview("Rtext") & "</td>"
    Response.Write "    <td width='16%'>会员组管理 " & ShowPurview("UserGroup") & "</td>"
    Response.Write "    <td width='16%'>充值卡管理 " & ShowPurview("Card") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>室场登记管理 " & ShowPurview("Equipment") & "</td>"
    Response.Write "    <td width='16%'>学生信息管理 " & ShowPurview("InfoManage") & "</td>"
    Response.Write "    <td width='16%'>学生成绩管理 " & ShowPurview("ScoreManage") & "</td>"
    Response.Write "    <td width='16%'>考试管理 " & ShowPurview("TestManage") & "</td>"
    Response.Write "    <td width='16%'></td>"
    Response.Write "    <td width='16%'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td colspan='6' height='22'><strong>会员管理权限</strong><strong> </strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>查看会员信息 " & ShowPurview("User_View") & "</td>"
    Response.Write "    <td width='16%'>修改会员信息 " & ShowPurview("User_ModifyInfo") & "</td>"
    Response.Write "    <td width='16%'>修改会员权限 " & ShowPurview("User_MofidyPurview") & "</td>"
    Response.Write "    <td width='16%'>锁住/解锁会员 " & ShowPurview("User_Lock") & "</td>"
    Response.Write "    <td width='16%'>删除会员 " & ShowPurview("User_Del") & "</td>"
    Response.Write "    <td width='16%'>升级为客户 " & ShowPurview("User_Update") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>会员资金管理 " & ShowPurview("User_Money") & "</td>"
    Response.Write "    <td width='16%'>会员点券管理 " & ShowPurview("User_Point") & "</td>"
    Response.Write "    <td width='16%'>会员有效期管理 " & ShowPurview("User_Valid") & "</td>"
    Response.Write "    <td width='16%'>会员消费明细 " & ShowPurview("ConsumeLog") & "</td>"
    Response.Write "    <td width='16%'>会员有效期明细 " & ShowPurview("RechargeLog") & "</td>"
    Response.Write "    <td width='16%'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td colspan='6' height='22'><strong>商城日常操作管理权限</strong><strong> </strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>查看订单 " & ShowPurview("Order_View") & "</td>"
    Response.Write "    <td width='16%'>确认订单 " & ShowPurview("Order_Confirm") & "</td>"
    Response.Write "    <td width='16%'>修改订单 " & ShowPurview("Order_Modify") & "</td>"
    Response.Write "    <td width='16%'>删除订单 " & ShowPurview("Order_Del") & "</td>"
    Response.Write "    <td width='16%'>收款处理 " & ShowPurview("Order_Payment") & "</td>"
    Response.Write "    <td width='16%'>开发票 " & ShowPurview("Order_Invoice") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>订单配送（实物） " & ShowPurview("Order_Deliver") & "</td>"
    Response.Write "    <td width='16%'>订单配送（软件） " & ShowPurview("Order_Download") & "</td>"
    Response.Write "    <td width='16%'>订单配送（点卡） " & ShowPurview("Order_SendCard") & "</td>"
    Response.Write "    <td width='16%'>结清订单 " & ShowPurview("Order_End") & "</td>"
    Response.Write "    <td width='16%'>订单过户 " & ShowPurview("Order_Transfer") & "</td>"
    Response.Write "    <td width='16%'>订单打印 " & ShowPurview("Order_Print") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>订单统计 " & ShowPurview("Order_Count") & "</td>"
    Response.Write "    <td width='16%'>销售明细情况 " & ShowPurview("Order_OrderItem") & "</td>"
    Response.Write "    <td width='16%'>销售统计/排行 " & ShowPurview("Order_SaleCount") & "</td>"
    Response.Write "    <td width='16%'>在线支付管理 " & ShowPurview("Payment") & "</td>"
    Response.Write "    <td width='16%'>资金明细查询 " & ShowPurview("Bankroll") & "</td>"
    Response.Write "    <td width='16%'>发退货记录 " & ShowPurview("Deliver") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='16%'>订单过户记录 " & ShowPurview("Transfer") & "</td>"
    Response.Write "    <td width='16%'>促销方案管理 " & ShowPurview("PresentProject") & "</td>"
    Response.Write "    <td width='16%'>付款方式管理 " & ShowPurview("PaymentType") & "</td>"
    Response.Write "    <td width='16%'>送货方式管理 " & ShowPurview("DeliverType") & "</td>"
    Response.Write "    <td width='16%'>银行帐户管理 " & ShowPurview("Bank") & "</td>"
    Response.Write "    <td width='16%'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td colspan='5' height='22'><strong>客户关系管理权限</strong><strong> </strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='17%'>查看客户信息 " & ShowPurview("Client_View") & "</td>"
    Response.Write "    <td width='17%'>添加客户 " & ShowPurview("Client_Add") & "</td>"
    Response.Write "    <td width='25%'>修改属于自己的客户信息 " & ShowPurview("Client_ModifyOwn") & "</td>"
    Response.Write "    <td width='25%'>修改所有客户信息 " & ShowPurview("Client_ModifyAll") & "</td>"
    Response.Write "    <td width='17%'>删除客户 " & ShowPurview("Client_Del") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='17%'>查看服务记录 " & ShowPurview("Service_View") & "</td>"
    Response.Write "    <td width='17%'>添加服务记录 " & ShowPurview("Service_Add") & "</td>"
    Response.Write "    <td width='25%'>修改自己添加的服务记录 " & ShowPurview("Service_ModifyOwn") & "</td>"
    Response.Write "    <td width='25%'>修改所有服务记录 " & ShowPurview("Service_ModifyAll") & "</td>"
    Response.Write "    <td width='17%'>删除服务记录 " & ShowPurview("Service_Del") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='17%'>查看投诉记录 " & ShowPurview("Complain_View") & "</td>"
    Response.Write "    <td width='17%'>添加投诉记录 " & ShowPurview("Complain_Add") & "</td>"
    Response.Write "    <td width='25%'>修改自己添加的投诉记录 " & ShowPurview("Complain_ModifyOwn") & "</td>"
    Response.Write "    <td width='25%'>修改所有投诉记录 " & ShowPurview("Complain_ModifyAll") & "</td>"
    Response.Write "    <td width='17%'>删除投诉记录 " & ShowPurview("Complain_Del") & "</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='17%'>查看回访记录 " & ShowPurview("Call_View") & "</td>"
    Response.Write "    <td width='17%'>添加回访记录 " & ShowPurview("Call_Add") & "</td>"
    Response.Write "    <td width='25%'>修改自己添加的回访记录 " & ShowPurview("Call_ModifyOwn") & "</td>"
    Response.Write "    <td width='25%'>修改所有回访记录 " & ShowPurview("Call_ModifyAll") & "</td>"
    Response.Write "    <td width='17%'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

End Sub

Sub ShowGuestBookPurview()
    Dim rsAdminGuest, GuestBook_Purview, arrKind_GuestBook
    Dim rsGuestBook, GuestBookDir, GuestBookName, GuestBookShortName

    Set rsGuestBook = Conn.Execute("select * from PE_Channel where ChannelID=4")
    If Not (rsGuestBook.BOF Or rsGuestBook.EOF) Then
        GuestBookDir = rsGuestBook("ChannelDir")
        GuestBookName = rsGuestBook("ChannelName")
        GuestBookShortName = rsGuestBook("ChannelShortName")
        Set rsAdminGuest = Conn.Execute("select AdminPurview_GuestBook,arrClass_GuestBook from PE_Admin where AdminName='" & AdminName & "'")
        If Not (rsAdminGuest.BOF And rsAdminGuest.EOF) Then
            GuestBook_Purview = rsAdminGuest(0)
            arrKind_GuestBook = Split(rsAdminGuest(1), "|||")
        End If
        rsAdminGuest.Close
        Set rsAdminGuest = Nothing
    End If
    rsGuestBook.Close
    Set rsGuestBook = Nothing

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td colspan='2' height='22'><strong>" & GuestBookName & "</strong> "
    If GuestBook_Purview = 1 Then Response.Write "（频道管理员）"
    If GuestBook_Purview = 2 Then Response.Write "（栏目总编）"
    If GuestBook_Purview = 3 Then Response.Write "（栏目管理员）"
    If GuestBook_Purview = 4 Then Response.Write "（无权限）"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>栏目管理权限，可以执行首页嵌入代码生成</td>"
    Response.Write "    <td align='center' width='20%'>"
    If AdminPurview = 1 Or GuestBook_Purview = 1 Then
        Response.Write "<font color=blue>√</font>"
    Else
        Response.Write "<font color=red>×</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>可以管理" & GuestBookShortName & "类别</td>"
    Response.Write "    <td align='center' width='20%'>"
    If AdminPurview = 1 Or GuestBook_Purview <= 2 Then
        Response.Write "<font color=blue>√</font>"
    Else
        Response.Write "<font color=red>×</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>" & GuestBookShortName & "其它管理权限</td>"
    Response.Write "    <td align='center' width='20%'>"
    If AdminPurview = 1 Or GuestBook_Purview = 1 Then
        Response.Write "<font color=blue>√</font>"
    Else
        Response.Write "<font color=red>×</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>各栏目" & GuestBookShortName & "修改、删除、移动、审核、精华、固顶、回复权限</td>"
    Response.Write "    <td align='center' width='20%'>"
    If GuestBook_Purview <= 2 Then
        Response.Write "<font color=blue>全部权限</font>"
    ElseIf GuestBook_Purview = 3 Then
        Response.Write "<font color=blue>部分权限</font>"
    Else
        Response.Write "<font color=red>无权限</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22'><strong>更多权限：</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td>"
    Response.Write "广告管理&nbsp;" & ShowPurview("AD_" & GuestBookDir) & "&nbsp;&nbsp;"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

    If GuestBook_Purview = 3 Then
        Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
        Response.Write "  <tr align='center' class='title'>"
        Response.Write "    <td height='22'><strong>栏目名称</strong></td>"
        Response.Write "    <td width='70'><strong>修改</strong></td>"
        Response.Write "    <td width='70'><strong>删除</strong></td>"
        Response.Write "    <td width='70'><strong>移动</strong></td>"
        Response.Write "    <td width='70'><strong>审核</strong></td>"
        Response.Write "    <td width='70'><strong>精华</strong></td>"
        Response.Write "    <td width='70'><strong>固顶</strong></td>"
        Response.Write "    <td width='70'><strong>回复</strong></td>"
        Response.Write "  </tr>"
        Dim rsGuestKind
        Set rsGuestKind = Conn.Execute("select * from PE_GuestKind order by OrderID,KindID")
        Do While Not rsGuestKind.EOF
            Response.Write "  <tr class='tdbg'>"
            Response.Write "    <td align='center'>" & rsGuestKind("KindName") & "</td>"
            Response.Write "    <td align='center'>"
            If FoundInArr(arrKind_GuestBook(0), rsGuestKind("KindID"), ",") = True Then
                Response.Write "<font color=blue>√</font>"
            Else
                Response.Write "<font color=red>×</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrKind_GuestBook(1), rsGuestKind("KindID"), ",") = True Then
                Response.Write "<font color=blue>√</font>"
            Else
                Response.Write "<font color=red>×</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrKind_GuestBook(2), rsGuestKind("KindID"), ",") = True Then
                Response.Write "<font color=blue>√</font>"
            Else
                Response.Write "<font color=red>×</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrKind_GuestBook(3), rsGuestKind("KindID"), ",") = True Then
                Response.Write "<font color=blue>√</font>"
            Else
                Response.Write "<font color=red>×</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrKind_GuestBook(4), rsGuestKind("KindID"), ",") = True Then
                Response.Write "<font color=blue>√</font>"
            Else
                Response.Write "<font color=red>×</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrKind_GuestBook(5), rsGuestKind("KindID"), ",") = True Then
                Response.Write "<font color=blue>√</font>"
            Else
                Response.Write "<font color=red>×</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrKind_GuestBook(6), rsGuestKind("KindID"), ",") = True Then
                Response.Write "<font color=blue>√</font>"
            Else
                Response.Write "<font color=red>×</font>"
            End If
            Response.Write "</td>"
            Response.Write "  </tr>"
            rsGuestKind.MoveNext
        Loop
        Set rsGuestKind = Nothing
        Response.Write "</table>"
    End If
End Sub

Sub ShowChannelPurview()
    Dim rsAdmin, Channel_Purview, arrClass_View, arrClass_Input, arrClass_Check, arrClass_Manage
    Dim rsChannel, ChannelDir, ChannelName, ChannelShortName, ModuleType

    If ChannelID > 0 Then
        Set rsChannel = Conn.Execute("select * from PE_Channel where ChannelID=" & ChannelID)
        If Not (rsChannel.BOF Or rsChannel.EOF) Then
            ChannelDir = rsChannel("ChannelDir")
            ChannelName = rsChannel("ChannelName")
            ChannelShortName = rsChannel("ChannelShortName")
            ModuleType = rsChannel("ModuleType")
            Set rsAdmin = Conn.Execute("select AdminPurview_" & ChannelDir & ",arrClass_View,arrClass_Input,arrClass_Check,arrClass_Manage from PE_Admin where AdminName='" & AdminName & "'")
            If Not (rsAdmin.BOF And rsAdmin.EOF) Then
                Channel_Purview = rsAdmin(0)
                arrClass_View = rsAdmin("arrClass_View")
                arrClass_Input = rsAdmin("arrClass_Input")
                arrClass_Check = rsAdmin("arrClass_Check")
                arrClass_Manage = rsAdmin("arrClass_Manage")
            End If
            rsAdmin.Close
            Set rsAdmin = Nothing
        End If
        rsChannel.Close
        Set rsChannel = Nothing
    End If

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td colspan='2' height='22'><strong>" & ChannelName & "</strong> "
    If Channel_Purview = 1 Then Response.Write "（频道管理员）"
    If Channel_Purview = 2 Then Response.Write "（栏目总编）"
    If Channel_Purview = 3 Then Response.Write "（栏目管理员）"
    If Channel_Purview = 4 Then Response.Write "（无权限）"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>栏目管理、专题管理、生成管理权限</td>"
    Response.Write "    <td align='center' width='20%'>"
    If AdminPurview = 1 Or Channel_Purview = 1 Then
        Response.Write "<font color=blue>√</font>"
    Else
        Response.Write "<font color=red>×</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>专题" & ChannelShortName & "管理权限</td>"
    Response.Write "    <td align='center' width='20%'>"
    If AdminPurview = 1 Or Channel_Purview <= 2 Then
        Response.Write "<font color=blue>√</font>"
    Else
        Response.Write "<font color=red>×</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>" & ChannelShortName & "评论、回收站及其它管理权限</td>"
    Response.Write "    <td align='center' width='20%'>"
    If AdminPurview = 1 Or Channel_Purview = 1 Then
        Response.Write "<font color=blue>√</font>"
    Else
        Response.Write "<font color=red>×</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td width='30%'>各栏目" & ChannelShortName & "录入、审核、管理权限</td>"
    Response.Write "    <td align='center' width='20%'>"
    If Channel_Purview <= 2 Then
        Response.Write "<font color=blue>全部权限</font>"
    ElseIf Channel_Purview = 3 Then
        Response.Write "<font color=blue>部分权限</font>"
    Else
        Response.Write "<font color=red>无权限</font>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22'><strong>更多权限：</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td>"
    Response.Write "模板管理&nbsp;" & ShowPurview("Template_" & ChannelDir) & "&nbsp;&nbsp;"
    Response.Write "JS文件管理&nbsp;" & ShowPurview("JsFile_" & ChannelDir) & "&nbsp;&nbsp;"
    Response.Write "顶部菜单&nbsp;" & ShowPurview("Keyword_" & ChannelDir) & "&nbsp;&nbsp;"
    Response.Write "关键字管理&nbsp;" & ShowPurview("Template_" & ChannelDir) & "&nbsp;&nbsp;"
    If ModuleType = 5 Then
        Response.Write "厂商管理&nbsp;" & ShowPurview("Producer_" & ChannelDir) & "&nbsp;&nbsp;"
        Response.Write "品牌管理&nbsp;" & ShowPurview("Trademark_" & ChannelDir) & "&nbsp;&nbsp;"
    Else
        Response.Write "作者管理&nbsp;" & ShowPurview("Author_" & ChannelDir) & "&nbsp;&nbsp;"
        Response.Write "来源管理&nbsp;" & ShowPurview("Copyfrom_" & ChannelDir) & "&nbsp;&nbsp;"
    End If
    Response.Write "更新XML&nbsp;" & ShowPurview("XML_" & ChannelDir) & "&nbsp;&nbsp;"
    Response.Write "自定义字段&nbsp;" & ShowPurview("Field_" & ChannelDir) & "&nbsp;&nbsp;"
    Response.Write "广告管理&nbsp;" & ShowPurview("AD_" & ChannelDir) & "&nbsp;&nbsp;"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<br>"

    If Channel_Purview = 3 Then
        Dim arrShowLine(20)
        Dim sqlClass, rsClass, i, iDepth
        For i = 0 To UBound(arrShowLine)
            arrShowLine(i) = False
        Next
        sqlClass = "select * from PE_Class where ChannelID=" & ChannelID & " order by RootID,OrderID"
        Set rsClass = Server.CreateObject("adodb.recordset")
        rsClass.Open sqlClass, Conn, 1, 1

        Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
        Response.Write "  <tr align='center' class='title'>"
        Response.Write "    <td height='22'><strong>栏目名称</strong></td>"
        Response.Write "    <td width='100'><strong>查看</strong></td>"
        Response.Write "    <td width='100'><strong>录入</strong></td>"
        Response.Write "    <td width='100'><strong>审核</strong></td>"
        Response.Write "    <td width='100'><strong>管理</strong></td>"
        Response.Write "  </tr>"
        Do While Not rsClass.EOF
            Response.Write "     <tr class='tdbg'><td>"
            iDepth = rsClass("Depth")
            If rsClass("NextID") > 0 Then
                arrShowLine(iDepth) = True
            Else
                arrShowLine(iDepth) = False
            End If
            If iDepth > 0 Then
                For i = 1 To iDepth
                    If i = iDepth Then
                        If rsClass("NextID") > 0 Then
                            Response.Write "<img src='../images/tree_line1.gif' width='17' height='16' valign='abvmiddle'>"
                        Else
                            Response.Write "<img src='../images/tree_line2.gif' width='17' height='16' valign='abvmiddle'>"
                        End If
                    Else
                        If arrShowLine(i) = True Then
                            Response.Write "<img src='../images/tree_line3.gif' width='17' height='16' valign='abvmiddle'>"
                        Else
                            Response.Write "<img src='../images/tree_line4.gif' width='17' height='16' valign='abvmiddle'>"
                        End If
                    End If
                Next
            End If
            If rsClass("Child") > 0 Then
                Response.Write "<img src='../images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>"
            Else
                Response.Write "<img src='../images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>"
            End If
            If rsClass("Depth") = 0 Then
                Response.Write "<b>"
            End If
            Response.Write rsClass("ClassName")
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrClass_View, rsClass("ClassID"), ",") = True Then
                Response.Write "<font color=blue>√</font>"
            Else
                Response.Write "<font color=red>×</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrClass_Input, rsClass("ClassID"), ",") = True Then
                Response.Write "<font color=blue>√</font>"
            Else
                Response.Write "<font color=red>×</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrClass_Check, rsClass("ClassID"), ",") = True Then
                Response.Write "<font color=blue>√</font>"
            Else
                Response.Write "<font color=red>×</font>"
            End If
            Response.Write "</td><td align='center'>"
            If FoundInArr(arrClass_Manage, rsClass("ClassID"), ",") = True Then
                Response.Write "<font color=blue>√</font>"
            Else
                Response.Write "<font color=red>×</font>"
            End If
            Response.Write "</td></tr>"
        rsClass.MoveNext
        Loop
        rsClass.Close
        Set rsClass = Nothing
        Response.Write "</table>"
    End If
End Sub

Function ShowPurview(strPurview)
    If CheckPurview_Other(AdminPurview_Others, strPurview) = True Then
        ShowPurview = "<font color=blue>√</font>"
    Else
        ShowPurview = "<font color=red>×</font>"
    End If
End Function

Function ShowChannelOtherPurview(Channel_Purview, strPurview)
    If ChannelPurview = 1 And CheckPurview_Other(AdminPurview_Others, strPurview) = True Then
        ShowChannelOtherPurview = "<font color=blue>√</font>"
    Else
        ShowChannelOtherPurview = "<font color=red>×</font>"
    End If
End Function

Function GetChannelList()
    Dim rsChannel, sqlChannel, strChannel, i
    If ChannelID = 0 Then
        strChannel = "<a href='" & strFileName & "?iChannelID=0'><font color=red>所有管理权限</font></a> | "
    Else
        strChannel = "<a href='" & strFileName & "?iChannelID=0'>所有管理权限</a> | "
    End If
    sqlChannel = "select * from PE_Channel where ChannelType<=1 and Disabled=" & PE_False & " order by OrderID"
    Set rsChannel = Server.CreateObject("adodb.recordset")
    rsChannel.Open sqlChannel, Conn, 1, 1
    If rsChannel.BOF And rsChannel.EOF Then
        strChannel = strChannel & "没有任何频道"
    Else
        i = 1
        Do While Not rsChannel.EOF
            If rsChannel("ChannelID") = ChannelID Then
                strChannel = strChannel & "<a href='" & strFileName & "?iChannelID=" & ChannelID & "'><font color=red>" & rsChannel("ChannelName") & "权限</font></a>"
            Else
                strChannel = strChannel & "<a href='" & strFileName & "?iChannelID=" & rsChannel("ChannelID") & "'>" & rsChannel("ChannelName") & "权限</a>"
            End If
            strChannel = strChannel & " | "
            i = i + 1
            If i Mod 10 = 0 Then
                strChannel = strChannel & "<br>"
            End If
            rsChannel.MoveNext
        Loop
    End If
    rsChannel.Close
    Set rsChannel = Nothing
    GetChannelList = strChannel
End Function

Function GetManagePath()
    Dim strPath, sqlPath, rsPath
    strPath = "您现在的位置：查看管理权限&nbsp;&gt;&gt;&nbsp;"
    If ChannelID = 0 Then
        strPath = strPath & "所有管理权限"
    Else
        sqlPath = "select ChannelID,ChannelName from PE_Channel where ChannelID=" & ChannelID
        Set rsPath = Server.CreateObject("adodb.recordset")
        rsPath.Open sqlPath, Conn, 1, 1
        If rsPath.BOF And rsPath.EOF Then
            strPath = strPath & "错误的频道参数"
        Else
            strPath = strPath & "<a href='" & strFileName & "?iChannelID=" & rsPath(0) & "'>" & rsPath(1) & "权限</a>"
        End If
        rsPath.Close
        Set rsPath = Nothing
    End If
    GetManagePath = strPath
End Function
%>
