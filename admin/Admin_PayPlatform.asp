<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 1      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<title>在线支付平台管理</title>" & vbCrLf
Response.Write "<link href='Admin_STYLE.CSS' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "<body>" & vbCrLf
Response.Write "<table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("在 线 支 付 平 台 管 理", 10212)
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td width='70' height='30' class='tdbg'>管理导航：</td>" & vbCrLf
Response.Write "    <td class='tdbg'><a href='Admin_PayPlatform.asp'>在线支付平台管理</a> | <a href='Admin_PayPlatform.asp?ManageType=Order'>在线支付平台排序</a></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
Select Case Action
Case "Modify"
    Call Modify
Case "SaveModify"
    Call SaveBank
Case "Disable", "Enable"
    Call DisableBank
Case "SetDefault"
    Call SetDefault
Case "Order"
    Call Order
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteEntry(2, AdminName, "在线支付平台管理操作失败，失败原因：" & ErrMsg)
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub main()
    Dim ManageType
    ManageType = Trim(Request("ManageType"))
    Response.Write "<br>" & vbCrLf
    Response.Write "<table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr align='center' class='title'>" & vbCrLf
    Response.Write "    <td width='30'>ID</td>" & vbCrLf
    Response.Write "    <td width='80'>平台名称</td>" & vbCrLf
    Response.Write "    <td width='120'>商户ID</td>" & vbCrLf
    Response.Write "    <td>平台说明</td>" & vbCrLf
    Response.Write "    <td width='60'>手续费率</td>" & vbCrLf
    Response.Write "    <td width='50'>是否默认</td>" & vbCrLf
    Response.Write "    <td width='40'>已启用</td>" & vbCrLf
    If ManageType <> "Order" Then
    Response.Write "    <td width='100'>常规操作</td>" & vbCrLf
    Else
    Response.Write "    <td width='100'>排序操作</td>" & vbCrLf
    End If
    Response.Write "  </tr>" & vbCrLf
    Dim rsPayPlatform, PayPlatformUrl
    Set rsPayPlatform = Conn.Execute("select * from PE_PayPlatform order by OrderID asc")
    If rsPayPlatform.BOF And rsPayPlatform.EOF Then
        Response.Write "<tr><td colspan='10' height='50' align='center'>没有任何在线支付平台</td></tr>"
    Else
        Do While Not rsPayPlatform.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbg2'"">" & vbCrLf
            Response.Write "    <td width='30' align='center'>" & rsPayPlatform("PlatformID") & "</td>" & vbCrLf
            Response.Write "    <td width='80' align='center'>" & rsPayPlatform("PlatformName") & "</td>" & vbCrLf
            Response.Write "    <td width='120' align='left'>" & rsPayPlatform("AccountsID") & "</td>" & vbCrLf
            Response.Write "    <td align='left' style=word-wrap:break-word;'>" & rsPayPlatform("Description") & "</td>" & vbCrLf
            Response.Write "    <td width='60' align='center'>" & rsPayPlatform("Rate") & "%</td>" & vbCrLf
            Response.Write "    <td width='50' align='center'>" & vbCrLf
            If rsPayPlatform("IsDefault") = True Then
                Response.Write "√"
            End If
            Response.Write "      </td>" & vbCrLf
            Response.Write "    <td width='40' align='center'>" & vbCrLf
            If rsPayPlatform("IsDisabled") = False Then
                Response.Write "√"
            Else
                Response.Write "<font color='red'>×</font>"
            End If
            Response.Write "</td>" & vbCrLf
            If ManageType <> "Order" Then
                Response.Write "    <td width='100' align='center'>" & vbCrLf
                Select Case rsPayPlatform("PlatformID")
                Case 1
                    PayPlatformUrl = "http://merchant3.chinabank.com.cn/register.do"
                Case 2
                    PayPlatformUrl = "http://www.ipay.cn"
                Case 3
                    PayPlatformUrl = "https://www.ips.com.cn"
                Case 4
                    PayPlatformUrl = "#"
                Case 5
                    PayPlatformUrl = "http://www.yeepay.com/"
                Case 6
                    PayPlatformUrl = "http://new.xpay.cn/SignUp/Default.aspx"
                Case 7
                    PayPlatformUrl = "https://www.cncard.net"
                Case 8
                    PayPlatformUrl = "https://www.alipay.com/"
                Case 9
                    PayPlatformUrl = "http://www.99bill.com/"
                Case 10
                    PayPlatformUrl = "#"
                Case 11
                    PayPlatformUrl = "http://www.99bill.com/"
                Case 12
                    PayPlatformUrl = "https://www.alipay.com/"
                Case 13
                    PayPlatformUrl = "http://union.tenpay.com/mch/mch_register.shtml?posid=123&actid=84&opid=50&whoid=31&sp_suggestuser=1201648901"
                End Select
                Response.Write "<a href='Admin_PayPlatform.asp?Action=Modify&PlatformID=" & rsPayPlatform("PlatformID") & "'>修改</a> "
                Response.Write "<a href='" & PayPlatformUrl & "' target='_blank'>申请商户</a><br>"
                If rsPayPlatform("IsDisabled") = True Then
                    Response.Write "<a href='Admin_PayPlatform.asp?Action=Enable&PlatformID=" & rsPayPlatform("PlatformID") & "'>启用</a> "
                Else
                    If rsPayPlatform("IsDefault") = True Then
                        Response.Write "<font color='gray'>禁用</font> "
                    Else
                        Response.Write "<a href='Admin_PayPlatform.asp?Action=Disable&PlatformID=" & rsPayPlatform("PlatformID") & "'>禁用</a> "
                    End If
                End If
                If rsPayPlatform("IsDisabled") = True Or rsPayPlatform("IsDefault") = True Then
                    Response.Write "<font color='gray'>设为默认</font> <br>"
                Else
                    Response.Write "<a href='Admin_PayPlatform.asp?Action=SetDefault&PlatformID=" & rsPayPlatform("PlatformID") & "'>设为默认</a>"
                End If
                Response.Write "</td>"
            Else
                Response.Write "<form name='orderform' method='post' action='Admin_PayPlatform.asp'>"
                Response.Write "    <td width='100' align='center'><input name='OrderID' type='text' id='OrderID' value='" & rsPayPlatform("OrderID") & "' size='4' maxlength='4' style='text-align:center '><input type='submit' name='Submit' value='修改'><input name='PlatformID' type='hidden' id='PlatformID' value='" & rsPayPlatform("PlatformID") & "'><input name='Action' type='hidden' id='Action' value='Order'></td></form>"
            End If
            Response.Write "  </tr>"
            rsPayPlatform.MoveNext
        Loop
    End If
    Set rsPayPlatform = Nothing
    Response.Write "</table>"
    Response.Write "<br>" & vbCrLf
    Response.Write "说明：“禁用”某在线支付平台后，在线支付时将不再显示此在线支付平台，但在在线支付记录管理中仍会显示。<br>" & vbCrLf
End Sub

Sub Modify()
    Dim PlatformID, rsPayPlatform
    PlatformID = Trim(Request("PlatformID"))
    If PlatformID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定在线支付平台ID</li>"
        Exit Sub
    Else
        PlatformID = PE_CLng(PlatformID)
    End If
    Set rsPayPlatform = Conn.Execute("select * from PE_PayPlatform where PlatformID=" & PlatformID & "")
    If rsPayPlatform.BOF And rsPayPlatform.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的在线支付平台！</li>"
        Set rsPayPlatform = Nothing
        Exit Sub
    End If
    Response.Write "<form name='myform' method='post' action='Admin_PayPlatform.asp'>" & vbCrLf
    Response.Write "  <table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center'>" & vbCrLf
    Response.Write "      <td colspan='2' class='title'><b>修 改 在 线 支 付 平 台</b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>平台名称：</td>" & vbCrLf
    Response.Write "      <td><input name='PlatformName' type='text' id='PlatformName' size='50' maxlength='20' value='" & rsPayPlatform("PlatformName") & "' disabled></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>对外显示的名称：</td>" & vbCrLf
    Response.Write "      <td><input name='ShowName' type='text' id='ShowName' size='50' maxlength='30' value='" & rsPayPlatform("ShowName") & "'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>说明：</td>" & vbCrLf
    Response.Write "      <td><textarea name='Description' cols='42' rows='5'>" & rsPayPlatform("Description") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>商户ID</td>" & vbCrLf
    Response.Write "      <td><input name='AccountsID' type='text' id='AccountsID' size='50' maxlength='50' value='" & rsPayPlatform("AccountsID") & "'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>MD5密钥：</td>" & vbCrLf
    Response.Write "      <td><input name='MD5Key' type='password' id='MD5Key' size='50' maxlength='255' value='" & rsPayPlatform("MD5Key") & "'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>手续费率：</td>" & vbCrLf
    Response.Write "      <td><input name='Rate' type='text' id='Rate' size='5' maxlength='5' value='" & rsPayPlatform("Rate") & "'>% <font color='#FF0000'>*</font><br>" & vbCrLf
    Response.Write "        <input name='PlusPoundage' type='checkbox' value='1' " & IsRadioChecked(rsPayPlatform("PlusPoundage"), True) & "> 手续费由付款人额外支付</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <td height='50' colspan='2'><input name='PlatformID' type='hidden' id='PlatformID' value='" & PlatformID & "'>" & vbCrLf
    Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveModify'>" & vbCrLf
    Response.Write "          <input type='submit' name='Submit' value='保存在线支付平台'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    Response.Write "" & vbCrLf
    Set rsPayPlatform = Nothing
End Sub

Sub SaveBank()
    Dim PlatformID, ShowName, AccountsID, Accounts, MD5Key, Rate
    Dim rsPayPlatform, sqlPlatform
    PlatformID = Trim(Request("PlatformID"))
    ShowName = Trim(Request("ShowName"))
    AccountsID = Trim(Request("AccountsID"))
    Accounts = Trim(Request("Accounts"))
    MD5Key = Trim(Request("MD5Key"))
    Rate = Trim(Request("Rate"))
    If PlatformID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定PlatformID！</li>"
    Else
        PlatformID = PE_CLng(PlatformID)
    End If
    If ShowName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定平台对外显示的名称</li>"
    End If
    If AccountsID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定商户ID</li>"
    End If
    If MD5Key = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定MD5密钥</li>"
    End If
    If Rate = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定手续费率</li>"
    Else
        Rate = PE_CDbl(Rate)
    End If
    
    
    If FoundErr = True Then Exit Sub
    
    sqlPlatform = "select * from PE_PayPlatform where PlatformID=" & PlatformID
    Set rsPayPlatform = Server.CreateObject("adodb.recordset")
    rsPayPlatform.Open sqlPlatform, Conn, 1, 3
    
    rsPayPlatform("ShowName") = ShowName
    rsPayPlatform("Description") = Trim(Request("Description"))
    rsPayPlatform("AccountsID") = AccountsID
    rsPayPlatform("MD5Key") = MD5Key
    rsPayPlatform("Rate") = Rate
    rsPayPlatform("PlusPoundage") = PE_CBool(Trim(Request("PlusPoundage")))
    rsPayPlatform.Update
    rsPayPlatform.Close
    Set rsPayPlatform = Nothing
    Call WriteEntry(2, AdminName, "保存在线支付平台信息成功：" & AccountsID)
    Call CloseConn
    Response.Redirect "Admin_PayPlatform.asp"
End Sub

Sub DisableBank()
    Dim PlatformID
    PlatformID = Trim(Request("PlatformID"))
    If PlatformID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定在线支付平台ID</li>"
    Else
        PlatformID = PE_CLng(PlatformID)
    End If
    If FoundErr = True Then Exit Sub
    Dim trs
    Set trs = Conn.Execute("select IsDefault from PE_PayPlatform where PlatformID=" & PlatformID & "")
    If trs.BOF And trs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的在线支付平台</li>"
    Else
        If trs(0) = True Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>不能禁用默认的在线支付平台</li>"
        End If
    End If
    Set trs = Nothing
    If FoundErr = True Then Exit Sub
    
    Select Case Action
    Case "Disable"
        Conn.Execute ("update PE_PayPlatform set IsDisabled=" & PE_True & " where PlatformID=" & PlatformID & "")
    Case "Enable"
        Conn.Execute ("update PE_PayPlatform set IsDisabled=" & PE_False & " where PlatformID=" & PlatformID & "")
    End Select

    Call CloseConn
    Response.Redirect "Admin_PayPlatform.asp"
End Sub

Sub Order()
    Dim PlatformID, OrderID
    PlatformID = Trim(Request("PlatformID"))
    OrderID = Trim(Request("OrderID"))
    If PlatformID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定在线支付平台ID</li>"
    Else
        PlatformID = PE_CLng(PlatformID)
    End If
    If OrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定OrderID</li>"
    Else
        OrderID = PE_CLng(OrderID)
    End If
    If FoundErr = True Then Exit Sub
    Conn.Execute ("update PE_PayPlatform set OrderID=" & OrderID & " where PlatformID=" & PlatformID & "")
    Call CloseConn
    Response.Redirect "Admin_PayPlatform.asp"
End Sub

Sub SetDefault()
    Dim PlatformID
    PlatformID = Trim(Request("PlatformID"))
    If PlatformID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定PlatformID</li>"
        Exit Sub
    Else
        PlatformID = PE_CLng(PlatformID)
    End If

    Conn.Execute ("update PE_PayPlatform set IsDefault=" & PE_False & "")
    Conn.Execute ("update PE_PayPlatform set IsDefault=" & PE_True & " where  PlatformID=" & PlatformID)
    Call CloseConn
    Response.Redirect "Admin_PayPlatform.asp"
End Sub
%>
