<!--#include file="Admin_Common.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = "Bank"   '其他权限
    
Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<title>银行帐户管理</title>" & vbCrLf
Response.Write "<link href='Admin_STYLE.CSS' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "<body>" & vbCrLf
Response.Write "<table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("银 行 帐 户 管 理", 10212)
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td width='70' height='30' class='tdbg'>管理导航：</td>" & vbCrLf
Response.Write "    <td class='tdbg'><a href='Admin_Bank.asp'>银行帐户管理首页</a> | <a href='Admin_Bank.asp?Action=Add'>添加银行帐户</a></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
Select Case Action
Case "Add"
    Call Add
Case "Modify"
    Call Modify
Case "SaveAdd", "SaveModify"
    Call SaveBank
Case "Disable", "Enable"
    Call DisableBank
Case "SetDefault"
    Call SetDefault
Case "Del"
    Call Del
Case "Order"
    Call Order
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn


Sub main()
    Response.Write "<br>" & vbCrLf
    Response.Write "<table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr align='center' class='title'>" & vbCrLf
    Response.Write "    <td width='30'>ID</td>" & vbCrLf
    Response.Write "    <td width='60'>帐户名称</td>" & vbCrLf
    Response.Write "    <td width='80'>开户行</td>" & vbCrLf
    Response.Write "    <td width='100'>户名</td>" & vbCrLf
    Response.Write "    <td>帐号/卡号</td>" & vbCrLf
    Response.Write "    <td width='50'>是否默认</td>" & vbCrLf
    Response.Write "    <td width='40'>已启用</td>" & vbCrLf
    Response.Write "    <td width='150'>常规操作</td>" & vbCrLf
    Response.Write "    <td width='80'>排序操作</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Dim rsBank
    Set rsBank = Conn.Execute("select * from PE_Bank order by OrderID asc")
    If rsBank.BOF And rsBank.EOF Then
        Response.Write "<tr><td colspan='10' height='50' align='center'>没有任何银行帐户</td></tr>"
    Else
        Do While Not rsBank.EOF
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbg2'"">" & vbCrLf
            Response.Write "    <td width='30' align='center'>" & rsBank("BankID") & "</td>" & vbCrLf
            Response.Write "    <td width='60' align='center'>" & rsBank("BankShortName") & "</td>" & vbCrLf
            Response.Write "    <td width='80' align='left'>" & rsBank("BankName") & "</td>" & vbCrLf
            Response.Write "    <td width='100' align='center'>" & rsBank("HolderName") & "</td>" & vbCrLf
            Response.Write "    <td align='left'>帐号：" & rsBank("Accounts") & "<br>卡号：" & rsBank("CardNum") & "</td>" & vbCrLf
            Response.Write "    <td width='50' align='center'>" & vbCrLf
            If rsBank("IsDefault") = True Then
                Response.Write "√"
            End If
            Response.Write "      </td>" & vbCrLf
            Response.Write "    <td width='40' align='center'>" & vbCrLf
            If rsBank("IsDisabled") = False Then
                Response.Write "√"
            Else
                Response.Write "<font color='red'>×</font>"
            End If
            Response.Write "</td>" & vbCrLf
            Response.Write "    <td width='150' align='center'>" & vbCrLf
            If rsBank("IsDefault") = True Then
                Response.Write "<font color='gray'>设为默认 禁用</font> "
            Else
                Response.Write "<a href='Admin_Bank.asp?Action=SetDefault&BankID=" & rsBank("BankID") & "'>设为默认</a> "
                If rsBank("IsDisabled") = True Then
                    Response.Write "<a href='Admin_Bank.asp?Action=Enable&BankID=" & rsBank("BankID") & "'>启用</a> "
                Else
                    Response.Write "<a href='Admin_Bank.asp?Action=Disable&BankID=" & rsBank("BankID") & "'>禁用</a> "
                End If
            End If
            Response.Write "<a href='Admin_Bank.asp?Action=Modify&BankID=" & rsBank("BankID") & "'>修改</a> "
            If rsBank("IsDefault") = True Then
                Response.Write "<font color='gray'>删除</font> "
            Else
                Response.Write "<a href='Admin_Bank.asp?Action=Del&BankID=" & rsBank("BankID") & "'>删除</a>"
            End If
            Response.Write "</td><form name='orderform' method='post' action='Admin_Bank.asp'>"
            Response.Write "    <td width='80' align='center'><input name='OrderID' type='text' id='OrderID' value='" & rsBank("OrderID") & "' size='4' maxlength='4' style='text-align:center '><input type='submit' name='Submit' value='修改'><input name='BankID' type='hidden' id='BankID' value='" & rsBank("BankID") & "'><input name='Action' type='hidden' id='Action' value='Order'></td></form>"
            Response.Write "  </tr>"
            rsBank.MoveNext
        Loop
    End If
    Set rsBank = Nothing
    Response.Write "</table>"
    Response.Write "<br>" & vbCrLf
    Response.Write "说明：“禁用”某银行帐户后，输入汇款信息时将不再显示此银行帐户，但在资金明细情况中仍会显示。<br>" & vbCrLf
End Sub

Sub Add()
    Response.Write "<form name='myform' method='post' action='Admin_Bank.asp'>" & vbCrLf
    Response.Write "  <table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center'>" & vbCrLf
    Response.Write "      <td colspan='2' class='title'><b>添 加 银 行 帐 户</b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>帐户名称：</td>" & vbCrLf
    Response.Write "      <td><input name='BankShortName' type='text' id='BankShortName' size='20' maxlength='20'> <font color='#FF0000'>*</font> 请认真填写，一旦录入就不可修改。</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>开户行：</td>" & vbCrLf
    Response.Write "      <td><input name='BankName' type='text' id='BankName' size='20' maxlength='50'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>户名：</td>" & vbCrLf
    Response.Write "      <td><input name='HolderName' type='text' id='HolderName' size='20' maxlength='20'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>帐号：</td>" & vbCrLf
    Response.Write "      <td><input name='Accounts' type='text' id='Accounts' size='20' maxlength='30'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>卡号：</td>" & vbCrLf
    Response.Write "      <td><input name='CardNum' type='text' id='CardNum' size='20' maxlength='30'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>银行图标：</td>" & vbCrLf
    Response.Write "      <td><input name='BankPic' type='text' id='BankPic' size='40' maxlength='200'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>帐户说明：</td>" & vbCrLf
    Response.Write "      <td><textarea name='BankIntro' cols='40' rows='3'></textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td align='right'>&nbsp;</td>" & vbCrLf
    Response.Write "      <td><input name='IsDefault' type='checkbox' id='IsDefault' value='Yes'> 设为默认银行帐户</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <td height='50' colspan='2'><input name='Action' type='hidden' id='Action' value='SaveAdd'><input type='submit' name='Submit' value='保存银行帐户'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

Sub Modify()
    Dim BankID, rsBank
    BankID = Trim(Request("BankID"))
    If BankID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定银行帐户ID</li>"
        Exit Sub
    Else
        BankID = PE_CLng(BankID)
    End If
    Set rsBank = Conn.Execute("select * from PE_Bank where BankID=" & BankID & "")
    If rsBank.BOF And rsBank.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的银行帐户！</li>"
        Set rsBank = Nothing
        Exit Sub
    End If
    Response.Write "<form name='myform' method='post' action='Admin_Bank.asp'>" & vbCrLf
    Response.Write "  <table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center'>" & vbCrLf
    Response.Write "      <td colspan='2' class='title'><b>修 改 银 行 帐 户</b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>帐户名称：</td>" & vbCrLf
    Response.Write "      <td><input name='BankShortName' type='text' id='BankShortName' size='20' maxlength='20' value='" & rsBank("BankShortName") & "' disabled> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>开户行：</td>" & vbCrLf
    Response.Write "      <td><input name='BankName' type='text' id='BankName' size='20' maxlength='50' value='" & rsBank("BankName") & "'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>户名：</td>" & vbCrLf
    Response.Write "      <td><input name='HolderName' type='text' id='HolderName' size='20' maxlength='20' value='" & rsBank("HolderName") & "'> <font color='#FF0000'>*</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>帐号：</td>" & vbCrLf
    Response.Write "      <td><input name='Accounts' type='text' id='Accounts' size='20' maxlength='30' value='" & rsBank("Accounts") & "'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>卡号：</td>" & vbCrLf
    Response.Write "      <td><input name='CardNum' type='text' id='CardNum' size='20' maxlength='30' value='" & rsBank("CardNum") & "'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>银行图标：</td>" & vbCrLf
    Response.Write "      <td><input name='BankPic' type='text' id='BankPic' size='40' maxlength='200' value='" & rsBank("BankPic") & "'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='30%' align='right'>帐户说明：</td>" & vbCrLf
    Response.Write "      <td><textarea name='BankIntro' cols='40' rows='3'>" & rsBank("BankIntro") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td align='right'>&nbsp;</td>" & vbCrLf
    Response.Write "      <td><input name='IsDefault' type='checkbox' id='IsDefault' value='Yes'"
    If rsBank("IsDefault") = True Then Response.Write " checked"
    Response.Write "> 设为默认银行帐户</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <td height='50' colspan='2'><input name='BankID' type='hidden' id='BankID' value='" & BankID & "'>" & vbCrLf
    Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveModify'>" & vbCrLf
    Response.Write "          <input type='submit' name='Submit' value='保存银行帐户'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    Response.Write "" & vbCrLf
    Set rsBank = Nothing
End Sub

Sub SaveBank()
    Dim BankID, BankShortName, BankName, Accounts, CardNum, HolderName, BankIntro, BankPic, IsDefault, OrderID
    Dim rsBank, sqlBank
    BankID = Trim(Request("BankID"))
    BankShortName = ReplaceBadChar(Trim(Request("BankShortName")))
    BankName = Trim(Request("BankName"))
    Accounts = Trim(Request("Accounts"))
    CardNum = Trim(Request("CardNum"))
    HolderName = Trim(Request("HolderName"))
    BankIntro = Trim(Request("BankIntro"))
    BankPic = Trim(Request("BankPic"))
    IsDefault = Trim(Request("IsDefault"))
    If Action = "SaveAdd" And BankShortName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定帐户名称</li>"
    End If
    If BankName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定开户行</li>"
    End If
    If HolderName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定户名</li>"
    End If
    If Accounts = "" And CardNum = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>至少要输入帐号和卡号的其中一个</li>"
    End If
    If IsDefault = "Yes" Then
        IsDefault = True
        Conn.Execute ("update PE_Bank set IsDefault=" & PE_False & "")
    Else
        IsDefault = False
    End If
    
    If Action = "SaveAdd" Then
        sqlBank = "select top 1 * from PE_Bank"
        Dim trs, mrs
        Set trs = Conn.Execute("select BankID from PE_Bank where BankShortName='" & BankShortName & "'")
        If Not (trs.BOF And trs.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>指定的帐户名称已经存在！</li>"
        Else
            Set mrs = Conn.Execute("select max(BankID) from PE_Bank")
            If IsNull(mrs(0)) Then
                BankID = 1
            Else
                BankID = mrs(0) + 1
            End If
            Set mrs = Nothing
            Set mrs = Conn.Execute("select max(OrderID) from PE_Bank")
            If IsNull(mrs(0)) Then
                OrderID = 1
            Else
                OrderID = mrs(0) + 1
            End If
            Set mrs = Nothing
        End If
        Set trs = Nothing
    Else
        If BankID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定BankID！</li>"
        Else
            BankID = PE_CLng(BankID)
        End If
        sqlBank = "select * from PE_Bank where BankID=" & BankID
    End If
    
    If FoundErr = True Then Exit Sub
    
    Set rsBank = Server.CreateObject("adodb.recordset")
    rsBank.Open sqlBank, Conn, 1, 3
    If Action = "SaveAdd" Then
        rsBank.AddNew
        rsBank("BankID") = BankID
        rsBank("OrderID") = OrderID
        rsBank("IsDisabled") = False
        rsBank("BankShortName") = BankShortName
    End If
    
    rsBank("BankName") = BankName
    rsBank("Accounts") = Accounts
    rsBank("CardNum") = CardNum
    rsBank("HolderName") = HolderName
    rsBank("BankIntro") = BankIntro
    rsBank("BankPic") = BankPic
    rsBank("IsDefault") = IsDefault
    rsBank.Update
    rsBank.Close
    Set rsBank = Nothing
    Call WriteEntry(2, AdminName, "保存银行帐户信息成功：" & BankName)
    Call CloseConn
    Response.Redirect "Admin_Bank.asp"
End Sub

Sub Del()
    Dim BankID
    BankID = Trim(Request("BankID"))
    If BankID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定银行帐户ID</li>"
        Exit Sub
    Else
        BankID = PE_CLng(BankID)
    End If
    
    Dim rsBank, trs
    Set rsBank = Server.CreateObject("adodb.recordset")
    rsBank.Open "select * from PE_Bank where BankID=" & BankID & "", Conn, 1, 3
    If rsBank.BOF And rsBank.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的银行帐户"
    Else
        If rsBank("IsDefault") = True Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>不能删除默认的银行帐户</li>"
        End If
        Set trs = Conn.Execute("select top 1 ItemID from PE_BankrollItem where Bank='" & rsBank("BankShortName") & "'")
        If Not (trs.BOF And trs.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>此银行帐户已经有资金明细记录，所以不能删除！但你可以禁用此银行帐户，以达到前台不显示此银行帐户的目的。</li>"
        End If
        Set trs = Nothing
        If FoundErr = False Then
            rsBank.Delete
            rsBank.Update
        End If
    End If
    rsBank.Close
    Set rsBank = Nothing
    Call WriteEntry(2, AdminName, "删除银行帐户成功，ID：" & BankID)

    Call CloseConn
    If FoundErr = False Then
        Response.Redirect "Admin_Bank.asp"
    End If
End Sub

Sub DisableBank()
    Dim BankID
    BankID = Trim(Request("BankID"))
    If BankID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定银行帐户ID</li>"
    Else
        BankID = PE_CLng(BankID)
    End If
    If FoundErr = True Then Exit Sub
    Dim trs
    Set trs = Conn.Execute("select IsDefault from PE_Bank where BankID=" & BankID & "")
    If trs.BOF And trs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的银行帐户</li>"
    Else
        If trs(0) = True Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>不能禁用默认的银行帐户</li>"
        End If
    End If
    Set trs = Nothing
    If FoundErr = True Then Exit Sub
    
    Select Case Action
    Case "Disable"
        Conn.Execute ("update PE_Bank set IsDisabled=" & PE_True & " where BankID=" & BankID & "")
    Case "Enable"
        Conn.Execute ("update PE_Bank set IsDisabled=" & PE_False & " where BankID=" & BankID & "")
    End Select

    Call CloseConn
    Response.Redirect "Admin_Bank.asp"
End Sub

Sub Order()
    Dim BankID, OrderID
    BankID = Trim(Request("BankID"))
    OrderID = Trim(Request("OrderID"))
    If BankID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定银行帐户ID</li>"
    Else
        BankID = PE_CLng(BankID)
    End If
    If OrderID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定OrderID</li>"
    Else
        OrderID = PE_CLng(OrderID)
    End If
    If FoundErr = True Then Exit Sub
    Conn.Execute ("update PE_Bank set OrderID=" & OrderID & " where BankID=" & BankID & "")
    Call CloseConn
    Response.Redirect "Admin_Bank.asp"
End Sub

Sub SetDefault()
    Dim BankID
    BankID = Trim(Request("BankID"))
    If BankID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定BankID</li>"
        Exit Sub
    Else
        BankID = PE_CLng(BankID)
    End If
    Conn.Execute ("update PE_Bank set IsDefault=" & PE_False & "")
    Conn.Execute ("update PE_Bank set IsDefault=" & PE_True & " where  BankID=" & BankID)
    Call CloseConn
    Response.Redirect "Admin_Bank.asp"
End Sub
%>
