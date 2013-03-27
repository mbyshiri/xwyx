<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Base64.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = "Card"   '其他权限

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>充值卡管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "<SCRIPT language=javascript>" & vbCrLf
Response.Write "function unselectall()" & vbCrLf
Response.Write "{" & vbCrLf
Response.Write "    if(document.myform.chkAll.checked){" & vbCrLf
Response.Write "    document.myform.chkAll.checked = document.myform.chkAll.checked&0;" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "function CheckAll(form)" & vbCrLf
Response.Write "{" & vbCrLf
Response.Write "  for (var i=0;i<form.elements.length;i++)" & vbCrLf
Response.Write "    {" & vbCrLf
Response.Write "    var e = form.elements[i];" & vbCrLf
Response.Write "    if (e.Name != 'chkAll'&&e.disabled!=true)" & vbCrLf
Response.Write "       e.checked = form.chkAll.checked;" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "}" & vbCrLf

Response.Write "function ShowGroup()" & vbCrLf
Response.Write "{" & vbCrLf
Response.Write "    var sel= document.myform.ValidUnit[document.all.ValidUnit.selectedIndex].value;"
Response.Write "    if(sel=='5')" & vbCrLf
Response.Write "    {" & vbCrLf
Response.Write "        document.myform.GroupList.style.display='';" & vbCrLf
Response.Write "        document.myform.ValidNum.disabled = true;" & vbCrLf
Response.Write "        document.myform.ValidNum.value = '1';" & vbCrLf
Response.Write "        HelpInfoForPoint.style.display='none';" & vbCrLf
Response.Write "        HelpInfoForGroup.style.display='';" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "    else" & vbCrLf
Response.Write "    {" & vbCrLf
Response.Write "        document.myform.GroupList.style.display='none';" & vbCrLf
Response.Write "        document.myform.ValidNum.disabled = false;" & vbCrLf
Response.Write "        document.myform.ValidNum.value = '500';" & vbCrLf
Response.Write "        HelpInfoForPoint.style.display='';" & vbCrLf
Response.Write "        HelpInfoForGroup.style.display='none';" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "}" & vbCrLf

Response.Write "function SetNumValue()" & vbCrLf
Response.Write "{" & vbCrLf
Response.Write "    document.myform.ValidNum.value = document.myform.GroupList[document.all.GroupList.selectedIndex].value;" & vbCrLf
Response.Write "}" & vbCrLf

Response.Write "</script>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("充 值 卡 管 理", 10043)
Response.Write "  <tr class='tdbg'> " & vbCrLf
Response.Write "    <td width='70' height='30'><strong>管理导航：</strong></td>" & vbCrLf
Response.Write "    <td height='30'><a href='Admin_Card.asp'>所有充值卡</a>&nbsp;| <a href='Admin_Card.asp?CardStatus=1'>所有未使用的充值卡</a> | <a href='Admin_Card.asp?CardStatus=2'>所有已使用的充值卡</a> | <a href='Admin_Card.asp?CardStatus=3'>所有已失效的充值卡</a> | <a href='Admin_Card.asp?Action=Add'>添加充值卡</a> |&nbsp;<a href='Admin_Card.asp?Action=BatchAdd'>批量生成充值卡</a></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Select Case Action
Case "Add"
    Call Add
Case "SaveAdd"
    Call SaveAdd
Case "BatchAdd"
    Call BatchAdd
Case "DoBatchAdd"
    Call DoBatchAdd
Case "Modify"
    Call Modify
Case "SaveModify"
    Call SaveModify
Case "Del"
    Call DelCard
Case "Show"
    Call Show
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim CardType, CardStatus, AgentName
    CardType = Trim(Request("CardType"))
    CardStatus = PE_CLng(Trim(Request("CardStatus")))
    AgentName = ReplaceBadChar(Trim(Request("AgentName")))
    strFileName = "Admin_Card.asp?CardType=" & CardType & "&CardStatus=" & CardStatus & "&Field=" & strField & "&Keyword=" & Keyword
    
    Response.Write "<br>" & vbCrLf
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "  <form name='myform' method='Post' action='Admin_Card.asp'>" & vbCrLf
    Response.Write "     <td>" & vbCrLf
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr align='center' class='title' height='22'>" & vbCrLf
    Response.Write "    <td width='30'><strong>选中</strong></td>" & vbCrLf
    Response.Write "    <td width=60><strong>类 型</strong></td>" & vbCrLf
    Response.Write "    <td width=100><strong>卡 号</strong></td>" & vbCrLf
    Response.Write "    <td width=40><strong>面值</strong></td>" & vbCrLf
    Response.Write "    <td width=40><strong>点 数</strong></td>" & vbCrLf
    Response.Write "    <td width='60'><strong>截止日期</strong></td>" & vbCrLf
    Response.Write "    <td><strong>所属商品</strong></td>" & vbCrLf
    Response.Write "    <td width='40'><strong>状 态</strong></td>" & vbCrLf
    Response.Write "    <td width='60'><strong>使用者</strong></td>" & vbCrLf
    Response.Write "    <td width='120'><strong>充值时间</strong></td>" & vbCrLf
    Response.Write "    <td width='60'><strong>代理商</strong></td>" & vbCrLf
    Response.Write "    <td width='60'><strong> 操 作</strong></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    
    Dim sqlCard, rsCard, i
    sqlCard = "select C.*,P.ProductName from PE_Card C left join PE_Product P on C.ProductID=P.ProductID where 1=1"
    Select Case CardType
    Case "0"
        sqlCard = sqlCard & " and C.CardType=0"
    Case "1"
        sqlCard = sqlCard & " and C.CardType=1"
    End Select
    Select Case CardStatus
    Case 1
        sqlCard = sqlCard & " and C.UserName='' and C.EndDate>=" & PE_Now & ""
    Case 2
        sqlCard = sqlCard & " and C.UserName<>''"
    Case 3
        sqlCard = sqlCard & " and C.UserName='' and C.EndDate<" & PE_Now & ""
    End Select
    If strField <> "" Then
        Select Case strField
        Case "CardNum"
            sqlCard = sqlCard & " and C.CardNum like '%" & Keyword & "%'"
        Case "Money"
            sqlCard = sqlCard & " and C.Money=" & PE_CDbl(Keyword) & ""
        Case "AgentName"
            sqlCard = sqlCard & " and C.AgentName='" & Keyword & "'"
        Case "UserName"
            sqlCard = sqlCard & " and C.UserName='" & Keyword & "'"
        End Select
    End If
    If AgentName <> "" Then
        sqlCard = sqlCard & " and C.AgentName='" & AgentName & "'"
    End If
    sqlCard = sqlCard & " order by C.CardID desc"
    Set rsCard = Server.CreateObject("Adodb.RecordSet")
    rsCard.Open sqlCard, Conn, 1, 1
    If rsCard.Bof And rsCard.EOF Then
        rsCard.Close
        Set rsCard = Nothing
        Response.Write "<tr class='tdbg'><td colspan='20' height='50' align='center'>没有任何充值卡！</td></tr></table>"
        Exit Sub
    End If
    
    totalPut = rsCard.RecordCount
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
            rsCard.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If
    i = 0
    Do While Not rsCard.EOF
        Response.Write "" & vbCrLf
        Response.Write "  <tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'""> " & vbCrLf
        Response.Write "    <td width='30'><input name='CardID' type='checkbox' id='CardID' value='" & rsCard("CardID") & "'  onclick='unselectall()'"
        If rsCard("UserName") <> "" Then Response.Write " disabled"
        Response.Write "></td>" & vbCrLf
        Response.Write "    <td width='60'>"
        If rsCard("CardType") = 0 Then
            Response.Write "本站充值卡"
        Else
            Response.Write "<font color='blue'>其他公司卡</font>"
        End If
        Response.Write "</td>"
        Response.Write "    <td width='100'><a href='Admin_Card.asp?Action=Show&CardID=" & rsCard("CardID") & "'>" & rsCard("CardNum") & "</a></td>" & vbCrLf
        Response.Write "    <td width='40'>" & rsCard("Money") & "元</td>" & vbCrLf
        Response.Write "    <td width='40'>" & GetValidNum(rsCard("ValidNum"), rsCard("ValidUnit")) & arrCardUnit(rsCard("ValidUnit")) & "</td>" & vbCrLf
        Response.Write "    <td width='60'>" & rsCard("EndDate") & "</td>" & vbCrLf
        Response.Write "    <td>"
        If IsNull(rsCard("ProductName")) Then
            Response.Write "<font color='blue'>不通过商城销售</font>"
        Else
            Response.Write rsCard("ProductName")
        End If
        Response.Write "</td>" & vbCrLf
        Response.Write "    <td width='40'>"
        If rsCard("UserName") <> "" Then
            Response.Write "<font color='gray'>已使用</font>"
        Else
            If rsCard("OrderFormItemID") > 0 Then
                Response.Write "已售出"
            Else
                If rsCard("EndDate") < Date Then
                    Response.Write "<font color='red'>已失效</font>"
                Else
                    If rsCard("ProductID") > 0 Then
                        Response.Write "<font color='green'>未售出</font>"
                    Else
                        Response.Write "<font color='green'>未使用</font>"
                    End If
                End If
            End If
        End If
        Response.Write "    </td>" & vbCrLf
        Response.Write "    <td width='60'><a href='Admin_User.asp?Action=Show&UserName=" & rsCard("UserName") & "'>" & rsCard("UserName") & "</a></td>" & vbCrLf
        Response.Write "    <td width='120'>" & rsCard("UseTime") & "</td>" & vbCrLf
        Response.Write "    <td width='60'><a href='Admin_Card.asp?AgentName=" & rsCard("AgentName") & "'>" & rsCard("AgentName") & "</a></td>" & vbCrLf
        Response.Write "    <td width='60'>"
        If rsCard("UserName") = "" And rsCard("OrderFormItemID") = 0 Then
            Response.Write "<a href='Admin_Card.asp?Action=Modify&CardID=" & rsCard("CardID") & "'>修改</a> "
            Response.Write "<a href='Admin_Card.asp?Action=Del&CardID=" & rsCard("CardID") & "' onclick=""return confirm('确定要删除此充值卡吗？')"">删除</a>"
        End If
        Response.Write "    </td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        i = i + 1
        If i >= MaxPerPage Then Exit Do
        rsCard.MoveNext
    Loop
    rsCard.Close
    Set rsCard = Nothing
    Response.Write "</table>  " & vbCrLf
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>" & vbCrLf
    Response.Write "              选中本页显示的所有充值卡</td>" & vbCrLf
    Response.Write "            <td><input name='Action' type='hidden' id='Action' value='Del'>" & vbCrLf
    Response.Write "              <input name='Submit' type='submit' id='Submit' value='删除选中的充值卡' onclick=""document.myform.Action.value='Del';return confirm('确定要删除选中的充值卡吗？');""></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</td>" & vbCrLf
    Response.Write "</form></tr></table>" & vbCrLf
    Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "张充值卡", True)
    Response.Write "<br>" & vbCrLf
    Response.Write "<form method='Get' name='SearchForm' action='Admin_Card.asp'>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='80' align='right'><strong>充值卡搜索：</strong></td>"
    Response.Write "   <td>"
    Response.Write "<select name='CardType'><option value='-1' selected>充值卡类型</option><option value='0'>本站充值卡</option><option value='1'>其他公司卡</option></select>"
    Response.Write "<select name='CardStatus'><option value='-1' selected>充值卡状态</option><option value='1'>未使用</option><option value='2'>已使用</option><option value='3'>已失效</option></select>"
    Response.Write "<select name='Field'><option value='CardNum'>卡号</option><option value='Money'>面值</option><option value='AgentName'>代理商</option><option value='UserName'>使用者</option></select>"
    Response.Write "<input type='text' name='keyword'  size='20' value='' maxlength='50' onFocus='this.select();'>"
    Response.Write "<input type='submit' name='Submit'  value='搜索'>"
    Response.Write "</td></tr></table></form>"
    Response.Write "<br>" & vbCrLf
    Response.Write "<table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td class='title'>小贴士</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height='50' class='tdbg'>" & vbCrLf
    Response.Write "      <li>通过商城销售的充值卡有四种状态：未售出、已售出、已使用、已失效</li>" & vbCrLf
    Response.Write "      <li><font color='blue'>不通过商城销售</font>的充值卡有三种状态：未使用、已使用、已失效</li>" & vbCrLf
    Response.Write "      <li>已经售出或已经使用过的充值卡将不能修改和删除</li>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub

Sub BatchAdd()
    Response.Write "<form method='post' action='Admin_Card.asp' name='myform'>" & vbCrLf
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >" & vbCrLf
    Response.Write "    <tr class='title'> " & vbCrLf
    Response.Write "      <td height='22' colspan='2'> <div align='center'><strong>批 量 生 成 充 值 卡</strong></div></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='40%'><strong>充值卡所属商品：</strong><br>商城中的某张点卡类商品可以对应多张实际的充值卡，会员在购买点卡类商品后，可以通过“获取虚拟充值卡”来得到这里输入的卡号和密码。</td>" & vbCrLf
    Response.Write "      <td width='60%'><select name='ProductID'><option value='0'>不通过商城销售</option>"
    Dim rsProduct
    Set rsProduct = Conn.Execute("select ProductID,ProductName from PE_Product where ProductKind=3 order by ProductID")
    Do While Not rsProduct.EOF
        Response.Write "<option value='" & rsProduct(0) & "'>" & rsProduct(1) & "</option>"
        rsProduct.MoveNext
    Loop
    Set rsProduct = Nothing
    Response.Write "</select>"
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='40%'><strong>充值卡数量：</strong></td>" & vbCrLf
    Response.Write "      <td width='60%'><input name='Nums' type='text' value='100' size='10' maxlength='10'>" & vbCrLf
    Response.Write "        张</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    'Response.Write "    <tr class='tdbg'>" & vbCrLf
    'Response.Write "      <td width='40%'><strong>充值卡号码前缀：</strong><br>例如：2008,PE2008等固定不变的字母或数字</td>" & vbCrLf
    'Response.Write "      <td width='60%'><input name='CardNumPrefix' type='text' id='CardNumPrefix' value='2008' size='10' maxlength='10'></td>" & vbCrLf
    'Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%'><strong>充值卡号码规则：</strong><br><span style='color:#0000ff'>说明：每个?代表一个英文字母，#代表一个数字，<br />                  *代表一个英文字母或数字(自定义符号必须是半角)</span></td>" & vbCrLf
    Response.Write "      <td width='60%'><input name='CardNumStr' type='text' id='CardNumStr' value='PE???###?#*' size='15' maxlength='15'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%'><strong>充值卡密码规则：</strong><br><span style='color:#0000ff'>说明：每个?代表一个英文字母，#代表一个数字，<br />                  *代表一个英文字母或数字(自定义符号必须是半角)</span></td>" & vbCrLf
    Response.Write "      <td width='60%'><input name='PasswordStr' type='text' id='PasswordStr' value='PE###?#*' size='10' maxlength='10'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%'><strong>充值卡面值：</strong><br>" & vbCrLf
    Response.Write "      即购买人需要花费的实际金额</td>" & vbCrLf
    Response.Write "      <td width='60%'><input name='Money' type='text' id='Money' value='50' size='10'>" & vbCrLf
    Response.Write "      元 <font color='red'>注意要与所属商品中描述的面值相同。</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='40%'><strong>充值卡点数、资金或有效期：</strong><br>" & vbCrLf
    Response.Write "        购买人可以得到的点数、资金或有效期" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td width='60%'><input name='ValidNum' type='text' id='ValidNum' value='500' size='10' maxlength='10'>" & vbCrLf
    Response.Write "        <select name='ValidUnit' id='ValidUnit' onchange='javascript:ShowGroup();'>" & vbCrLf
    Response.Write "          <option value='0' selected>点</option>" & vbCrLf
    Response.Write "          <option value='1'>天</option>" & vbCrLf
    Response.Write "          <option value='2'>月</option>" & vbCrLf
    Response.Write "          <option value='3'>年</option>" & vbCrLf
    Response.Write "          <option value='4'>元</option>" & vbCrLf
    Response.Write "          <option value='5'>组</option>" & vbCrLf
    Response.Write "        </select>"

    Response.Write "        <select name='GroupList' onchange='javascript:SetNumValue();' id='GroupList' style='display:none'>" & vbCrLf
    Dim rsGroupList
    Set rsGroupList = Conn.Execute("Select GroupID,GroupName from PE_UserGroup Order by GroupID asc")
    Do While Not rsGroupList.EOF
        Response.Write "         <option value='" & rsGroupList("GroupID") & "'>" & rsGroupList("GroupName") & "</option>" & vbCrLf
        rsGroupList.MoveNext
    Loop
    rsGroupList.Close
    Set rsGroupList = Nothing
    Response.Write "         </select>"
            
    Response.Write "<span id='HelpInfoForPoint'><font color='red'>注意要与所属商品中描述的点数相同。</font></span>"
    Response.Write "<span id='HelpInfoForGroup' style='display:none'><font color='red'>请选择充值卡对应的会员组。</font></span></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%'><strong>充值截止期限：</strong><br>" & vbCrLf
    Response.Write "      购买人必须在此日期前进行充值，否则自动失效</td>" & vbCrLf
    Response.Write "      <td width='60%' class='tdbg'><input name='EndDate' type='text' id='EndDate' value='" & DateAdd("yyyy", 1, Date) & "' size='10' maxlength='10'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='40%'><strong>代理商：</strong></td>" & vbCrLf
    Response.Write "      <td width='60%'><input name='AgentName' type='text' value='' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td height='40' colspan='2' align='center'><input name='Action' type='hidden' id='Action' value='DoBatchAdd'> " & vbCrLf
    Response.Write "        <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'> " & vbCrLf
    Response.Write "        &nbsp; <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='Admin_Card.asp'"" style='cursor:hand;'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

Sub Add()
    Response.Write "<form name='myform' method='post' action='Admin_Card.asp'>" & vbCrLf
    Response.Write "  <table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center'>" & vbCrLf
    Response.Write "      <td colspan='2' class='title'><strong>添 加 充 值 卡</strong></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='40%'><strong>充值卡类型：</strong></td>" & vbCrLf
    Response.Write "      <td width='60%'><input name='CardType' type='radio' value='0' checked>本站充值卡&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>购买者得到卡号和密码后，可以直接在本站进行充值</font><br><input name='CardType' type='radio' value='1'>其他公司卡&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>购买者得到卡号和密码后，需要去相关公司或网站进行充值</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='40%'><b>充值卡所属商品：</b><br>商城中的某张点卡类商品可以对应多张实际的充值卡，会员在购买点卡类商品后，可以通过“获取虚拟充值卡”来得到这里输入的卡号和密码。</td>" & vbCrLf
    Response.Write "      <td width='60%'><select name='ProductID'><option value='0'>不通过商城销售</option>"
    Dim rsProduct
    Set rsProduct = Conn.Execute("select ProductID,ProductName from PE_Product where ProductKind=3 order by ProductID")
    Do While Not rsProduct.EOF
        Response.Write "<option value='" & rsProduct(0) & "'>" & rsProduct(1) & "</option>"
        rsProduct.MoveNext
    Loop
    Set rsProduct = Nothing
    Response.Write "</select>"
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='40%'><strong>添加方式：</strong></td>" & vbCrLf
    Response.Write "      <td width='60%'><input name='AddType' type='radio' value='0' checked onclick=""trSingle1.style.display='';trSingle2.style.display='';trBatch.style.display='none';""> 单张充值卡&nbsp;&nbsp;&nbsp;&nbsp;<input name='AddType' type='radio' value='1' onclick=""trSingle1.style.display='none';trSingle2.style.display='none';trBatch.style.display='';"">批量添加充值卡</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    
    Response.Write "    <tr class='tdbg' id='trSingle1'>" & vbCrLf
    Response.Write "      <td width='40%'><b>充值卡卡号：</b></td>" & vbCrLf
    Response.Write "      <td><input name='CardNum' type='text' id='CardNum' size='20' maxlength='30'>" & vbCrLf
    Response.Write "        <font color='#0000FF'>建议设为10--15位</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='trSingle2'>" & vbCrLf
    Response.Write "      <td width='40%'><b>充值卡密码：</b></td>" & vbCrLf
    Response.Write "      <td><input name='Password' type='text' id='Password' size='20' maxlength='30'>" & vbCrLf
    Response.Write "        <font color='#0000FF'>建议设为6--10位 </font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    
    Response.Write "    <tr class='tdbg' id='trBatch' style='display:none'>" & vbCrLf
    Response.Write "      <td width='40%'><b>格式文本：</b><br><font color='red'>请按照每行一张卡，每张卡按“卡号＋分隔符＋密码”的格式录入</font><br>例1：734534759*kSo94Sf4Xs（以“*”作为分隔符）<br>例2：98273305834|lo23ji6x（以“|”作为分隔符）</td>" & vbCrLf
    Response.Write "      <td><textarea name='CardText' rows='10' cols='50'></textarea><br>分隔符：<input name='strSplit' type='text' id='strSplit' size='20' maxlength='30'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%'><b>充值卡面值：</b><br>即购买人需要花费的实际金额</td>" & vbCrLf
    Response.Write "      <td><input name='Money' type='text' id='Money' size='10' maxlength='10'>" & vbCrLf
    Response.Write "        元</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%'><b>充值卡点数、资金或有效期：</b><br>购买人可以得到的点数、资金或有效期</td>" & vbCrLf
    Response.Write "      <td><input name='ValidNum' type='text' id='ValidNum' size='10' maxlength='10'>" & vbCrLf
    Response.Write "          <select name='ValidUnit' id='ValidUnit' onchange='javascript:ShowGroup();'>" & vbCrLf
    Response.Write "            <option value='0' selected>点</option>" & vbCrLf
    Response.Write "            <option value='1'>天</option>" & vbCrLf
    Response.Write "            <option value='2'>月</option>" & vbCrLf
    Response.Write "            <option value='3'>年</option>" & vbCrLf
    Response.Write "            <option value='4'>元</option>" & vbCrLf
    Response.Write "            <option value='5'>组</option>" & vbCrLf
    Response.Write "          </select>" & vbCrLf

    Response.Write "        <select name='GroupList' onchange='javascript:SetNumValue();' id='GroupList' style='display:none'>" & vbCrLf
    Dim rsGroupList
    Set rsGroupList = Conn.Execute("Select GroupID,GroupName from PE_UserGroup Order by GroupID asc")
    Do While Not rsGroupList.EOF
        Response.Write "         <option value='" & rsGroupList("GroupID") & "'>" & rsGroupList("GroupName") & "</option>" & vbCrLf
        rsGroupList.MoveNext
    Loop
    rsGroupList.Close
    Set rsGroupList = Nothing
    Response.Write "         </select>"
            
    Response.Write "<span id='HelpInfoForPoint'><font color='red'>注意要与所属商品中描述的点数相同。</font></span>"
    Response.Write "<span id='HelpInfoForGroup' style='display:none'><font color='red'>请选择充值卡对应的会员组。</font></span></td>" & vbCrLf
    Response.Write "    </td></tr>" & vbCrLf

    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%'><b>充值截止日期：</b><br>购买人必须在此日期前进行充值，否则自动失效</td>" & vbCrLf
    Response.Write "      <td><input name='EndDate' type='text' id='EndDate' value='" & DateAdd("yyyy", 1, Date) & "' size='10' maxlength='10'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='40%'><strong>代理商：</strong></td>" & vbCrLf
    Response.Write "      <td width='60%'><input name='AgentName' type='text' value='' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <td height='30' colspan='2'><input name='Action' type='hidden' id='Action' value='SaveAdd'>" & vbCrLf
    Response.Write "          <input type='submit' name='Submit' value=' 添 加 '></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub

Sub Modify()
    Dim CardID, rsCard
    CardID = Trim(Request("CardID"))
    If CardID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要删除的充值卡ID</li>"
        Exit Sub
    Else
        CardID = PE_CLng(CardID)
    End If
    Set rsCard = Conn.Execute("select * from PE_Card where CardID=" & CardID & "")
    If rsCard.Bof And rsCard.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的充值卡！</li>"
    Else
        If rsCard("UserName") <> "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>此充值卡已经被使用，不能再修改！</li>"
        End If
        If rsCard("OrderFormItemID") > 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>此充值卡已经售出，不能再修改！</li>"
        End If
    End If
    If FoundErr = True Then
        Set rsCard = Nothing
        Exit Sub
    End If
    Response.Write "<form name='myform' method='post' action='Admin_Card.asp'>" & vbCrLf
    Response.Write "  <table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center'>" & vbCrLf
    Response.Write "      <td colspan='2' class='title'><strong>修 改 充 值 卡</strong></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>所属商品：</b></td>" & vbCrLf
    Response.Write "      <td>" & GetProductName(rsCard("ProductID")) & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>充值卡卡号：</b></td>" & vbCrLf
    Response.Write "      <td><input name='CardNum' type='text' id='CardNum' value='" & rsCard("CardNum") & "' size='20' maxlength='30' disabled></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>充值卡密码：</b></td>" & vbCrLf
    Response.Write "      <td><input name='Password' type='text' id='Password' value='" & Base64decode(rsCard("Password")) & "' size='20' maxlength='30'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>充值卡面值：</b></td>" & vbCrLf
    Response.Write "      <td><input name='Money' type='text' id='Money' value='" & rsCard("Money") & "' size='10' maxlength='10'>" & vbCrLf
    Response.Write "      元</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>充值卡点数：</b></td>" & vbCrLf
    Response.Write "      <td><input name='ValidNum'"
    If rsCard("ValidUnit") = 5 Then Response.Write " disabled "
    Response.Write "type='text' id='ValidNum' value='" & rsCard("ValidNum") & "' size='10' maxlength='10'>" & vbCrLf
    Response.Write "        <select name='ValidUnit' id='ValidUnit' onchange='javascript:ShowGroup();'>" & vbCrLf
    Response.Write "          <option value='0'"
    If rsCard("ValidUnit") = 0 Then Response.Write " selected"
    Response.Write ">点</option>" & vbCrLf
    Response.Write "          <option value='1'"
    If rsCard("ValidUnit") = 1 Then Response.Write " selected"
    Response.Write ">天</option>" & vbCrLf
    Response.Write "          <option value='2'"
    If rsCard("ValidUnit") = 2 Then Response.Write " selected"
    Response.Write ">月</option>" & vbCrLf
    Response.Write "          <option value='3'"
    If rsCard("ValidUnit") = 3 Then Response.Write " selected"
    Response.Write ">年</option>" & vbCrLf
    Response.Write "          <option value='4'"
    If rsCard("ValidUnit") = 4 Then Response.Write " selected"
    Response.Write ">元</option>" & vbCrLf
    Response.Write "          <option value='5'"
    If rsCard("ValidUnit") = 5 Then Response.Write " selected"
    Response.Write ">组</option>" & vbCrLf
    Response.Write "        </select>"

    Response.Write "        <select name='GroupList' onchange='javascript:SetNumValue();' id='GroupList'"
    If rsCard("ValidUnit") <> 5 Then Response.Write " style='display:none'"
    Response.Write " >" & vbCrLf
    Dim rsGroupList
    Set rsGroupList = Conn.Execute("Select GroupID,GroupName from PE_UserGroup Order by GroupID asc")
    Do While Not rsGroupList.EOF
        Response.Write "         <option value='" & rsGroupList("GroupID") & "'"
        If rsCard("ValidNum") = rsGroupList("GroupID") Then Response.Write " selected"
        Response.Write " >" & rsGroupList("GroupName") & "</option>" & vbCrLf
        rsGroupList.MoveNext
    Loop
    rsGroupList.Close
    Set rsGroupList = Nothing
    Response.Write "         </select>"
            
    Response.Write "<span id='HelpInfoForPoint'"
    If rsCard("ValidUnit") = 5 Then Response.Write " style='display:none'"
    Response.Write "><font color='red'>注意要与所属商品中描述的点数相同。</font></span>"
    Response.Write "<span id='HelpInfoForGroup'"
    If rsCard("ValidUnit") <> 5 Then Response.Write "style='display:none'"
    Response.Write "><font color='red'>请选择充值卡对应的会员组。</font></span></td>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>充值截止日期：</b></td>" & vbCrLf
    Response.Write "      <td><input name='EndDate' type='text' id='EndDate' value='" & rsCard("EndDate") & "' size='10' maxlength='10'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='40%' align='right'><strong>代理商：</strong></td>" & vbCrLf
    Response.Write "      <td width='60%'><input name='AgentName' type='text' value='" & rsCard("AgentName") & "' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr align='center' class='tdbg'>" & vbCrLf
    Response.Write "      <td height='30' colspan='2'><input name='Action' type='hidden' id='Action' value='SaveModify'>" & vbCrLf
    Response.Write "      <input name='CardID' type='hidden' id='CardID' value='" & CardID & "'>" & vbCrLf
    Response.Write "      <input type='submit' name='Submit' value='保存修改结果'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    Set rsCard = Nothing
End Sub

Sub Show()
    Dim CardID, rsCard
    CardID = Trim(Request("CardID"))
    If CardID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要删除的充值卡ID</li>"
        Exit Sub
    Else
        CardID = PE_CLng(CardID)
    End If
    Set rsCard = Conn.Execute("select * from PE_Card where CardID=" & CardID & "")
    If rsCard.Bof And rsCard.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的充值卡！</li>"
    End If
    If FoundErr = True Then
        Set rsCard = Nothing
        Exit Sub
    End If
    Response.Write "<br><table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center'>" & vbCrLf
    Response.Write "      <td colspan='2' class='title'><strong>查 看 充 值 卡 信 息</strong></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>充值卡类型：</b></td>" & vbCrLf
    Response.Write "      <td>"
    If rsCard("CardType") = 0 Then
        Response.Write "本站充值卡"
    Else
        Response.Write "其他公司卡"
    End If
    Response.Write "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>所属商品：</b></td>" & vbCrLf
    Response.Write "      <td>" & GetProductName(rsCard("ProductID")) & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>充值卡卡号：</b></td>" & vbCrLf
    Response.Write "      <td>" & rsCard("CardNum") & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>充值卡密码：</b></td>" & vbCrLf
    Response.Write "      <td>" & Base64decode(rsCard("Password")) & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>充值卡面值：</b></td>" & vbCrLf
    Response.Write "      <td>" & rsCard("Money") & " 元</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>充值卡点数：</b></td>" & vbCrLf
    Response.Write "      <td>" & GetValidNum(rsCard("ValidNum"), rsCard("ValidUnit")) & arrCardUnit(rsCard("ValidUnit")) & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>充值截止日期：</b></td>" & vbCrLf
    Response.Write "      <td>" & rsCard("EndDate") & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>充值卡生成时间：</b></td>" & vbCrLf
    Response.Write "      <td>" & rsCard("CreateTime") & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>充值卡状态：</b></td>" & vbCrLf
    Response.Write "      <td>"

        If rsCard("UserName") <> "" Then
            Response.Write "已使用"
        Else
            If rsCard("OrderFormItemID") > 0 Then
                Response.Write "已售出"
            Else
                If rsCard("EndDate") < Date Then
                    Response.Write "<font color='red'>已失效</font>"
                Else
                    If rsCard("ProductID") > 0 Then
                        Response.Write "<font color='green'>未售出</font>"
                    Else
                        Response.Write "<font color='green'>未使用</font>"
                    End If
                End If
            End If
        End If
    Response.Write "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>使用者：</b></td>" & vbCrLf
    Response.Write "      <td>" & rsCard("UserName") & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' align='right'><b>充值时间：</b></td>" & vbCrLf
    Response.Write "      <td>" & rsCard("UseTime") & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'> " & vbCrLf
    Response.Write "      <td width='40%' align='right'><strong>代理商：</strong></td>" & vbCrLf
    Response.Write "      <td width='60%'>" & rsCard("AgentName") & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Set rsCard = Nothing
End Sub

Sub DoBatchAdd()
    Dim arrCardNum, arrPassword, CardNum, Password
    Dim ProductID, Nums, CardNumStr, PasswordStr, Money, ValidNum, ValidUnit, EndDate, AgentName
    arrCardNum = ""
    arrPassword = ""
    ProductID = PE_CLng(Trim(Request.Form("ProductID")))
    Nums = PE_CLng(Trim(Request.Form("Nums")))
    'CardNumPrefix = Trim(Request("CardNumPrefix"))
    CardNumStr = Trim(Request.Form("CardNumStr"))
    PasswordStr = Trim(Request.Form("PasswordStr"))
    Money = PE_CDbl(Trim(Request.Form("Money")))
    ValidUnit = PE_CLng(Trim(Request.Form("ValidUnit")))
    If ValidUnit = 5 Then
        ValidNum = PE_CLng(Trim(Request.Form("GroupList")))
    Else
        ValidNum = PE_CLng(Trim(Request.Form("ValidNum")))
    End If
    EndDate = Trim(Request.Form("EndDate"))
    AgentName = Trim(Request.Form("AgentName"))
    If Nums < 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要生成的充值卡数量！</li>"
    End If
    If CardNumStr = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定充值卡卡号规则！</li>"
    End If
    If PasswordStr = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定充值卡密码规则！</li>"
    End If
    If Money <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定充值卡的面值！</li>"
    End If
    If ValidNum <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定充值卡的点数！</li>"
    End If
    If IsDate(EndDate) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定正确的充值截止日期！</li>"
    Else
        EndDate = CDate(EndDate)
        If EndDate <= Date Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>充值截止日期不能比当前日期还早</li>"
        End If
    End If
    If FoundErr = True Then Exit Sub
    
    Dim sqlCard, rsCard, i
    sqlCard = "select top 1 * from PE_Card"
    Set rsCard = Server.CreateObject("adodb.recordset")
    rsCard.Open sqlCard, Conn, 1, 3
    For i = 1 To Nums
        CardNum = GetRndCharAndNum(CardNumStr)
        Password = GetRndCharAndNum(PasswordStr)
        If arrCardNum = "" Then
            arrCardNum = CardNum
            arrPassword = Password
        Else
            arrCardNum = arrCardNum & "," & CardNum
            arrPassword = arrPassword & "," & Password
        End If
        rsCard.AddNew
        rsCard("CardType") = 0
        rsCard("ProductID") = ProductID
        rsCard("CardNum") = CardNum
        rsCard("Password") = Base64encode(Password)
        rsCard("Money") = Money
        rsCard("ValidNum") = ValidNum
        rsCard("ValidUnit") = ValidUnit
        rsCard("EndDate") = EndDate
        rsCard("AgentName") = AgentName
        rsCard("UserName") = ""
        rsCard("CreateTime") = Now()
        rsCard("OrderFormItemID") = 0
        rsCard.Update
    Next
    rsCard.Close
    Set rsCard = Nothing
    If ProductID > 0 Then
        Conn.Execute ("update PE_Product set Stocks=Stocks+" & Nums & " where ProductID=" & ProductID & "")
    End If
    
    arrCardNum = Split(arrCardNum, ",")
    arrPassword = Split(arrPassword, ",")
    
    Response.Write "  <br>" & vbCrLf
    Response.Write "  <table width='300'  border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr class='title'>" & vbCrLf
    Response.Write "      <td colspan='2' align='center'><strong>本次生成的点卡信息如下：</strong></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='100'>充值卡所属商品：</td>" & vbCrLf
    Response.Write "      <td>" & GetProductName(ProductID) & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='100'>充值卡数量：</td>" & vbCrLf
    Response.Write "      <td>" & Nums & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='100'>充值卡面值：</td>" & vbCrLf
    Response.Write "      <td>" & Money & " 元</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='100'>充值卡点数：</td>" & vbCrLf
    Response.Write "      <td>" & GetValidNum(ValidNum, ValidUnit) & arrCardUnit(ValidUnit) & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='100'>充值截止日期：</td>" & vbCrLf
    Response.Write "      <td>" & EndDate & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='100'>代理商：</td>" & vbCrLf
    Response.Write "      <td>" & AgentName & "</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "</table><br>" & vbCrLf
    Response.Write "<table width='300' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr align='center' class='title'>" & vbCrLf
    Response.Write "    <td  width=150 height='22'><strong> 卡 号 </strong></td>" & vbCrLf
    Response.Write "    <td  width=150 height='22'><strong> 密 码 </strong></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    For i = 0 To Nums - 1
        Response.Write "  <tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
        Response.Write "    <td>" & arrCardNum(i) & "</td>" & vbCrLf
        Response.Write "    <td>" & arrPassword(i) & "</td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
    Next
    Response.Write "</table>" & vbCrLf
End Sub

'PE###?#*
'PE???###?#*
'每个?代表一个英文字母，
'#代表一个数字，
'*代表一个英文字母或数字(自定义符号必须是半角)
Function GetRndCharAndNum(str)
    Dim arrNum, arrChar, arrMix, strLen, strTemp, i, c
    arrNum = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
    arrChar = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    arrMix = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    Randomize
    strLen = Len(str)
    strTemp = ""
    For i = 1 To strLen + 1
        c = Mid(str, i, 1)
     '   Randomize
        Select Case c
        Case "?"
            '10, 62
            strTemp = strTemp & arrChar(CInt(Rnd * 51))
        Case "#"
            '0, 10
            strTemp = strTemp & arrNum(CInt(Rnd * 9))
        Case "*"
            '62
            strTemp = strTemp & arrMix(CInt(Rnd * 61))
        Case Else
            strTemp = strTemp & c
        End Select
    Next
    GetRndCharAndNum = strTemp
End Function

Sub SaveAdd()
    Dim rsCard
    Dim ProductID, CardType
    Dim AddType
    Dim CardNum, Password, Money, ValidNum, ValidUnit, EndDate, AgentName
    Dim CardText, strSplit
    
    CardType = PE_CLng(Trim(Request.Form("CardType")))
    ProductID = PE_CLng(Trim(Request.Form("ProductID")))
    AddType = PE_CLng(Trim(Request.Form("AddType")))
    
    CardNum = ReplaceBadChar(Trim(Request.Form("CardNum")))
    Password = ReplaceBadChar(Trim(Request.Form("Password")))
    
    CardText = Trim(Request.Form("CardText"))
    strSplit = Trim(Request.Form("strSplit"))
    
    Money = PE_CDbl(Trim(Request.Form("Money")))
    ValidUnit = PE_CLng(Trim(Request.Form("ValidUnit")))
    If ValidUnit = 5 Then
        ValidNum = PE_CLng(Trim(Request.Form("GroupList")))
    Else
        ValidNum = PE_CLng(Trim(Request.Form("ValidNum")))
    End If
    EndDate = Trim(Request.Form("EndDate"))
    AgentName = Trim(Request.Form("AgentName"))
    If CardType = 1 And ProductID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>其他公司卡必须通过商城销售。请指定所属商品。</li>"
    End If
    If AddType = 0 Then
        If CardNum = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定充值卡ID</li>"
        End If
        If Password = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定充值卡密码</li>"
        End If
    Else
        If CardText = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请输入批量添加的充值卡格式文本</li>"
        End If
        If strSplit = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定分隔符</li>"
        End If
    End If
    If Money <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定充值卡的面值！</li>"
    End If
    If ValidNum <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定充值卡的点数！</li>"
    End If
    If IsDate(EndDate) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定正确的充值截止日期！</li>"
    Else
        EndDate = CDate(EndDate)
        If EndDate <= Date Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>充值截止日期不能比当前日期还早</li>"
        End If
    End If
    If FoundErr = True Then Exit Sub
    
    If AddType = 0 Then
        Set rsCard = Server.CreateObject("Adodb.Recordset")
        rsCard.Open "select * from PE_Card where CardNum='" & CardNum & "' and ProductID=" & ProductID & "", Conn, 1, 3
        If Not (rsCard.Bof And rsCard.EOF) Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>输入的充值卡卡号已经存在！</li>"
        End If
        If FoundErr = True Then
            Set rsCard = Nothing
            Exit Sub
        End If
        rsCard.AddNew
        rsCard("CardType") = CardType
        rsCard("ProductID") = ProductID
        rsCard("CardNum") = CardNum
        rsCard("Password") = Base64encode(Password)
        rsCard("Money") = Money
        rsCard("ValidNum") = ValidNum
        rsCard("ValidUnit") = ValidUnit
        rsCard("EndDate") = EndDate
        rsCard("AgentName") = AgentName
        rsCard("UserName") = ""
        rsCard("CreateTime") = Now()
        rsCard("OrderFormItemID") = 0
        rsCard.Update
        rsCard.Close
        Set rsCard = Nothing
        If ProductID > 0 Then
            Conn.Execute ("update PE_Product set Stocks=Stocks+1 where ProductID=" & ProductID & "")
        End If
        Call CloseConn
        Response.Redirect "Admin_Card.asp"
    Else
        Dim arrCard, arrCard2, i, iCount
        arrCard = Split(CardText, vbCrLf)
        iCount = 0
        Set rsCard = Server.CreateObject("Adodb.Recordset")
        For i = 0 To UBound(arrCard)
            If Trim(arrCard(i)) <> "" Then
                If InStr(arrCard(i), strSplit) <> 0 Then
                    arrCard2 = Split(Trim(arrCard(i)), strSplit)
                    CardNum = ReplaceBadChar(Trim(arrCard2(0)))
                    If CardNum <> "" And Trim(arrCard2(1)) <> "" Then
                        rsCard.Open "select * from PE_Card where CardNum='" & CardNum & "' and ProductID=" & ProductID & "", Conn, 1, 3
                        If rsCard.Bof And rsCard.EOF Then
                            rsCard.AddNew
                            rsCard("CardType") = CardType
                            rsCard("ProductID") = ProductID
                            rsCard("CardNum") = CardNum
                            rsCard("Password") = Base64encode(Trim(arrCard2(1)))
                            rsCard("Money") = Money
                            rsCard("ValidNum") = ValidNum
                            rsCard("ValidUnit") = ValidUnit
                            rsCard("EndDate") = EndDate
                            rsCard("UserName") = ""
                            rsCard("CreateTime") = Now()
                            rsCard("OrderFormItemID") = 0
                            rsCard.Update
                            rsCard.Close
                            iCount = iCount + 1
                            Response.Write "<li>卡号为：" & CardNum & " 的充值卡成功添加到数据库中！</li>"
                        Else
                            Response.Write "<li>卡号为：" & CardNum & " 的充值卡已经存在！</li>"
                            rsCard.Close
                        End If
                    End If
                Else
                    Response.Write "<li>添加的第" & i + 1 & "条卡号信息有误"
                End If
            End If
            Response.Flush
        Next
        If ProductID > 0 Then
            Conn.Execute ("update PE_Product set Stocks=Stocks+" & iCount & " where ProductID=" & ProductID & "")
        End If
        Set rsCard = Nothing
    End If
End Sub

Sub SaveModify()
    Dim CardID, rsCard
    Dim Password, Money, ValidNum, ValidUnit, EndDate
    Password = ReplaceBadChar(Trim(Request.Form("Password")))
    Money = PE_CDbl(Trim(Request.Form("Money")))
    ValidUnit = PE_CLng(Trim(Request.Form("ValidUnit")))
    If ValidUnit = 5 Then
        ValidNum = PE_CLng(Trim(Request.Form("GroupList")))
    Else
        ValidNum = PE_CLng(Trim(Request.Form("ValidNum")))
    End If
    EndDate = Trim(Request.Form("EndDate"))
    CardID = PE_CLng(Trim(Request("CardID")))
    If CardID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定充值卡ID</li>"
    End If
    If Money <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定充值卡的面值！</li>"
    End If
    If Password = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定充值卡的密码！</li>"
    End If
    If ValidNum <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定充值卡的点数！</li>"
    End If
    If IsDate(EndDate) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定正确的充值截止日期！</li>"
    Else
        EndDate = CDate(EndDate)
        If EndDate <= Date Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>充值截止日期不能比当前日期还早</li>"
        End If
    End If
    If FoundErr = True Then Exit Sub

    Set rsCard = Server.CreateObject("Adodb.Recordset")
    rsCard.Open "select * from PE_Card where CardID=" & CardID & "", Conn, 1, 3
    If rsCard.Bof And rsCard.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的充值卡！</li>"
    Else
        If rsCard("UserName") <> "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>此充值卡已经被使用，不能再修改！</li>"
        End If
        If rsCard("OrderFormItemID") > 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>此充值卡已经售出，不能再修改！</li>"
        End If
    End If
    If FoundErr = True Then
        Set rsCard = Nothing
        Exit Sub
    End If
    rsCard("Password") = Base64encode(Password)
    rsCard("Money") = Money
    rsCard("ValidNum") = ValidNum
    rsCard("ValidUnit") = ValidUnit
    rsCard("EndDate") = EndDate
    rsCard.Update
    rsCard.Close
    Set rsCard = Nothing
    Call CloseConn
    Response.Redirect "Admin_Card.asp"
End Sub

Sub DelCard()
    Dim CardID, rsCard
    CardID = Trim(Request("CardID"))
    If CardID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要删除的充值卡ID</li>"
        Exit Sub
    Else
        If IsValidID(CardID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定正确的充值卡ID</li>"
            Exit Sub
        End If
    End If
    If InStr(CardID, ",") > 0 Then
        Set rsCard = Conn.Execute("select CardID,ProductID from PE_Card where CardID in (" & CardID & ") and UserName='' and OrderFormItemID=0")
        Do While Not rsCard.EOF
            If rsCard("ProductID") > 0 Then
                Conn.Execute ("update PE_Product set Stocks=Stocks-1 where ProductID=" & rsCard("ProductID") & "")
            End If
            rsCard.MoveNext
        Loop
        Set rsCard = Nothing
        
        Conn.Execute ("delete from PE_Card where CardID in (" & CardID & ") and UserName='' and OrderFormItemID=0")
    Else
        Set rsCard = Conn.Execute("select CardID,ProductID from PE_Card where CardID =" & CardID & " and UserName='' and OrderFormItemID=0")
        Do While Not rsCard.EOF
            If rsCard("ProductID") > 0 Then
                Conn.Execute ("update PE_Product set Stocks=Stocks-1 where ProductID=" & rsCard("ProductID") & "")
            End If
            rsCard.MoveNext
        Loop
        Set rsCard = Nothing
        Conn.Execute ("delete from PE_Card where CardID=" & CardID & " and UserName='' and OrderFormItemID=0")
    End If
    Call main
End Sub

Function GetProductName(iProductID)
    If iProductID = 0 Then
        GetProductName = "不通过商城销售"
    Else
        Dim rsProduct
        Set rsProduct = Conn.Execute("select ProductName from PE_Product where ProductID=" & iProductID & "")
        If rsProduct.Bof And rsProduct.EOF Then
            GetProductName = "找不到所属商品"
        Else
            GetProductName = rsProduct(0)
        End If
        Set rsProduct = Nothing
    End If
End Function

Function GetValidNum(intValidNum, intValidUnit)
    If intValidUnit = 5 Then
        Dim rsGroupList
        Set rsGroupList = Conn.Execute("Select GroupName from PE_UserGroup where GroupID = " & intValidNum)
        If Not (rsGroupList.EOF And rsGroupList.Bof) Then
            GetValidNum = rsGroupList("GroupName")
        Else
            GetValidNum = intValidNum
        End If
        rsGroupList.Close
        Set rsGroupList = Nothing
    Else
        GetValidNum = intValidNum
    End If
End Function
%>
