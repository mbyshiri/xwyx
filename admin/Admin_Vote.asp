<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_ContentEx.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = "Vote"   '其他权限

Dim ItemName, ID, VoteType

VoteType = PE_CLng(Trim(Request("VoteType")))
strFileName = "Admin_Vote.asp?Action=" & Action & "&VoteType=" & VoteType
ItemName = "调查"
ID = Trim(Request("ID"))
ChannelID = PE_CLng(Trim(Request("ChannelID")))
If IsValidID(ID) = False Then
    ID = ""
End If

Response.Write "<html><head><title>调查管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Call ShowPageTitle("网 站 调 查 管 理", 10024)
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td width='70' height='30'><strong>管理导航：</strong></td>" & vbCrLf
Response.Write "    <td><a href='Admin_Vote.asp'>调查管理首页</a>&nbsp;|&nbsp;<a href='Admin_Vote.asp?Action=Add'>添加新调查</a>"
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
Select Case Action
Case "Add"
    Call Add
Case "Modify"
    Call Modify
Case "SaveAdd", "SaveModify"
    Call SaveVote
Case "SetNew", "CancelNew", "Move", "Del"
    Call SetProperty
Case Else
    Call main
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim rs, sql
    Call ShowJS_Main(ItemName)
    Response.Write "<br><table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22'>" & GetChannelList(ChannelID) & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table><br>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "    <td height='22'>"
    Call ShowManagePath(ChannelID)
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "  <form name='myform' method='Post' action='Admin_Vote.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "  <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td width='30' height='22' align='center'><strong>选中</strong></td>"
    Response.Write "    <td width='30' height='22' align='center'><strong>ID</strong></td>"
    Response.Write "    <td height='22' align='center'><strong>主题</strong></td>"
    Response.Write "    <td width='60' height='22' align='center'><strong>调查状态</strong></td>"
    Response.Write "    <td width='60' height='22' align='center'><strong>调查类型</strong></td>"
    Response.Write "    <td width='120' height='22' align='center'><strong>发布时间</strong></td>"
    Response.Write "    <td width='120' height='22' align='center'><strong>终止时间</strong></td>"
    Response.Write "    <td width='80' height='22' align='center'><strong>操作</strong></td>"
    Response.Write "  </tr>"

    sql = "select * from PE_Vote"
    If VoteType = 0 Then
        sql = sql & " where IsItem=" & PE_False
    Else
        sql = sql & " where IsItem=" & PE_True
    End If
    If ChannelID >= -1 Then
        sql = sql & " and ChannelID=" & ChannelID
    End If
    sql = sql & " order by IsSelected,ID desc"
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>没有任何调查！<br><br></td></tr>"
    Else
        Do While Not rs.EOF
            Response.Write "    <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "      <td width='30' align='center'><input name='ID' type='checkbox' onclick='unselectall()' value='" & rs("ID") & "'></td>"
            Response.Write "      <td width='30' align='center'>" & rs("ID") & "</td>"
            If VoteType = 0 Then
                Response.Write "      <td><a href='Admin_Vote.asp?Action=Modify&ID=" & rs("ID") & "'>" & rs("Title") & "</a></td>"
            Else
                Response.Write "      <td><a href='" & InstallDir & "Vote.asp?ID=" & rs("ID") & "&Action=Show'>" & rs("Title") & "</a></td>"
            End If
            Response.Write "      <td width='60' align='center'>"
            If rs("IsSelected") = True And Now() <= rs("EndTime") Then
                Response.Write "<font color=green>启用</font>"
            Else
                If Now() > rs("EndTime") Then
                    Response.Write "<font color=red>过期</font>"
                Else
                    Response.Write "<font color=red>停止</font>"
                End If
            End If
            Response.Write "      </td>"
            Response.Write "      <td width='60' align='center'>"
            If rs("VoteType") = "Single" Then
                Response.Write "单选"
            ElseIf rs("VoteType") = "Multi" Then
                Response.Write "多选"
            End If
            Response.Write "      <td align='center'>" & rs("VoteTime") & "</td>"
            Response.Write "      <td align='center'>" & rs("EndTime") & "</td>"
            Response.Write "      <td width='80' align='center'>"
            If VoteType = 0 Then
                Response.Write "      <a href='Admin_Vote.asp?Action=Modify&ID=" & rs("ID") & "'>修改</a>&nbsp;"
            End If
            Response.Write "      <a href='Admin_Vote.asp?Action=Del&ID=" & rs("ID") & "' onClick=""return confirm('确定要删除此调查吗？');"">删除</a>&nbsp;"
            Response.Write "      </td>"
            Response.Write "    </tr>"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='130' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>选中所有的调查</td><td>"
    Response.Write "<input type='submit' value='删除选定的调查' name='submit' onClick=""document.myform.Action.value='Del'"">&nbsp;&nbsp;"
    If VoteType = 0 Then
        Response.Write "<input type='submit' value='启用调查' name='submit1' onClick=""document.myform.Action.value='SetNew'"">&nbsp;&nbsp;"
        Response.Write "<input type='submit' value='停止调查' name='submit2' onClick=""document.myform.Action.value='CancelNew'"">&nbsp;&nbsp;"
        Response.Write "<input type='submit' value='将选定的调查移动到 ->' name='submit3' onClick=""document.myform.Action.value='Move'"">"
    End If
    Response.Write "<select name='ChannelID' id='ChannelID'>" & GetChannel_Option(0) & "</select>"
    Response.Write "<input name='Action' type='hidden' id='Action' value='ExportExcel'>"
    Response.Write "<input name='VoteType' type='hidden' id='VoteType' value='" & VoteType & "'>"
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    If VoteType = 0 Then
        Response.Write "<br><b>说明：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;只有将调查设为最新调查后才会在前台显示"
    End If
    Response.Write "<br><br>"
End Sub

Sub ShowJS_AddModify()
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if (document.myform.Title.value==''){" & vbCrLf
    Response.Write "     alert('调查主题不能为空！');" & vbCrLf
    Response.Write "     document.myform.Title.focus();" & vbCrLf
    Response.Write "     return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  return true;  " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function SetVoteNum(type){" & vbCrLf
    Response.Write "  if (type=='Multi'){" & vbCrLf
    Response.Write "     document.getElementById('VoteNum').style.display='';" & vbCrLf
    Response.Write "     document.getElementById('VoteNumtips').style.display='';" & vbCrLf	
    Response.Write "  }" & vbCrLf
    Response.Write "  else{" & vbCrLf
    Response.Write "     document.getElementById('VoteNum').style.display='None';" & vbCrLf
    Response.Write "     document.getElementById('VoteNumtips').style.display='None';" & vbCrLf	
    Response.Write "  }" & vbCrLf	
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf		
End Sub

Sub Addtr(TempNum)
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "  var count="&TempNum&";" & vbCrLf	
    Response.Write "  function AddRow(){" & vbCrLf
    Response.Write "   if (count>20){" & vbCrLf
    Response.Write "   alert('最多只能添加20项')" & vbCrLf	
    Response.Write "   }" & vbCrLf
    Response.Write "   else{" & vbCrLf			
    Response.Write "   var tr = document.createElement('tr');" & vbCrLf
    Response.Write "   var td1 = document.createElement('td');" & vbCrLf
    Response.Write "   var td2 = document.createElement('td');" & vbCrLf	
    Response.Write "   var td3 = document.createElement('td');" & vbCrLf	
    Response.Write "   var td4 = document.createElement('td');" & vbCrLf			
    Response.Write "   var input1 = document.createElement('input');" & vbCrLf	
    Response.Write "   var input2 = document.createElement('input');" & vbCrLf			
    Response.Write "   var Items = document.createTextNode('选项'+count+'：')" & vbCrLf
    Response.Write "   var Nums = document.createTextNode('票数：');" & vbCrLf
    Response.Write "   tr.setAttribute('class','tdbg');" & vbCrLf	
    Response.Write "   td1.setAttribute('width','20%');" & vbCrLf	
    Response.Write "   td1.setAttribute('align','right');" & vbCrLf		
    Response.Write "   td2.setAttribute('width','35%');" & vbCrLf			
    Response.Write "   td3.setAttribute('width','10%');" & vbCrLf	
    Response.Write "   td3.setAttribute('align','right');" & vbCrLf					
    Response.Write "   td4.setAttribute('width','80');" & vbCrLf			
    Response.Write "   input1.setAttribute('type','text');" & vbCrLf
    Response.Write "   input1.setAttribute('name','select'+count);" & vbCrLf
    Response.Write "   input1.setAttribute('size','36');" & vbCrLf		
    Response.Write "   input2.setAttribute('type','text');" & vbCrLf
    Response.Write "   input2.setAttribute('name','answer'+count);" & vbCrLf
    Response.Write "   input2.setAttribute('size','10');" & vbCrLf		
    Response.Write "   td1.appendChild(Items);" & vbCrLf	
    Response.Write "   td2.appendChild(input1);" & vbCrLf			
    Response.Write "   td3.appendChild(Nums);" & vbCrLf		
    Response.Write "   td4.appendChild(input2);" & vbCrLf				
    Response.Write "   tr.appendChild(td1);" & vbCrLf
    Response.Write "   tr.appendChild(td2);" & vbCrLf
    Response.Write "   tr.appendChild(td3);" & vbCrLf	
    Response.Write "   tr.appendChild(td4);" & vbCrLf	
    Response.Write "   var AddRow = document.getElementById('AddRow');" & vbCrLf
    Response.Write "   AddRow.appendChild(tr);" & vbCrLf
    Response.Write "   count++;" & vbCrLf	
    Response.Write "  }" & vbCrLf
    Response.Write "  }" & vbCrLf	
    Response.Write " </script>" & vbCrLf
End Sub

Sub Add()
    Dim i
    Call ShowJS_AddModify
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Vote.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' class='title' colspan=4 align=center><b>添 加 调 查</b></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>所属频道：</td>"
    Response.Write "      <td colspan='3'>"
    Response.Write "        <select name='ChannelID' id='ChannelID'>" & GetChannel_Option(0) & "</select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>调查主题：</td>"
    Response.Write "      <td colspan='3'>"
    Response.Write "        <textarea name='Title' cols='60' rows='4'></textarea>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    For i = 1 To 8
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td width='20%' align='right'>选项" & i & "：</td>"
        Response.Write "      <td width='35%'>"
        Response.Write "        <input type='text' name='select" & i & "' size='36'>"
        Response.Write "      </td>"
        Response.Write "      <td width='10%' align='right'>票数：</td>"
        Response.Write "      <td width='35%' width='80'>"
        Response.Write "        <input type='text' name='answer" & i & "' size='10'>"
        Response.Write "      </td>"
        Response.Write "    </tr>"

    Next
    Call Addtr(i)		
    Response.Write "    <tbody id='AddRow' class='tdbg'>"
    Response.Write "</tbody>"	
    Response.Write "    <tr class='tdbg'><td align='right'><a onclick='AddRow()' style=""cursor:hand;color:#f00;"">>>继续添加</a></td><td colspan='5'></td></tr>"		
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>调查类型：</td>"
    Response.Write "      <td colspan='5'>"
    Response.Write "        <select onchange=""SetVoteNum(this.options[this.selectedIndex].value)""  name='VoteType' id='VoteType'>"
    Response.Write "          <option value='Single' selected>单选</option>"
    Response.Write "          <option value='Multi'>多选</option>"
    Response.Write "        </select>"
    Response.Write "     <input name='VoteNum'  style='display:none' type='text' value='0' onclick=""this.value=''"" title='设置可选票数,0为不限制' id='VoteNum' size='2' maxlength='2'><font id='VoteNumtips'  color='#FF0000' style='display:none'> （最大可选票数，0为不限制）</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>发布时间：</td>"
    Response.Write "      <td colspan='3'>"
    Response.Write "        <input name='VoteTime' type='text' id='VoteTime' value='" & Now() & "' size='20' maxlength='20'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>终止时间：</td>"
    Response.Write "      <td colspan='3'>"
    Response.Write "        <input name='EndTime' type='text' id='EndTime' value='" & Now() + 30 & "' size='20' maxlength='20'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>&nbsp;</td>"
    Response.Write "      <td colspan='3'>"
    Response.Write "        <input name='IsSelected' type='checkbox' id='IsSelected' value='yes' checked>"
    Response.Write "        启用本调查</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td colspan=4 align=center>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 添 加 '>"
    Response.Write "        &nbsp;"
    Response.Write "        <input  name='Reset' type='reset' id='Reset' value=' 清 除 '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
End Sub

Sub Modify()
    Dim rs, sql
    Dim i
    Dim TempNum
    TempNum=1	
    If ID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的调查ID！</li>"
        Exit Sub
    Else
        ID = PE_CLng(ID)
    End If
    sql = "select * from PE_Vote where ID=" & ID
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到指定的调查！</li>"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    Call ShowJS_AddModify

    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='Admin_Vote.asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' class='title' colspan=4 align=center><b>修 改 调 查</b></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>所属频道：</td>"
    Response.Write "      <td colspan='3'>"
    Response.Write "        <select name='ChannelID' id='ChannelID'>" & GetChannel_Option(rs("ChannelID")) & "</select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>调查主题：</td>"
    Response.Write "      <td colspan='3'>"
    Response.Write "        <textarea name='Title' cols='60' rows='4'>" & PE_ConvertBR(rs("Title")) & "</textarea>"
    Response.Write "      </td>"
    Response.Write "    </tr>"

    For i = 1 To 20
        If rs("select" & i) <>"" then
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td width='20%' align='right'>选项" & i & "：</td>"
            Response.Write "      <td width='35%' >"
            Response.Write "        <input type='text' name='select" & i & "' value='" & rs("select" & i) & "' size='36'>"
            Response.Write "      </td>"
            Response.Write "      <td width='10%' align='right'>票数：</td>"
            Response.Write "      <td width='35%'>"
            Response.Write "        <input type='text' name='answer" & i & "' value='" & rs("answer" & i) & "' size='10'>"
            Response.Write "      </td>"
            Response.Write "    </tr>"
            TempNum = TempNum + 1			
        End IF	
    Next
    Call Addtr(TempNum)		
    Response.Write "    <tbody id='AddRow' class='tdbg'>"
    Response.Write "</tbody>"	
    Response.Write "    <tr class='tdbg'><td align='right' cols=5><a onclick='AddRow()' style=""cursor:hand;color:#f00;"">>>继续添加</a></td><td colspan='5'></td></tr>"		
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>调查类型：</td>"
    Response.Write "      <td colspan='3'>"
    Response.Write "        <select name='VoteType' onchange=""SetVoteNum(this.options[this.selectedIndex].value)""   id='VoteType'>"
    Response.Write "          <option value='Single' "
    If rs("VoteType") = "Single" Then Response.Write " selected"
    Response.Write "          >单选</option>"
    Response.Write "          <option value='Multi' "
    If rs("VoteType") = "Multi" Then Response.Write " selected"
    Response.Write "          >多选</option>"
    Response.Write "        </select>"
    Response.Write "     <input name='VoteNum' "
    If rs("VoteType") = "Single" Then Response.Write " style='display:none'"
    Response.Write "  type='text' value='"& rs("VoteNum")&"' onclick=""this.value=''"" title='设置可选票数,0为不限制' id='VoteNum' size='2' maxlength='2'><font id='VoteNumtips'  color='#FF0000' "
    If rs("VoteType") = "Single" Then Response.Write "  style='display:none'"
    Response.Write "	  > （最大可选票数，0为不限制）</font></td>"
    Response.Write "    </tr>"	
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>发布时间：</td>"
    Response.Write "      <td colspan='3'>"
    Response.Write "        <input name='VoteTime' type='text' id='VoteTime' value='" & rs("VoteTime") & "' size='20' maxlength='20'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>终止时间：</td>"
    Response.Write "      <td colspan='3'>"
    Response.Write "        <input name='EndTime' type='text' id='EndTime' value='" & rs("EndTime") & "' size='20' maxlength='20'>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td align='right'>&nbsp;</td>"
    Response.Write "      <td colspan='3'>"
    Response.Write "        <input name='IsSelected' type='checkbox' id='IsSelected' value='yes' "
    If rs("IsSelected") = True Then Response.Write " checked"
    Response.Write "        >启用本调查</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan=4 align=center>"
    Response.Write "        <input name='ID' type='hidden' id='ID' value='" & ID & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "        <input name='Submit' type='submit' id='Submit' value=' 保 存 '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
    rs.Close
    Set rs = Nothing
End Sub

Sub SaveVote()
    Dim Title, VoteTime, EndTime, VoteType, IsSelected, VoteNum
    Dim rs, sql
    Dim i
    ChannelID = PE_CLng(Request("ChannelID"))
    Title = Trim(Request("Title"))
    VoteTime = PE_CDate(Trim(Request("VoteTime")))
    EndTime = Trim(Request("EndTime"))
    VoteType = Trim(Request("VoteType"))
    IsSelected = Trim(Request("IsSelected"))
    VoteNum = PE_Clng(Trim(Request("VoteNum")))	
    If Title = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>调查主题不能为空！</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If

    Title = PE_HTMLEncode(Title)
    If IsSelected = "yes" Then
        IsSelected = True
        '只有一个有效的调查
        'Conn.Execute("update PE_Vote set IsSelected=False Where ChannelID=" & ChannelID)
    Else
        IsSelected = False
    End If

    Set rs = Server.CreateObject("adodb.recordset")
    If Action = "SaveAdd" Then
        sql = "select top 1 * from PE_Vote"
        rs.Open sql, Conn, 1, 3
        rs.addnew
    ElseIf Action = "SaveModify" Then
        If ID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>不能确定调查ID</li>"
            Exit Sub
        Else
            sql = "select * from PE_Vote where ID=" & PE_CLng(ID)
            Set rs = Server.CreateObject("adodb.recordset")
            rs.Open sql, Conn, 1, 3
            If rs.BOF And rs.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>找不到指定的调查！</li>"
                rs.Close
                Set rs = Nothing
                Exit Sub
            End If
        End If
    End If

    rs("ChannelID") = ChannelID
    rs("Title") = Title
    Dim j
    j=1
    For i = 1 To 20
        If Trim(Request("select" & i))<>"" Then 
            rs("select" & j) = Trim(Request("select" & i))
            If Request("answer" & i) = "" Then
                rs("answer" & j) = 0
            Else
                rs("answer" & j) = PE_CLng(Request("answer" & i))
            End If
            j=j+1
        End If	
    Next
    For i = j To 20
        rs("select" & i) = ""	
        rs("answer" & i) = 0	
    Next
	
    rs("VoteTime") = VoteTime
    If EndTime <> "" And IsDate(EndTime) Then
        rs("EndTime") = EndTime
    End If
    rs("VoteType") = VoteType
    rs("VoteNum") = VoteNum	
    rs("IsSelected") = IsSelected
    rs("IsItem") = False
    rs.Update
    rs.Close
    Set rs = Nothing
    
    If IsSelected = "yes" Then
        PE_Cache.DelCache (ChannelID & "_Site_Vote")
    End If
    Call CreateJS_Vote
    Call CloseConn
    Response.Redirect "admin_Vote.asp?ChannelID=" & ChannelID
End Sub

Sub SetProperty()
    Dim sqlProperty, rsProperty
    Dim MoveChannelID
    If ID = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定调查ID</li>"
    End If
    If Action = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>参数不足！</li>"
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    If InStr(ID, ",") > 0 Then
        sqlProperty = "select * from PE_Vote where ID in (" & ID & ")"
    Else
        sqlProperty = "select * from PE_Vote where ID=" & ID
    End If
    Set rsProperty = Server.CreateObject("ADODB.Recordset")
    rsProperty.Open sqlProperty, Conn, 1, 3
    Do While Not rsProperty.EOF
        Select Case Action
        Case "SetNew"
            rsProperty("IsSelected") = True
            PE_Cache.DelCache (rsProperty("ChannelID") & "_Site_Vote")
        Case "CancelNew"
            rsProperty("IsSelected") = False
            PE_Cache.DelCache (rsProperty("ChannelID") & "_Site_Vote")
        Case "Move"
            MoveChannelID = PE_CLng(Trim(Request("ChannelID")))
            PE_Cache.DelCache (rsProperty("ChannelID") & "_Site_Vote")
            PE_Cache.DelCache (MoveChannelID & "_Site_Vote")
            rsProperty("ChannelID") = MoveChannelID
        Case "Del"
            PE_Cache.DelCache (rsProperty("ChannelID") & "_Site_Vote")
            rsProperty.Delete
        End Select
        rsProperty.Update
        rsProperty.MoveNext
    Loop
    rsProperty.Close
    Set rsProperty = Nothing
    Call CreateJS_Vote
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Sub CreateJS_Vote()
    If ObjInstalled_FSO = False Then
        Exit Sub
    End If
    Dim sqlVote, rsVote, i, strVote
    sqlVote = "select * from PE_Vote where IsSelected=" & PE_True & " and (ChannelID=-1 or ChannelID=" & ChannelID & ") and IsItem=" & PE_False & " order by ID Desc"
    Set rsVote = Conn.Execute(sqlVote)
    If rsVote.BOF And rsVote.EOF Then
        strVote = "&nbsp;没有任何调查"
    Else
        Do While Not rsVote.EOF
            strVote = strVote & "<form name='VoteForm' method='post' action='" & InstallDir & "vote.asp' target='_blank'>"
            strVote = strVote & "&nbsp;&nbsp;&nbsp;&nbsp;" & rsVote("Title") & "<br>"
            If rsVote("VoteType") = "Single" Then
                For i = 1 To 8
                    If Trim(rsVote("Select" & i) & "") = "" Then Exit For
                    strVote = strVote & "<input type='radio' name='VoteOption' value='" & i & "' style='border:0'>" & rsVote("Select" & i) & "<br>"
                Next
            Else
                For i = 1 To 20
                    If Trim(rsVote("Select" & i) & "") = "" Then Exit For
                    strVote = strVote & "<input type='checkbox' name='VoteOption' value='" & i & "' style='border:0'>" & rsVote("Select" & i) & "<br>"
                Next
            End If
            strVote = strVote & "<br><input name='VoteType' type='hidden'value='" & rsVote("VoteType") & "'>"
            strVote = strVote & "<input name='Action' type='hidden' value='Vote'>"
            strVote = strVote & "<input name='ID' type='hidden' value='" & rsVote("ID") & "'>"
            strVote = strVote & "<div align='center'>"
            strVote = strVote & "<a href='javascript:VoteForm.submit();'><img src='" & InstallDir & "images/voteSubmit.gif' width='52' height='18' border='0'></a>&nbsp;&nbsp;"
            strVote = strVote & "<a href='" & InstallDir & "Vote.asp?ID=" & rsVote("ID") & "&Action=Show' target='_blank'><img src='" & InstallDir & "images/voteView.gif' width='52' height='18' border='0'></a>"
            strVote = strVote & "</div></form>"
            rsVote.MoveNext
        Loop
    End If
    rsVote.Close
    Set rsVote = Nothing

    Dim JSPath
    If ChannelDir <> "" Then
        JSPath = InstallDir & ChannelDir & "/js"
    Else
        JSPath = InstallDir & "js"
    End If
    If Not fso.FolderExists(Server.MapPath(JSPath)) Then
        fso.CreateFolder (Server.MapPath(JSPath))
    End If
    Call WriteToFile(JSPath & "/ShowVote.js", "document.write(""" & strVote & """);")
End Sub

%>
