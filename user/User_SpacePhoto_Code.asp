<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim BlogID
BlogID = PE_CLng(Trim(Request("ID")))

Sub Execute()
    BlogID = PE_CLng(Trim(Request("ID")))
    If BlogID = 0 Then
        Exit Sub
    End If

    FileName = "User_SpacePhoto.asp?ID=" & BlogID & "&Action=" & Action
    strFileName = FileName & "&Field=" & strField & "&keyword=" & Keyword

    Response.Write "<table align='center'><tr align='center' valign='top'>"
    Response.Write "<td width='90'><a href='User_SpacePhoto.asp?ID=" & BlogID & "&Action=Add'><img src='images/photo_add.gif' border='0' align='absmiddle'><br>添加我的图片</a></td>"
    Response.Write "<td width='90'><a href='User_SpacePhoto.asp?ID=" & BlogID & "&Action=Add1'><img src='images/photo_up.gif' border='0' align='absmiddle'><br>上传我的图片</a></td>"
    Response.Write "<td width='90'><a href='User_SpacePhoto.asp?ID=" & BlogID & "&Action=AddFlash'><img src='images/photo_cam.gif' border='0' align='absmiddle'><br>拍摄我的图片</a></td>"
    Response.Write "<td width='90'><a href='User_SpacePhoto.asp?ID=" & BlogID & "'><img src='images/photo_all.gif' border='0' align='absmiddle'><br>管理我的图片</a></td>"
    Response.Write "</tr></table>" & vbCrLf
    
    Select Case Action
    Case "Add"
        Call Add(0)
    Case "Add1"
        Call Add(1)
    Case "AddFlash"
        Call Add(2)
    Case "Modify"
        Call Modify
    Case "SaveAdd", "SaveModify"
        Call SaveItem
    Case "Del"
        Call Del
    Case Else
        Call main
    End Select
    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
    End If
End Sub

Sub main()
    If FoundErr = True Then Exit Sub
    Call ShowJS_Main("图片")

    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td height='22'>图片管理</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr>"
    Response.Write "    <form name='myform' method='Post' action='User_SpacePhoto.asp' onsubmit='return ConfirmDel();'>"
    Response.Write "     <td><table class='border' border='0' cellspacing='1' width='100%' cellpadding='0'>"
    Response.Write "          <tr class='title' height='22'> "
    Response.Write "            <td height='22' width='30' align='center'><strong>选中</strong></td>"
    Response.Write "            <td width='25' align='center'><strong>ID</strong></td>"
    Response.Write "            <td align='center' ><strong>标题</strong></td>"
    Response.Write "            <td width='40' align='center' ><strong>点击数</strong></td>"
    Response.Write "            <td width='130' align='center' ><strong>日期</strong></td>"
    Response.Write "            <td width='80' align='center' ><strong>管理操作</strong></td>"
    Response.Write "    </tr>"

    Dim rsItem, sql
    sql = "select * from PE_SpacePhoto Where UserID=" & UserID
    If Keyword <> "" Then
        Select Case strField
        Case "Title"
            sql = sql & " and Title like '%" & Keyword & "%' "
        Case "Content"
            sql = sql & " and Content like '%" & Keyword & "%' "
        Case "Time"
            sql = sql & " and Datetime='" & Keyword & "' "
        Case Else
            sql = sql & " and Title like '%" & Keyword & "%' "
        End Select
    End If
    sql = sql & " order by ID desc"

    Set rsItem = Server.CreateObject("ADODB.Recordset")
    rsItem.Open sql, Conn, 1, 1
    If rsItem.BOF And rsItem.EOF Then
        totalPut = 0
        Response.Write "<tr class='tdbg'><td colspan='20' align='center'><br>您尚未添加图片<br><br></td></tr>"
    Else
        totalPut = rsItem.RecordCount
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
                rsItem.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If

        Dim itemNum: itemNum = 0
        Do While Not rsItem.EOF
            Response.Write "<tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td align='center'><input name='ItemID' type='checkbox' onclick='unselectall()' id='ItemID' value='" & rsItem("ID") & "'></td>"
            Response.Write "    <td align='center'>" & rsItem("ID") & "</td>"
            Response.Write "    <td><a href='../Space/Showphoto.asp?ID=" & rsItem("ID") & "' target='_blank'>" & rsItem("title") & "</a></td>"
            Response.Write "    <td align='center'>" & rsItem("Hits") & "</td>"
            Response.Write "    <td align='center'>" & rsItem("Datetime") & "</td>"
            Response.Write "    <td align='center'>"
            Response.Write "<a href='User_SpacePhoto.asp?Action=Modify&ID=" & BlogID & "&ItemID=" & rsItem("ID") & "'>修改</a>&nbsp;"
            Response.Write "<a href='User_SpacePhoto.asp?Action=Del&ID=" & BlogID & "&ItemID=" & rsItem("ID") & "' onclick=""return confirm('确定要删除此图片吗？一旦删除将不能恢复！');"">删除</a>"
            Response.Write "</td></tr>"
            itemNum = itemNum + 1
            If itemNum >= MaxPerPage Then Exit Do
            rsItem.MoveNext
        Loop
    End If
    rsItem.Close
    Set rsItem = Nothing
    Response.Write "</table>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>选中本页显示的所有图片</td><td>"
    Response.Write "<input name='submit' type='submit' value='删除选定的图片' onClick=""document.myform.Action.value='Del'"">"
    Response.Write "<input name='Action' type='hidden' id='Action' value=''>"
    Response.Write "  </td></tr>"
    Response.Write "</table>"
    Response.Write "</td>"
    Response.Write "</form></tr></table>"
    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "篇图片", True)
    End If

    Response.Write "<br>"
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "   <td width='80' align='right'><strong>图片搜索：</strong></td>"
    Response.Write "   <td>" & GetSearchForm(FileName) & "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub

Sub ShowJS_Item(iType)
    Response.Write "<script language = 'JavaScript'>" & vbCrLf
    Response.Write "function CheckForm(){" & vbCrLf
    Response.Write "  if (document.myform.url.value==''){" & vbCrLf
    If iType = 0 Then
        Response.Write "document.myform.url.focus();" & vbCrLf
        Response.Write "    alert('您尚未添加图片！');" & vbCrLf
    ElseIf iType = 1 Then
        Response.Write "    alert('您尚未上传图片！');" & vbCrLf
    ElseIf iType = 2 Then
        Response.Write "    alert('您尚未拍摄图片！');" & vbCrLf
    End If
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  if (document.myform.Title.value==''){" & vbCrLf
    Response.Write "    alert('图片名称不能为空！');" & vbCrLf
    Response.Write "    document.myform.Title.focus();" & vbCrLf
    Response.Write "    return false;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "  return true;  " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub Add(iType)
    Call ShowJS_Item(iType)
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='User_SpacePhoto.asp' target='_self'>"
    Response.Write "<table border=0 cellpadding=0 cellspacing=0 align=center width='95%'>"
    If iType = 1 Then
        Response.Write "<tr><td><fieldset><legend>预览图片</legend>"
        Response.Write "  <table border='0' cellpadding='0' cellspacing='5' width='100%'>"
        Response.Write "    <tr height='200'><td align='center'>"
        Response.Write "        <img id='img' src='../images/nopic.gif' border='0' width='460' height='350'><Input name='url' type='hidden' id='url' size='30'>"
        Response.Write "    </td></tr>"
        Response.Write "  </table></fieldset></td></tr>"
    ElseIf iType = 2 Then
        Response.Write "<tr><td><fieldset><legend>拍摄照片</legend>"
        Response.Write "  <table border='0' cellpadding='0' cellspacing='5' width='100%'>"
        Response.Write "    <tr height='200'><td align='center'>"
        Response.Write "<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" width=""377"" height=""187"" id=""myFlash"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab"">"
        Response.Write "<param name=""movie"" value=""FlashCam.swf""><param name=""quality"" value=""high""><param name=""wmode"" value=""transparent""><param name=""bgcolor"" value=""#ffffff"">"
        Response.Write "<embed src=""FlashCam.swf"" quality=""high"" wmode=""transparent"" bgcolor=""#ffffff"" width=""377"" height=""187"" name=""bmap"" align=""middle"""
        Response.Write " play=""true"" loop=""false"" quality=""high"" allowScriptAccess=" 'sameDomain"" NAME=""myFlash"" swLiveConnect=""true"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"">"
        Response.Write "</embed></object>"
        Response.Write "  <Input name='url' type='hidden' id='url' size='30'></td></tr>"
        Response.Write "  </table></fieldset></td></tr>"
    Else
        Response.Write "    <tr><td align='center'>图片地址：<Input name='url' type='Text' id='url' size='60'></td></tr>"
    End If
    If iType = 1 Then
        Response.Write "<tr><td align='center'><fieldset align='center'>"
        Response.Write "<legend align=left>上传图片</legend>"
        Response.Write "  <iframe class='TBGen' style='top:2px' id='UploadFiles' src='upload.asp?DialogType=userblogpic&size=" & UserSetting(27) & "' frameborder='0' scrolling='no' width='350' height='32'></iframe>"
        Response.Write "</fieldset></td></tr>"
    End If
    Response.Write "    <tr><td align='center'>图片说明：<Input name='Title' type='Text' id='Title' size='60'></td></tr>"
    Response.Write "</table>"
    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='SaveAdd'>"
    Response.Write "   <input name='ID' type='hidden' id='ID' value='" & BlogID & "'>"
    Response.Write "   <input name='add' type='submit'  id='Add' value=' 确 认 ' onClick=""document.myform.Action.value='SaveAdd';document.myform.target='_self';"" style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='User_SpacePhoto.asp?ID=" & BlogID & "';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
    If iType = 2 Then
        Response.Write "<script language = 'JavaScript'>" & vbCrLf
        Response.Write "document.myform.add.disabled = true;" & vbCrLf
        Response.Write "function myFlash_DoFSCommand(command, args) {" & vbCrLf
        Response.Write "　  document.myform.url.value = args;" & vbCrLf
        Response.Write "    document.myform.add.disabled = false;" & vbCrLf
        Response.Write "}" & vbCrLf
        Response.Write "document.write('<SCRIPT LANGUAGE=VBScript\> \n');" & vbCrLf
        Response.Write "document.write('on error resume next \n');" & vbCrLf
        Response.Write "document.write('Sub myFlash_FSCommand(ByVal command, ByVal args)\n');" & vbCrLf
        Response.Write "document.write(' call myFlash_DoFSCommand(command, args)\n');" & vbCrLf
        Response.Write "document.write('end sub\n');" & vbCrLf
        Response.Write "document.write('</SCRIPT\> \n');" & vbCrLf
        Response.Write "</script>" & vbCrLf
    End If
End Sub

Sub Modify()
    Dim rsItem, sql, ItemID
    ItemID = PE_CLng(Trim(Request("ItemID")))
    If ItemID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定要修改的图片</li>"
        Exit Sub
    End If

    sql = "select * from PE_SpacePhoto where ID=" & ItemID & " and UserID=" & UserID
    Set rsItem = Server.CreateObject("ADODB.Recordset")
    rsItem.Open sql, Conn, 1, 1
    If rsItem.BOF And rsItem.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到图片</li>"
        rsItem.Close
        Set rsItem = Nothing
        Exit Sub
    End If

    Call ShowJS_Item(1)
    Response.Write "<form method='POST' name='myform' onSubmit='return CheckForm();' action='User_SpacePhoto.asp' target='_self'>"
    Response.Write "<table border=0 cellpadding=0 cellspacing=0 align=center width='95%'>"
    Response.Write "<tr><td><fieldset><legend>预览图片</legend>"
    Response.Write "  <table border='0' cellpadding='0' cellspacing='5' width='100%'>"
    Response.Write "    <tr height='300'><td align='center'>"
    Response.Write "        <img id='img' src='" & rsItem("Content") & "' border='0' width='460' height='350'>"
    Response.Write "    </td></tr>"
    Response.Write "  </table></fieldset></td></tr>"
    Response.Write "<tr><td align='center'><fieldset align='center'>"
    Response.Write "<legend align=left>上传图片</legend>"
    Response.Write "  <iframe class='TBGen' style='top:2px' id='UploadFiles' src='upload.asp?DialogType=userblogpic&size=" & UserSetting(27) & "' frameborder='0' scrolling='no' width='350' height='32'></iframe>"
    Response.Write "</fieldset></td></tr>"
    Response.Write "    <tr><td align='center'>图片地址：<Input name='url' type='Text' id='url' size='60' value='" & rsItem("Content") & "'></td></tr>"
    Response.Write "    <tr><td align='center'>图片说明：<Input name='Title' type='Text' id='Title' size='60' value='" & rsItem("Title") & "'></td></tr>"
    Response.Write "</table>"
    Response.Write "  <p align='center'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='SaveModify'>"
    Response.Write "   <input name='ItemID' type='hidden' id='ID' value='" & rsItem("ID") & "'>"
    Response.Write "   <input name='ID' type='hidden' id='ID' value='" & BlogID & "'>"
    Response.Write "   <input name='Save' type='submit' value='保存修改结果' style='cursor:hand;'>&nbsp;"
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='User_SpacePhoto.asp?ID=" & BlogID & "';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
    rsItem.Close
    Set rsItem = Nothing
End Sub

Sub SaveItem()
    Dim rsItem, sql, i
    Dim ItemID, Title, Content

    ItemID = PE_CLng(Trim(Request.Form("ItemID")))
    Title = Trim(Request.Form("Title"))
    Content = Trim(Request.Form("url"))

    If Title = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>图片说明不能为空</li>"
    Else
        Title = ReplaceText(Title, 2)
    End If

    If Content = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>图片地址不能为空</li>"
    End If
    
    If FoundErr = True Then
        Exit Sub
    End If

    Title = PE_HTMLEncode(Title)
    Content = ReplaceBadUrl(Content)
    Set rsItem = Server.CreateObject("adodb.recordset")
    If Action = "SaveAdd" Then
        ItemID = PE_CLng(Conn.Execute("Select max(ID) from PE_SpacePhoto")(0)) + 1
        sql = "select top 1 * from PE_SpacePhoto"
        rsItem.Open sql, Conn, 1, 3
        rsItem.addnew
        rsItem("ID") = ItemID
        rsItem("UserID") = UserID
        rsItem("BlogID") = BlogID
        rsItem("Title") = Title
        rsItem("Content") = Content
        rsItem("Datetime") = Now()

        rsItem.Update
        Conn.Execute ("update PE_Space set LastUseTime=" & PE_Now & " where ID=" & BlogID & "")
        rsItem.Close
    ElseIf Action = "SaveModify" Then
        If ItemID = 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>不能确定图片的值</li>"
        Else
            sql = "select top 1 * from PE_SpacePhoto where ID=" & ItemID & " and UserID=" & UserID
            rsItem.Open sql, Conn, 1, 3
            If rsItem.BOF And rsItem.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>找不到此图片。</li>"
            Else
                rsItem("Title") = Title
                rsItem("Content") = Content
                rsItem("Datetime") = Now()
                rsItem.Update
            End If
            rsItem.Close
        End If
    End If
    Set rsItem = Nothing
    
    If FoundErr = True Then Exit Sub
    
    Response.Write "<br><br>"
    Response.Write "<table class='border' align=center width='400' border='0' cellpadding='0' cellspacing='0' bordercolor='#999999'>"
    Response.Write "  <tr align=center> "
    Response.Write "    <td  height='22' align='center' class='title'> "
    If Action = "SaveAdd" Then
        Response.Write "<b>添加图片成功</b>"
    Else
        Response.Write "<b>修改图片成功</b>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td><table width='100%' border='0' cellpadding='2' cellspacing='1'>"
    Response.Write "          <td width='100' align='right'><strong>图片名称：</strong></td>"
    Response.Write "          <td>" & Title & "</td>"
    Response.Write "        </tr>"
    Response.Write "        <tr class='tdbg'> "
    Response.Write "          <td width='100' align='right'><strong>发表日期：</strong></td>"
    Response.Write "          <td>" & Now() & "</td>"
    Response.Write "        </tr>"
    Response.Write "      </table></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='30' align='center'>"
    Response.Write "【<a href='User_SpacePhoto.asp?Action=Modify&ID=" & BlogID & "&ItemID=" & ItemID & "'>修改图片</a>】&nbsp;"
    Response.Write "【<a href='User_SpacePhoto.asp?Action=Add&ID=" & BlogID & "'>继续添加图片</a>】&nbsp;"
    Response.Write "【<a href='User_SpacePhoto.asp?ID=" & BlogID & "'>图片管理</a>】"
    Response.Write "【<a href='../Space/Showphoto.asp?ID=" & ItemID & "'>图片预览</a>】"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>" & vbCrLf
End Sub

Sub Del()
    Dim ItemID
    ItemID = Trim(Request("ItemID"))
    If IsValidID(ItemID) = False Then
        ItemID = ""
    End If
    If ItemID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请先选定图片！</li>"
        Exit Sub
    End If
    If InStr(ItemID, ",") > 0 Then
        Conn.Execute ("delete from PE_SpacePhoto Where ID in (" & ItemID & ") and UserID=" & UserID)
    Else
        Conn.Execute ("delete from PE_SpacePhoto Where ID=" & ItemID & " and UserID=" & UserID)
    End If
    Call CloseConn
    Response.Redirect ComeUrl
End Sub

Function GetSearchForm(Action)
    Dim strForm
    strForm = "<table border='0' cellpadding='0' cellspacing='0'>"
    strForm = strForm & "<form method='Get' name='SearchForm' action='" & Action & "'>"
    strForm = strForm & "<tr><td height='28' align='center'>"
    strForm = strForm & "<select name='Field' size='1'>"
    strForm = strForm & "<option value='Title' selected>图片名称</option>"
    strForm = strForm & "<option value='Content'>图片说明</option>"
    strForm = strForm & "</select>"
    strForm = strForm & "<input type='text' name='keyword'  size='20' value='关键字' maxlength='50' onFocus='this.select();'>"
    strForm = strForm & "<input type='submit' name='Submit'  value='搜索'>"
    strForm = strForm & "<input name='ID' type='hidden' id='ID' value='" & BlogID & "'>"
    strForm = strForm & "</td></tr></form></table>"
    GetSearchForm = strForm
End Function
%>
