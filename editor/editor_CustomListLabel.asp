<!-- #include File="../Start.asp" -->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim Title, ModuleType, ChannelShortName, ChannelShowType, imageproperty
Dim editLabel, Labletemp
Dim ClassID, NClassID, IncludeChild, SpecialID, Num, ProductType, IsHot, IsElite, IsPic, AuthorName, DateNum
Dim OrderType, ShowType, TitleLen, ContentLen, ShowClassName, ShowPropertyType, ShowIncludePic, ShowAuthor
Dim ShowDateType, ShowHits, ShowHotSign, ShowNewSign, ShowTips, ShowCommentLink, UsePage, OpenType, Cols
Dim ImgWidth, ImgHeight, iTimeOut, urltype, CssNameA, CssName1, CssName2, effectID
Dim Template
Dim ChannelID, iChannelID, dChannelID

ChannelID = Trim(request("ChannelID"))
dChannelID = Trim(request("dChannelID"))

TemplateID = PE_CLng(request("TemplateID"))

NClassID = False

If dChannelID = "" Then
   dChannelID = ChannelID
End If
If ChannelID = "" And iChannelID = "" Then
    Response.write "频道参数丢失！"
    Response.End
End If

If ChannelID = "ChannelID" Then
    iChannelID = Trim(dChannelID)
Else
    ChannelID = PE_CLng(ChannelID)
    iChannelID = ChannelID
End If

Action = Trim(request("Action"))
Title = Trim(request("Title"))
ModuleType = PE_CLng(Trim(request("ModuleType")))
ChannelShowType = Trim(request("ChannelShowType"))
   
If ModuleType = 1 Then
    ChannelShortName = "文章"
    imageproperty = "Article"
ElseIf ModuleType = 2 Then
    ChannelShortName = "下载"
    imageproperty = "Soft"
ElseIf ModuleType = 3 Then
    ChannelShortName = "图片"
    imageproperty = "Photo"
ElseIf ModuleType = 5 Then
    iChannelID = 1000
    ChannelShortName = "商品"
    imageproperty = "Product"
End If

If SpecialID = "" Then SpecialID = 0
If Trim(request.querystring("editLabel")) <> "" Then
    editLabel = True
End If

Response.write "<html>" & vbCrLf
Response.write "<head>" & vbCrLf
Response.write "<title>" & Title & "</title>" & vbCrLf
Response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.write "<script language=""javascript"">" & vbCrLf
Response.write "function NClassIDChild(){" & vbCrLf
Response.write "    if (document.myform.NClassChild.checked==true){" & vbCrLf
Response.write "        document.myform.ClassID.size=2;" & vbCrLf
Response.write "        document.myform.ClassID.style.height=250;" & vbCrLf
Response.write "        document.myform.ClassID.style.width=400;" & vbCrLf
Response.write "        document.myform.ClassID.multiple=true" & vbCrLf
Response.write "        for(var i=0;i<document.myform.ClassID.length;i++){" & vbCrLf
Response.write "            if (document.myform.ClassID.options[i].value==""rsClass_arrChildID""||document.myform.ClassID.options[i].value==""ClassID""||document.myform.ClassID.options[i].value==""arrChildID""||document.myform.ClassID.options[i].value==""0""){" & vbCrLf
Response.write "                document.myform.ClassID.options[i].style.background=""red"";" & vbCrLf
Response.write "            }" & vbCrLf
Response.write "        }" & vbCrLf
Response.write "    }else{" & vbCrLf
Response.write "        document.myform.ClassID.size=1;" & vbCrLf
Response.write "        document.myform.ClassID.style.width=200;" & vbCrLf
Response.write "        document.myform.ClassID.multiple=false;" & vbCrLf
Response.write "        for(var i=0;i<document.myform.ClassID.length;i++){" & vbCrLf
Response.write "            if (document.myform.ClassID.options[i].value==""rsClass_arrChildID""||document.myform.ClassID.options[i].value==""ClassID""||document.myform.ClassID.options[i].value==""arrChildID""||document.myform.ClassID.options[i].value==""0""){" & vbCrLf
Response.write "                document.myform.ClassID.options[i].style.background="""";" & vbCrLf
Response.write "            }" & vbCrLf
Response.write "        }" & vbCrLf
Response.write "    }" & vbCrLf
Response.write "}" & vbCrLf
Response.write "function objectTag() {" & vbCrLf
Response.write "    var strJS,OrderType;" & vbCrLf
Response.write "    if (document.myform.ClassID.value==''){" & vbCrLf
Response.write "        alert('所属栏目不能指定为外部栏目！');" & vbCrLf
Response.write "        document.myform.ClassID.focus();" & vbCrLf
Response.write "        return false;" & vbCrLf
Response.write "    }" & vbCrLf
Response.write "    var UsePage,ShowAll;" & vbCrLf
Response.write "    for (var i=0;i<document.myform.UsePage.length;i++){" & vbCrLf
Response.write "    var PowerEasy = document.myform.UsePage[i];" & vbCrLf
Response.write "    if (PowerEasy.checked==true)       " & vbCrLf
Response.write "        UsePage = PowerEasy.value" & vbCrLf
Response.write "    }" & vbCrLf
'If ModuleType = 2 Then
'    Response.write "    for (var i=0;i<document.myform.ShowAll.length;i++){" & vbCrLf
'    Response.write "    var PowerEasy = document.myform.ShowAll[i];" & vbCrLf
'    Response.write "    if (PowerEasy.checked==true)       " & vbCrLf
'    Response.write "        ShowAll = PowerEasy.value" & vbCrLf
'    Response.write "    }" & vbCrLf
'End If
Response.write "    strJS=""【" & imageproperty & "List("";" & vbCrLf
If ModuleType = 1 Or ModuleType = 2 Or ModuleType = 3 Then
    Response.write "strJS+=document.myform.iChannelID.value+ "","";" & vbCrLf
'ElseIf ModuleType = 2 Then
'    Response.write "strJS+=document.myform.ItemNum.value+ "","";" & vbCrLf
End If
Response.write "    if (document.myform.NClassChild.checked==true){" & vbCrLf
Response.write "        var Nclassidzhi=""""" & vbCrLf
Response.write "        for(var i=0;i<document.myform.ClassID.length;i++){" & vbCrLf
Response.write "            if (document.myform.ClassID.options[i].selected==true){" & vbCrLf
Response.write "                if (document.myform.ClassID.options[i].value==""rsClass_arrChildID""||document.myform.ClassID.options[i].value==""ClassID""||document.myform.ClassID.options[i].value==""arrChildID""||document.myform.ClassID.options[i].value==""0""){" & vbCrLf
Response.write "                    alert(""您在多选中选择了红色部分，多选栏目中是不能包含那部分的。"");" & vbCrLf
Response.write "                    return false" & vbCrLf
Response.write "                }else{" & vbCrLf
Response.write "                    if (Nclassidzhi==""""){" & vbCrLf
Response.write "                        Nclassidzhi+=document.myform.ClassID.options[i].value;" & vbCrLf
Response.write "                    }else{" & vbCrLf
Response.write "                        Nclassidzhi+=""|""+document.myform.ClassID.options[i].value;" & vbCrLf
Response.write "                    }" & vbCrLf
Response.write "                }" & vbCrLf
Response.write "            }" & vbCrLf
Response.write "        }" & vbCrLf
Response.write "        strJS+=Nclassidzhi;" & vbCrLf
Response.write "    }else{" & vbCrLf
Response.write "        strJS+=document.myform.ClassID.value;" & vbCrLf
Response.write "    }" & vbCrLf
If ModuleType = 1 Or ModuleType = 2 Or ModuleType = 3 Or ModuleType = 5 Then
    Response.write "strJS+="",""+document.myform.IncludeChild.checked;" & vbCrLf
    Response.write "strJS+="",""+document.myform.SpecialID.value;" & vbCrLf
    Response.write "strJS+="",""+document.myform.ItemNum.value;" & vbCrLf
'ElseIf ModuleType = 2 Then
'    Response.write "strJS+="",""+ShowAll;" & vbCrLf
End If
If ModuleType = 5 Then
    Response.write "strJS+="",""+document.myform.ProductType.value;" & vbCrLf
End If
Response.write "strJS+="",""+document.myform.IsHot.checked;" & vbCrLf
Response.write "strJS+="",""+document.myform.IsElite.checked;" & vbCrLf
If ModuleType <> 5 Then
    Response.write "strJS+="",""+document.myform.AuthorName.value;" & vbCrLf
End If
Response.write "strJS+="",""+document.myform.DateNum.value;" & vbCrLf
Response.write "strJS+="",""+document.myform.OrderType.value;" & vbCrLf
Response.write "strJS+="",""+UsePage;" & vbCrLf
'If ModuleType = 2 Then
'    Response.write "strJS+="",""+document.myform.OpenType.value;" & vbCrLf
'End If
If ModuleType = 1 Or ModuleType = 2 Or ModuleType = 3 Or ModuleType = 5 Then
    Response.write "strJS+="",""+document.myform.TitleLen.value;" & vbCrLf
End If
Response.write "strJS+="",""+document.myform.ContentLen.value;" & vbCrLf
If ModuleType <> 3 Then
    Response.write "strJS+="",""+document.myform.IsPic.checked;" & vbCrLf
End If
Response.write "strJS+="")】"";" & vbCrLf
Response.write " if (document.myform.Cols.value!=0){" & vbCrLf
Response.write "    strJS+=""【Cols=""+document.myform.Cols.value+""|""+document.myform.ColsHtml.value+""】"";" & vbCrLf
Response.write "}" & vbCrLf
Response.write " if (document.myform.Rows.value!=0){" & vbCrLf
Response.write "    strJS+=""【Rows=""+document.myform.Rows.value+""|""+document.myform.RowsHtml.value+""】"";" & vbCrLf
Response.write "}" & vbCrLf
Response.write "strJS+=document.myform.Content.value;" & vbCrLf
Response.write "strJS+=""【/" & imageproperty & "List】"";" & vbCrLf
Response.write "window.returnValue = strJS;" & vbCrLf
Response.write "window.close();" & vbCrLf
Response.write "}" & vbCrLf
Response.write "function insertLabel(strLabel)" & vbCrLf
Response.write "{" & vbCrLf
Response.write "  myform.Content.focus();" & vbCrLf
Response.write "  var str = document.selection.createRange();" & vbCrLf
Response.write "  str.text = strLabel" & vbCrLf
Response.write "}" & vbCrLf
Response.write " function previewContent() {" & vbCrLf
Response.write "    var Content=document.myform.Content.value" & vbCrLf
Response.write "    Content = Content.replace(""&"", ""{$ID}"");" & vbCrLf
Response.write "    window.showModalDialog(""editor_previewContent.asp?Content=""+Content+""&ChannelID=" & ChannelID & ",toolbar=no, menubar=no, top=0,left=0,dialogwidth=800,dialogheight=600,help: no; scroll:yes; status: yes"");" & vbCrLf
Response.write " }" & vbCrLf
Response.write "</script>" & vbCrLf

Response.write "<link href='../Images/Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.write "<base target=""_self"">" & vbCrLf
Response.write "</head>" & vbCrLf
Response.write "<body>" & vbCrLf
Response.write "<form action='editor_CustomListLabel.asp' method='post' name='myform' id='myform'>" & vbCrLf
Response.write "<table width='700' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>  " & vbCrLf
Response.write "    <tr class=title>" & vbCrLf
Response.write "      <td align=middle colSpan=2 height=22><STRONG>" & ChannelShortName & "自定义列表标签 </STRONG></td>" & vbCrLf
Response.write "    </tr>" & vbCrLf

If ModuleType <> 5 Then
    Response.write "    <tr class='tdbg'>"
    Response.write "      <td height='25' width='130'  class='tdbg5' align='right'><strong>所属频道：</strong></td>" & vbCrLf
    Response.write "      <td height='25' class='tdbg5'><input type='hidden' name='iChannelID' value='" & ChannelID & "'><select name='ChannelID' onChange='document.myform.submit();'>" & GetChannel_Option(ModuleType, ChannelID) & "</select></td>"
    Response.write "    </tr>"
End If
If PE_CLng(iChannelID) > 0 Then
    Response.write "    <tr class='tdbg'>"
    Response.write "      <td height='25'  class='tdbg5' width='130' align='right'><strong>所属栏目：</strong></td>" & vbCrLf
    Response.write "      <td height='25' ><select name='ClassID' "
    If NClassID = True Then
        Response.write "size='2' multiple style='height:250px;width:400px;'"
    Else
        Response.write "size='1'"
    End If
    Response.write ">" & GetClass_Channel(iChannelID, Trim(ClassID), NClassID) & "</select>"
    Response.write " <input type='checkbox' name='IncludeChild' value='1' "
    If LCase(Trim(IncludeChild)) = "true" Then
    Response.write " checked "
    End If
    Response.write " >包含子栏目&nbsp;&nbsp;<font color='red'><b>注意：</b></font>不能指定为外部栏目 </font>"
    Response.write "  <br><input type='checkbox' name='NClassChild' value='1' onClick=""javascript:NClassIDChild()"" "
    If NClassID = True Then
    Response.write " checked "
    End If
    Response.write " >是否选择多个栏目&nbsp;&nbsp;<font color='red'><b>注意：</b></font>红色的栏目不能选 </font>"
    Response.write "      </td>"
    Response.write "    </tr>"
    'If ModuleType <> 2 Then
        Response.write "    <tr class='tdbg'>"
        Response.write "      <td height='25'  class='tdbg5' width='130' align='right'><strong>所属专题：</strong></td>"
        Response.write "      <td height='25' ><select name='SpecialID' id='SpecialID'>" & GetSpecial_Option(SpecialID) & "</select></td>"
        Response.write "    </tr>"
    'End If
Else
    Response.write "<INPUT TYPE='hidden' name='ClassID' value='0' >"
    Response.write "<INPUT TYPE='hidden' name='NClassChild' value='0' >"
    Response.write "<INPUT TYPE='hidden' name='IncludeChild' value='true' >"
    Response.write "<INPUT TYPE='hidden' name='SpecialID' value='0' >"
End If
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width='130' class='tdbg5' align='right' height=25><STRONG>" & ChannelShortName & "数：</STRONG></td>" & vbCrLf
Response.write "      <td height=25><input type='text' name='ItemNum' value='0' size=""8"">&nbsp;&nbsp;&nbsp;<font color='#FF0000'>如果为0，将显示所有" & ChannelShortName & "。</font></td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
'If ModuleType = 2 Then
'    Response.write "    <tr class=tdbg>" & vbCrLf
'    Response.write "      <td width=""130"" class='tdbg5' align=right height=25><STRONG>是否显示所有" & ChannelShortName & "：</STRONG></td>" & vbCrLf
'    Response.write "      <td height=25 >" & vbCrLf
'    Response.write "      <input type=""radio"" name=""ShowAll"" checked value=""True"">是" & vbCrLf
'    Response.write "      <input type=""radio"" name=""ShowAll"" value=""False"">否</td>" & vbCrLf
'    Response.write "    </tr>" & vbCrLf
'End If
If ModuleType = 5 Then
    Response.write "    <tr class='tdbg'>"
    Response.write "      <td height='25' class='tdbg5' align='right'><strong> 产品类型：</strong></td>"
    Response.write "      <td height='25' ><select name='ProductType' id='ProductType'>"
    Response.write "        <option value='1'"
    If Trim(ProductType) = "1" Then Response.write "selected"
    Response.write ">正常销售商品</option>"
    Response.write "        <option value='2'"
    If Trim(ProductType) = "2" Then Response.write "selected"
    Response.write ">涨价商品</option>"
    Response.write "        <option value='3'"
    If Trim(ProductType) = "3" Then Response.write "selected"
    Response.write ">特价商品</option>"
    Response.write "        <option value='4'"
    If Trim(ProductType) = "4" Then Response.write "selected"
    Response.write ">促销礼品</option>"
    Response.write "        <option value='5'"
    If Trim(ProductType) = "5" Then Response.write "selected"
    Response.write ">正常销售和涨价商品</option>"
    Response.write "        <option value='6'"
    If Trim(ProductType) = "6" Then Response.write "selected"
    Response.write ">降价商品</option>"
    Response.write "        <option value='0'"
    If Trim(ProductType) = "0" Then Response.write "selected"
    Response.write ">所有商品</option>"
    Response.write "        </select> </td>"
    Response.write "    </tr>"
End If
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130"" class='tdbg5' align=right height=25><STRONG>" & ChannelShortName & "属性：</STRONG></td>" & vbCrLf
Response.write "      <td height=25 >" & vbCrLf
Response.write "        <Input id=IsHot type=checkbox value=1 name=IsHot> 热门" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
If ModuleType <> 3 Then
    Response.write "        <Input id=IsPic type=checkbox value=1 name=IsPic> 图片" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
End If
Response.write "        <Input id=IsElite type=checkbox value=1 name=IsElite> 推荐" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<FONT color=#ff0000>如果都不选，将显示所有文章。</FONT></td>" & vbCrLf
Response.write "    </tr>" & vbCrLf

If ModuleType <> 5 Then
    Response.write "    <tr class=tdbg>" & vbCrLf
    Response.write "      <td width=""130"" class='tdbg5' align=right height=25><STRONG>显示指定作者的" & ChannelShortName & "：</STRONG></td>" & vbCrLf
    Response.write "      <td height=25 > " & vbCrLf
    Response.write "         <Input id=AuthorName  maxLength=10 size=10 value='' name=AuthorName>&nbsp;&nbsp;&nbsp;&nbsp;<FONT color=#ff0000>如果都不添，将不指定。</FONT>" & vbCrLf
    Response.write "    </tr>" & vbCrLf
End If
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130"" class='tdbg5' align=right height=25><STRONG>日期范围：</STRONG></td>" & vbCrLf
Response.write "      <td height=25>只显示最近 " & vbCrLf
Response.write "        <Input id=DateNum maxLength=3 size=5 value=0 name=DateNum> 天内更新的" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<FONT color=#ff0000>&nbsp;&nbsp;如果为空或0，则显示所有天数的文章。</FONT></td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130"" class='tdbg5' align=right height=25><STRONG>排序方法：</STRONG></td>" & vbCrLf
Response.write "      <td height=25 >" & vbCrLf
Response.write "        <Select id='OrderType' name='OrderType'> " & vbCrLf
Response.write "          <Option value=1 selected>" & ChannelShortName & "ID（降序）</Option> " & vbCrLf
Response.write "          <Option value=2>" & ChannelShortName & "ID（升序）</Option> " & vbCrLf
Response.write "          <Option value=3>更新时间（降序）</Option> " & vbCrLf
Response.write "          <Option value=4>更新时间（升序）</Option> " & vbCrLf
Response.write "          <Option value=5>点击次数（降序）</Option> " & vbCrLf
Response.write "          <Option value=6>点击次数（升序）</Option>" & vbCrLf
Response.write "        </Select>" & vbCrLf
Response.write "      </td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130""  class='tdbg5' align=right height=25><STRONG>是否分页显示：</STRONG></td>" & vbCrLf
Response.write "      <td height=25 >" & vbCrLf
Response.write "        <input type=""radio"" name=""UsePage"" checked value=""True"">是" & vbCrLf
Response.write "        <input type=""radio"" name=""UsePage"" value=""False"">否</td> " & vbCrLf
Response.write "    </tr>" & vbCrLf
'If ModuleType = 2 Then
'    Response.write "    <tr class=tdbg>" & vbCrLf
'    Response.write "      <td width=""130""  class='tdbg5' align=right height=25><STRONG>" & ChannelShortName & "打开方式：</STRONG></td>" & vbCrLf
'    Response.write "      <td height=25 >" & vbCrLf
'    Response.write "        <Select id=OpenType name=OpenType> " & vbCrLf
'    Response.write "          <Option value=0 selected>在原窗口打开</Option> " & vbCrLf
'    Response.write "          <Option value=1>在新窗口打开</Option>" & vbCrLf
'    Response.write "        </Select></td>" & vbCrLf
'    Response.write "    </tr>       " & vbCrLf
'End If
'If ModuleType <> 2 Then
    Response.write "    <tr class=tdbg>" & vbCrLf
    Response.write "      <td width=""130""  class='tdbg5' align=right height=25><STRONG>" & ChannelShortName & "标题最多字符：</STRONG></td>" & vbCrLf
    Response.write "      <td height=25 >" & vbCrLf
    Response.write "        <Input  maxLength=3 size=5 value=0 name=TitleLen> &nbsp;&nbsp;&nbsp;&nbsp;<FONT color=#ff0000>一个汉字=两个英文字符，为0时不显示</FONT></td>" & vbCrLf
    Response.write "    </tr>" & vbCrLf
'End If
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130""  class='tdbg5' align=right height=25><STRONG>" & ChannelShortName & "内容最多字符：</STRONG></td>" & vbCrLf
Response.write "      <td height=25>" & vbCrLf
Response.write "        <Input  maxLength=3 size=5 value=0 name=ContentLen> &nbsp;&nbsp;&nbsp;&nbsp;<FONT color=#ff0000>一个汉字=两个英文字符，为0时不显示</FONT></td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130""  class='tdbg5' align=right height=25><STRONG>" & ChannelShortName & "循环列设置：</STRONG></td>" & vbCrLf
Response.write "      <td height=25>" & vbCrLf
Response.write "        每显示<Input  maxLength=3 size=2 value=0 name=Cols>列后,向自定义循环列表中插入 <Input size=10 value="""" name=ColsHtml>&nbsp;&nbsp;<FONT color=#ff0000>支持Html代码</FONT></td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130""  class='tdbg5' align=right height=25><STRONG>" & ChannelShortName & "循环行设置：</STRONG></td>" & vbCrLf
Response.write "      <td height=25>" & vbCrLf
Response.write "        每显示<Input  maxLength=3 size=2 value=0 name=Rows>行后,向自定义循环列表中插入 <Input size=10 value="""" name=RowsHtml>&nbsp;&nbsp;<FONT color=#ff0000>支持Html代码</FONT></td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130"" class='tdbg5' align=right height=25><STRONG>循环标签支持标签：</STRONG></td>" & vbCrLf
Response.write "      <td>" & vbCrLf
Response.write "        <table border='0' cellpadding='0' cellspacing='0' width='100%' height='100%' align='center'>" & vbCrLf
If ModuleType = 1 Then
    Response.write "         <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ArticleUrl}')"" title=""文章的链接地址"">{$ArticleUrl}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ArticleID}')"" title=""文章的ID"">{$ArticleID}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$UpdateTime}')"" title=""文章更新时间"">{$UpdateTime}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Stars}')"" title=""文章评分等级"">{$Stars}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Author}')"" title=""文章作者"">{$Author}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$CopyFrom}')"" title=""文章来源"">{$CopyFrom}</a></td>" & vbCrLf
    Response.write "         </tr>" & vbCrLf
    Response.write "         <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Hits}')"" title=""点击次数"">{$Hits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Inputer}')"" title=""文章录入者"">{$Inputer}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Editor}')"" title=""责任编辑"">{$Editor}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ReadPoint}')"" title=""阅读点数"">{$ReadPoint}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Property}')"" title=""文章属性"">{$Property}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Top}')"" title=""固顶"">{$Top}</a></td>" & vbCrLf
    Response.write "         </tr>" & vbCrLf
    Response.write "         <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Elite}')"" title=""推荐"">{$Elite}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Hot}')"" title=""热门"">{$Hot}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Title}')"" title=""文章正标题，字数由参数TitleLen控制"">{$Title}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Subheading}')"" title=""自定义列表副标题"">{$Subheading}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Intro}')"" title=""文章简介"">{$Intro}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Content}')"" title=""文章正文内容，字数由参数ContentLen控制"">{$Content}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ArticlePic(130,90)}')"" title=""显示图片文章，width为图片宽度，height为图片高度"">{$ArticlePic(130,90)}</a></td>" & vbCrLf
    Response.write "           <td valign='top'></td>" & vbCrLf
    Response.write "           <td valign='top'></td>" & vbCrLf
    Response.write "           <td valign='top'></td>" & vbCrLf
    Response.write "           <td valign='top'></td>" & vbCrLf
    Response.write "           <td valign='top'></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
ElseIf ModuleType = 2 Then
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftUrl}')"" title=""软件地址"">{$SoftUrl}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftID}')"" title=""软件ID"">{$SoftID}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftName}')"" title=""软件名称"">{$SoftName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftVersion}')"" title=""软件版本"">{$SoftVersion}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftProperty}')"" title=""软件属性（固顶、推荐等）"">{$SoftProperty}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftSize}')"" title=""软件大小"">{$SoftSize}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$UpdateTime}')"" title=""更新时间"">{$UpdateTime}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$CopyrightType}')"" title=""版权类型"">{$CopyrightType}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Stars}')"" title=""评分等级"">{$Stars}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftIntro}')"" title=""软件简介"">{$SoftIntro}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$OperatingSystem}')"" title=""运行平台"">{$OperatingSystem}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftType}')"" title=""软件类型"">{$SoftType}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Hits}')"" title=""点击数"">{$Hits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$DayHits}')"" title=""显示每日点击数"">{$DayHits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$WeekHits}')"" title=""显示每周点击数"">{$WeekHits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$MonthHits}')"" title=""显示每月点击数"">{$MonthHits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftPoint}')"" title=""显示下载软件时所需的点数"">{$SoftPoint}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftAuthor}')"" title=""软件作者"">{$SoftAuthor}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftLanguage}')"" title=""语言种类"">{$SoftLanguage}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$AuthorEmail}')"" title=""作者Email"">{$AuthorEmail}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$DemoUrl}')"" title=""软件演示地址"">{$DemoUrl}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$RegUrl}')"" title=""软件注册地址"">{$RegUrl}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftPic(130,90)}')"" title=""显示图片软件，width为图片宽度，height为图片高度"">{$SoftPic(130,90)}</a></td>" & vbCrLf
    Response.write "           <td valign='top'></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
ElseIf ModuleType = 3 Then
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PhotoID}')"" title=""图片ID"">{$PhotoID}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PhotoUrl}')"" title=""图片地址"">{$PhotoUrl}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$UpdateTime}')"" title=""更新时间"">{$UpdateTime}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Stars}')"" title=""评分等级"">{$Stars}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Author}')"" title=""图片作者"">{$Author}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$CopyFrom}')"" title=""图片来源"">{$CopyFrom}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Hits}')"" title=""点击数"">{$Hits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Inputer}')"" title=""图片的录入作者"">{$Inputer}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Editor}')"" title=""图片的编辑者"">{$Editor}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$InfoPoint}')"" title=""查看点数"">{$InfoPoint}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Keyword}')"" title=""关键字"">{$Keyword}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Keyword}')"" title=""关键字"">{$Keyword}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Property}')"" title=""图片属性（固顶，热门，推荐等）"">{$Property}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Top}')"" title=""显示固顶"">{$Top}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Elite}')"" title=""显示推荐"">{$Elite}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Hot}')"" title=""显示热门"">{$Hot}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PhotoName}')"" title=""显示图片名称"">{$PhotoName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PhotoName}')"" title=""显示图片名称"">{$PhotoName}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PhotoIntro}')"" title=""显示图片介绍"">{$PhotoIntro}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PhotoThumb}')"" title=""显示图片的缩略图"">{$PhotoThumb}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$DayHits}')"" title=""显示当日点击数"">{$DayHits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$WeekHits}')"" title=""显示本周点击数"">{$WeekHits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$MonthHits}}')"" title=""显示本月点击数"">{$MonthHits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
ElseIf ModuleType = 5 Then
    Response.write "         <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ClassUrl}')"" title=""栏目链接"">{$ClassUrl}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ClassID}')"" title=""栏目ID"">{$ClassID}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ClassName}')"" title=""栏目名称"">{$ClassName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ParentDir}')"" title=""父目录"">{$ParentDir}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ClassDir}')"" title=""栏目目录"">{$ClassDir}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductUrl}')"" title=""商品链接"">{$ProductUrl}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductID}')"" title=""商品ID"">{$ProductID}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductNum}')"" title=""商品数"">{$ProductNum}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductModel}')"" title=""商品型号"">{$ProductModel}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductStandard}')"" title=""商品规格"">{$ProductStandard}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Top}')"" title=""固顶"">{$Top}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Elite}')"" title=""推荐"">{$Elite}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Hot}')"" title=""热门"">{$Hot}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$UpdateTime}')"" title=""更新时间"">{$UpdateTime}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Stars}')"" title=""评分等级"">{$Stars}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductName}')"" title=""商品名称"">{$ProductName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductTypeName}')"" title=""商品类别"">{$ProductTypeName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductIntro}')"" title=""商品简介"">{$ProductIntro}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductExplain}')"" title=""显示商品说明"">{$ProductExplain}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductThumb(130,90)}')"" title=""显示商品图片，width为图片宽度，height为图片高度"">{$ProductThumb(130,90)}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$BeginDate}')"" title=""显示优惠开始日期"">{$BeginDate}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$EndDate}')"" title=""显示优惠结束日期"">{$EndDate}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Discount}')"" title=""降价折扣"">{$Discount}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$LimitNum}')"" title=""限够数量"">{$LimitNum}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Price_Original}')"" title=""原始零售价"">{$Price_Original}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Price_Market}')"" title=""显示市场价"">{$Price_Market}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Price_Member}')"" title=""显示会员价"">{$Price_Member}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Price}')"" title=""显示商城价"">{$Price}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProducerName}')"" title=""生 产 商"">{$ProducerName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$TrademarkName}')"" title=""品牌商标"">{$TrademarkName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PresentExp}')"" title=""购物积分"">{$PresentExp}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PresentPoint}')"" title=""赠送的点数"">{$PresentPoint}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PresentMoney}')"" title=""返还的现金券"">{$PresentMoney}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PointName}')"" title=""点券的名称"">{$PointName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PointUnit}')"" title=""点券的单位"">{$PointUnit}</a></td>" & vbCrLf
    Response.write "           <td valign='top'></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf

End If
Response.write "          <tr>" & vbCrLf
Set rs = Conn.Execute("select * from PE_Field where ChannelID=" & ChannelID & "")
If rs.bof And rs.EOF Then
Else
    Do While Not rs.EOF
        Response.write "    <td valign='top'><a href=""javascript:insertLabel('" & rs("LabelName") & "')"" title=""" & rs("title") & """>" & rs("LabelName") & "</a></td>" & vbCrLf
        rs.movenext
    Loop
End If
Set rs = Nothing
Response.write "          </tr>" & vbCrLf
Response.write "        </table>" & vbCrLf
Response.write "    </td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
Response.write "    <tr class='tdbg'>" & vbCrLf
Response.write "        <td width=""130""  class='tdbg5' align=right height=25><STRONG>请输入循环标签Html代码：</STRONG>" & vbCrLf
Response.write "        <br>" & vbCrLf

If TemplateID <> 0 Then
    sql = "select * from PE_Template where TemplateID=" & TemplateID
    Set rs = Conn.Execute(sql)
    If rs.bof And rs.EOF Then
        Template = ""
    Else
        Dim StrBody, arrContent
        Template = rs("TemplateContent")
        If InStr(Template, "<body") > 0 Then
            regEx.Pattern = "(\<body)(.[^\<]*)(\>)"
            Set Matches = regEx.Execute(Template)
            For Each Match In Matches
                StrBody = Match.Value
            Next
        Else
            StrBody = "<body>"
        End If
        arrContent = Split(Template, StrBody)
        If UBound(arrContent) <> 0 Then
            Template = Replace(Replace(arrContent(1), "</body>", ""), "</html>", "")
        End If
    End If
    rs.Close
    Set rs = Nothing
End If

Dim rs, sql
sql = "select * from PE_Template where ChannelID=0 and TemplateType=101"
Set rs = Conn.Execute(sql)
Response.write "<select name='TemplateID' onChange='document.myform.submit();'>" & vbCrLf
If rs.bof And rs.EOF Then
    Response.write "<option value=""0"" selected>还没有列表模板！</option> " & vbCrLf
Else
    Response.write "<option value=""0"" selected>选择调用的模板样式</option>" & vbCrLf
    Do While Not rs.EOF
        Response.write "<option value=" & rs("TemplateID") & ">" & rs("TemplateName") & "</option>" & vbCrLf
        rs.movenext
    Loop
End If
Response.write "</select>" & vbCrLf
rs.Close
Set rs = Nothing

Response.write "        </td>" & vbCrLf

Response.write "        <td>" & vbCrLf
Response.write "        <textarea NAME='Content'  style='width:550px;height:200px'>" & Template & "</TEXTAREA>     " & vbCrLf
'Response.Write "        <textarea name='Content' id='Content' style='display:none' ></textarea>"
'Response.Write "        <iframe ID='editor' src='../editor.asp?ChannelID=1&ShowType=2&tContentid=Content' frameborder='1' scrolling='no' width='550' height='280' ></iframe>"
Response.write "        </td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
Response.write "    <tr class='tdbg'>" & vbCrLf
Response.write "      <td height='10' colspan='2' align='center'>" & vbCrLf
Response.write "        <input name='Title' type='hidden' id='Title' value='" & Title & "'>" & vbCrLf
Response.write "        <input name='Action' type='hidden' id='Action' value='" & Action & "'>" & vbCrLf
Response.write "        <input name='editLabel' type='hidden' id='editLabel' value='" & editLabel & "'>" & vbCrLf
Response.write "        <input name='dChannelID' type='hidden' id='dChannelID' value='" & dChannelID & "'> " & vbCrLf
Response.write "        <input name='ModuleType' type='hidden' id='ModuleType' value='" & ModuleType & "'>" & vbCrLf
Response.write "        <input name='ChannelShowType' type='hidden' id='ChannelShowType' value='"" & ChannelShowType & ""'> " & vbCrLf
Response.write "      </td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
Response.write "    <tr class='tdbg'>" & vbCrLf
Response.write "        <td colspan=2 align='center'>" & vbCrLf
Response.write "          <input TYPE='button' value=' 确 定 ' onCLICK='objectTag()'>&nbsp;&nbsp;" & vbCrLf
Response.write "          <input name='EditorpreviewContent' type='button' id='EditorpreviewContent' value=' 预 览 ' onclick='previewContent();'>" & vbCrLf
Response.write "        </td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
Response.write "  </table>" & vbCrLf
Response.write "</FORM>" & vbCrLf
Response.write "</body>" & vbCrLf
Response.write "</html>" & vbCrLf

Function GetSpecial_Option(SpecialID)
    Dim sqlSpecial, rsSpecial, strOption, strOptionTemp
    sqlSpecial = "select ChannelID,SpecialID,SpecialName,OrderID from PE_Special where ChannelID=0 or ChannelID=" & ChannelID & "   order by ChannelID,OrderID"
    Set rsSpecial = Conn.Execute(sqlSpecial)
    If LCase(SpecialID) <> "specialid" Then
        If PE_CLng(SpecialID) = 0 Then
            strOption = "<option value='0'>不属于任何专题</option>"
        Else
            strOption = "<option value='0' selected>不属于任何专题</option>"
        End If
    End If
    If rsSpecial.bof And rsSpecial.bof Then
    Else
        Do While Not rsSpecial.EOF
            If rsSpecial("ChannelID") > 0 Then
                strOptionTemp = rsSpecial("SpecialName") & "（本频道）"
            Else
                strOptionTemp = rsSpecial("SpecialName") & "（全站）"
            End If
            If rsSpecial("SpecialID") = PE_CLng(SpecialID) Then
                strOption = strOption & "<option value='" & rsSpecial("SpecialID") & "' selected>" & strOptionTemp & "</option>"
            Else
                strOption = strOption & "<option value='" & rsSpecial("SpecialID") & "'>" & strOptionTemp & "</option>"
            End If
            rsSpecial.movenext
        Loop
    End If
    rsSpecial.Close
    Set rsSpecial = Nothing
    strOption = strOption & "<option value='SpecialID'"
    If SpecialID = "SpecialID" Then strOption = strOption & " selected"
    strOption = strOption & ">当前频道</option>"

    GetSpecial_Option = strOption
End Function

Function GetChannel_Option(iModuleType, ChannelID)
    Dim strChannel, sqlChannel, rsChannel
    sqlChannel = "select ChannelID,ChannelName from PE_Channel  where ModuleType=" & iModuleType & " and Disabled=" & PE_False & " and ChannelType<=1 order by OrderID"
    Set rsChannel = Conn.Execute(sqlChannel)
    Do While Not rsChannel.EOF
        If rsChannel(0) = PE_CLng(ChannelID) Then
            strChannel = strChannel & "<option value='" & rsChannel(0) & "' selected>" & rsChannel(1) & "</option>"
        Else
            strChannel = strChannel & "<option value='" & rsChannel(0) & "'>" & rsChannel(1) & "</option>"
        End If
        rsChannel.movenext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    strChannel = strChannel & "<option value='0'"
    If ChannelID = "0" Then strChannel = strChannel & " selected"
    strChannel = strChannel & ">所有同类频道</option>"
    strChannel = strChannel & "<option value='ChannelID'"
    If ChannelID = "ChannelID" Then strChannel = strChannel & " selected"
    strChannel = strChannel & ">当前频道</option>"
    GetChannel_Option = strChannel
End Function

Function GetClass_Channel(ChannelID, ClassID, NClassID)
    Dim rsClass, sqlClass, strClass_Option, tmpDepth, i, classcss
    Dim arrShowLine(20)
    For i = 0 To UBound(arrShowLine)
    arrShowLine(i) = False
    Next
    sqlClass = "Select * from PE_Class where ChannelID=" & ChannelID & " order by RootID,OrderID"
    Set rsClass = Conn.Execute(sqlClass)
    If rsClass.bof And rsClass.bof Then
    strClass_Option = strClass_Option & "<option value='0'>请先添加栏目</option>"
    Else
        Do While Not rsClass.EOF
            tmpDepth = rsClass("Depth")
            If rsClass("NextID") > 0 Then
                arrShowLine(tmpDepth) = True
            Else
                arrShowLine(tmpDepth) = False
            End If

            If rsClass("ClassType") = 2 Then
                strClass_Option = strClass_Option & "<option value=''"
            Else
                strClass_Option = strClass_Option & "<option value='" & rsClass("ClassID") & "'"
                If NClassID = False Then
                    If ClassID <> "rsClass_arrChildID" Or ClassID <> "ClassID" Or ClassID <> "arrChildID" Then
                        If rsClass("ClassID") = PE_CLng(ClassID) Then
                            strClass_Option = strClass_Option & " selected"
                        End If
                    End If
                Else
                    If FoundInArr(ClassID, rsClass("ClassID"), "|") = True Then
                        strClass_Option = strClass_Option & " selected"
                    End If
                End If
            End If
            strClass_Option = strClass_Option & ">"
            
            If tmpDepth > 0 Then
            For i = 1 To tmpDepth
                strClass_Option = strClass_Option & "&nbsp;&nbsp;"
                If i = tmpDepth Then
                If rsClass("NextID") > 0 Then
                    strClass_Option = strClass_Option & "├&nbsp;"
                Else
                    strClass_Option = strClass_Option & "└&nbsp;"
                End If
                Else
                If arrShowLine(i) = True Then
                    strClass_Option = strClass_Option & "│"
                Else
                    strClass_Option = strClass_Option & "&nbsp;"
                End If
                End If
            Next
            End If
            strClass_Option = strClass_Option & rsClass("ClassName")
            If rsClass("ClassType") = 2 Then
                strClass_Option = strClass_Option & "（外）"
            End If
            strClass_Option = strClass_Option & "</option>"
            rsClass.movenext
        Loop
    End If
    rsClass.Close
    Set rsClass = Nothing
    If NClassID = False Then
        classcss = "style=''"
    Else
        classcss = "style='background:red'"
    End If
    
    If Trim(ClassID) = "rsClass_arrChildID" Then
        strClass_Option = strClass_Option & "<option value='rsClass_arrChildID' " & classcss & " selected >栏目循环中的栏目</option>"
    Else
        strClass_Option = strClass_Option & "<option value='rsClass_arrChildID' " & classcss & " >栏目循环中的栏目</option>"
    End If
    If Trim(ClassID) = "ClassID" Then
        strClass_Option = strClass_Option & "<option value='ClassID' " & classcss & " selected>当前栏目（不包含子栏目）</option>"
    Else
        strClass_Option = strClass_Option & "<option value='ClassID' " & classcss & ">当前栏目（不包含子栏目）</option>"
    End If
    If Trim(ClassID) = "arrChildID" Then
        strClass_Option = strClass_Option & "<option value='arrChildID' " & classcss & " selected>当前栏目及子栏目</option>"
    Else
        strClass_Option = strClass_Option & "<option value='arrChildID' " & classcss & ">当前栏目及子栏目</option>"
    End If
    If Trim(ClassID) = "0" Then
        strClass_Option = strClass_Option & "<option value='0' " & classcss & " selected>显示所有栏目</option>"
    Else
        strClass_Option = strClass_Option & "<option value='0' " & classcss & ">显示所有栏目</option>"
    End If

    GetClass_Channel = strClass_Option
End Function
%>

