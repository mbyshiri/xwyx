<!-- #include File="../Start.asp" -->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
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
    Response.write "Ƶ��������ʧ��"
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
    ChannelShortName = "����"
    imageproperty = "Article"
ElseIf ModuleType = 2 Then
    ChannelShortName = "����"
    imageproperty = "Soft"
ElseIf ModuleType = 3 Then
    ChannelShortName = "ͼƬ"
    imageproperty = "Photo"
ElseIf ModuleType = 5 Then
    iChannelID = 1000
    ChannelShortName = "��Ʒ"
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
Response.write "        alert('������Ŀ����ָ��Ϊ�ⲿ��Ŀ��');" & vbCrLf
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
Response.write "    strJS=""��" & imageproperty & "List("";" & vbCrLf
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
Response.write "                    alert(""���ڶ�ѡ��ѡ���˺�ɫ���֣���ѡ��Ŀ���ǲ��ܰ����ǲ��ֵġ�"");" & vbCrLf
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
Response.write "strJS+="")��"";" & vbCrLf
Response.write " if (document.myform.Cols.value!=0){" & vbCrLf
Response.write "    strJS+=""��Cols=""+document.myform.Cols.value+""|""+document.myform.ColsHtml.value+""��"";" & vbCrLf
Response.write "}" & vbCrLf
Response.write " if (document.myform.Rows.value!=0){" & vbCrLf
Response.write "    strJS+=""��Rows=""+document.myform.Rows.value+""|""+document.myform.RowsHtml.value+""��"";" & vbCrLf
Response.write "}" & vbCrLf
Response.write "strJS+=document.myform.Content.value;" & vbCrLf
Response.write "strJS+=""��/" & imageproperty & "List��"";" & vbCrLf
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
Response.write "      <td align=middle colSpan=2 height=22><STRONG>" & ChannelShortName & "�Զ����б��ǩ </STRONG></td>" & vbCrLf
Response.write "    </tr>" & vbCrLf

If ModuleType <> 5 Then
    Response.write "    <tr class='tdbg'>"
    Response.write "      <td height='25' width='130'  class='tdbg5' align='right'><strong>����Ƶ����</strong></td>" & vbCrLf
    Response.write "      <td height='25' class='tdbg5'><input type='hidden' name='iChannelID' value='" & ChannelID & "'><select name='ChannelID' onChange='document.myform.submit();'>" & GetChannel_Option(ModuleType, ChannelID) & "</select></td>"
    Response.write "    </tr>"
End If
If PE_CLng(iChannelID) > 0 Then
    Response.write "    <tr class='tdbg'>"
    Response.write "      <td height='25'  class='tdbg5' width='130' align='right'><strong>������Ŀ��</strong></td>" & vbCrLf
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
    Response.write " >��������Ŀ&nbsp;&nbsp;<font color='red'><b>ע�⣺</b></font>����ָ��Ϊ�ⲿ��Ŀ </font>"
    Response.write "  <br><input type='checkbox' name='NClassChild' value='1' onClick=""javascript:NClassIDChild()"" "
    If NClassID = True Then
    Response.write " checked "
    End If
    Response.write " >�Ƿ�ѡ������Ŀ&nbsp;&nbsp;<font color='red'><b>ע�⣺</b></font>��ɫ����Ŀ����ѡ </font>"
    Response.write "      </td>"
    Response.write "    </tr>"
    'If ModuleType <> 2 Then
        Response.write "    <tr class='tdbg'>"
        Response.write "      <td height='25'  class='tdbg5' width='130' align='right'><strong>����ר�⣺</strong></td>"
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
Response.write "      <td width='130' class='tdbg5' align='right' height=25><STRONG>" & ChannelShortName & "����</STRONG></td>" & vbCrLf
Response.write "      <td height=25><input type='text' name='ItemNum' value='0' size=""8"">&nbsp;&nbsp;&nbsp;<font color='#FF0000'>���Ϊ0������ʾ����" & ChannelShortName & "��</font></td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
'If ModuleType = 2 Then
'    Response.write "    <tr class=tdbg>" & vbCrLf
'    Response.write "      <td width=""130"" class='tdbg5' align=right height=25><STRONG>�Ƿ���ʾ����" & ChannelShortName & "��</STRONG></td>" & vbCrLf
'    Response.write "      <td height=25 >" & vbCrLf
'    Response.write "      <input type=""radio"" name=""ShowAll"" checked value=""True"">��" & vbCrLf
'    Response.write "      <input type=""radio"" name=""ShowAll"" value=""False"">��</td>" & vbCrLf
'    Response.write "    </tr>" & vbCrLf
'End If
If ModuleType = 5 Then
    Response.write "    <tr class='tdbg'>"
    Response.write "      <td height='25' class='tdbg5' align='right'><strong> ��Ʒ���ͣ�</strong></td>"
    Response.write "      <td height='25' ><select name='ProductType' id='ProductType'>"
    Response.write "        <option value='1'"
    If Trim(ProductType) = "1" Then Response.write "selected"
    Response.write ">����������Ʒ</option>"
    Response.write "        <option value='2'"
    If Trim(ProductType) = "2" Then Response.write "selected"
    Response.write ">�Ǽ���Ʒ</option>"
    Response.write "        <option value='3'"
    If Trim(ProductType) = "3" Then Response.write "selected"
    Response.write ">�ؼ���Ʒ</option>"
    Response.write "        <option value='4'"
    If Trim(ProductType) = "4" Then Response.write "selected"
    Response.write ">������Ʒ</option>"
    Response.write "        <option value='5'"
    If Trim(ProductType) = "5" Then Response.write "selected"
    Response.write ">�������ۺ��Ǽ���Ʒ</option>"
    Response.write "        <option value='6'"
    If Trim(ProductType) = "6" Then Response.write "selected"
    Response.write ">������Ʒ</option>"
    Response.write "        <option value='0'"
    If Trim(ProductType) = "0" Then Response.write "selected"
    Response.write ">������Ʒ</option>"
    Response.write "        </select> </td>"
    Response.write "    </tr>"
End If
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130"" class='tdbg5' align=right height=25><STRONG>" & ChannelShortName & "���ԣ�</STRONG></td>" & vbCrLf
Response.write "      <td height=25 >" & vbCrLf
Response.write "        <Input id=IsHot type=checkbox value=1 name=IsHot> ����" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
If ModuleType <> 3 Then
    Response.write "        <Input id=IsPic type=checkbox value=1 name=IsPic> ͼƬ" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
End If
Response.write "        <Input id=IsElite type=checkbox value=1 name=IsElite> �Ƽ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<FONT color=#ff0000>�������ѡ������ʾ�������¡�</FONT></td>" & vbCrLf
Response.write "    </tr>" & vbCrLf

If ModuleType <> 5 Then
    Response.write "    <tr class=tdbg>" & vbCrLf
    Response.write "      <td width=""130"" class='tdbg5' align=right height=25><STRONG>��ʾָ�����ߵ�" & ChannelShortName & "��</STRONG></td>" & vbCrLf
    Response.write "      <td height=25 > " & vbCrLf
    Response.write "         <Input id=AuthorName  maxLength=10 size=10 value='' name=AuthorName>&nbsp;&nbsp;&nbsp;&nbsp;<FONT color=#ff0000>�������������ָ����</FONT>" & vbCrLf
    Response.write "    </tr>" & vbCrLf
End If
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130"" class='tdbg5' align=right height=25><STRONG>���ڷ�Χ��</STRONG></td>" & vbCrLf
Response.write "      <td height=25>ֻ��ʾ��� " & vbCrLf
Response.write "        <Input id=DateNum maxLength=3 size=5 value=0 name=DateNum> ���ڸ��µ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<FONT color=#ff0000>&nbsp;&nbsp;���Ϊ�ջ�0������ʾ�������������¡�</FONT></td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130"" class='tdbg5' align=right height=25><STRONG>���򷽷���</STRONG></td>" & vbCrLf
Response.write "      <td height=25 >" & vbCrLf
Response.write "        <Select id='OrderType' name='OrderType'> " & vbCrLf
Response.write "          <Option value=1 selected>" & ChannelShortName & "ID������</Option> " & vbCrLf
Response.write "          <Option value=2>" & ChannelShortName & "ID������</Option> " & vbCrLf
Response.write "          <Option value=3>����ʱ�䣨����</Option> " & vbCrLf
Response.write "          <Option value=4>����ʱ�䣨����</Option> " & vbCrLf
Response.write "          <Option value=5>�������������</Option> " & vbCrLf
Response.write "          <Option value=6>�������������</Option>" & vbCrLf
Response.write "        </Select>" & vbCrLf
Response.write "      </td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130""  class='tdbg5' align=right height=25><STRONG>�Ƿ��ҳ��ʾ��</STRONG></td>" & vbCrLf
Response.write "      <td height=25 >" & vbCrLf
Response.write "        <input type=""radio"" name=""UsePage"" checked value=""True"">��" & vbCrLf
Response.write "        <input type=""radio"" name=""UsePage"" value=""False"">��</td> " & vbCrLf
Response.write "    </tr>" & vbCrLf
'If ModuleType = 2 Then
'    Response.write "    <tr class=tdbg>" & vbCrLf
'    Response.write "      <td width=""130""  class='tdbg5' align=right height=25><STRONG>" & ChannelShortName & "�򿪷�ʽ��</STRONG></td>" & vbCrLf
'    Response.write "      <td height=25 >" & vbCrLf
'    Response.write "        <Select id=OpenType name=OpenType> " & vbCrLf
'    Response.write "          <Option value=0 selected>��ԭ���ڴ�</Option> " & vbCrLf
'    Response.write "          <Option value=1>���´��ڴ�</Option>" & vbCrLf
'    Response.write "        </Select></td>" & vbCrLf
'    Response.write "    </tr>       " & vbCrLf
'End If
'If ModuleType <> 2 Then
    Response.write "    <tr class=tdbg>" & vbCrLf
    Response.write "      <td width=""130""  class='tdbg5' align=right height=25><STRONG>" & ChannelShortName & "��������ַ���</STRONG></td>" & vbCrLf
    Response.write "      <td height=25 >" & vbCrLf
    Response.write "        <Input  maxLength=3 size=5 value=0 name=TitleLen> &nbsp;&nbsp;&nbsp;&nbsp;<FONT color=#ff0000>һ������=����Ӣ���ַ���Ϊ0ʱ����ʾ</FONT></td>" & vbCrLf
    Response.write "    </tr>" & vbCrLf
'End If
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130""  class='tdbg5' align=right height=25><STRONG>" & ChannelShortName & "��������ַ���</STRONG></td>" & vbCrLf
Response.write "      <td height=25>" & vbCrLf
Response.write "        <Input  maxLength=3 size=5 value=0 name=ContentLen> &nbsp;&nbsp;&nbsp;&nbsp;<FONT color=#ff0000>һ������=����Ӣ���ַ���Ϊ0ʱ����ʾ</FONT></td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130""  class='tdbg5' align=right height=25><STRONG>" & ChannelShortName & "ѭ�������ã�</STRONG></td>" & vbCrLf
Response.write "      <td height=25>" & vbCrLf
Response.write "        ÿ��ʾ<Input  maxLength=3 size=2 value=0 name=Cols>�к�,���Զ���ѭ���б��в��� <Input size=10 value="""" name=ColsHtml>&nbsp;&nbsp;<FONT color=#ff0000>֧��Html����</FONT></td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130""  class='tdbg5' align=right height=25><STRONG>" & ChannelShortName & "ѭ�������ã�</STRONG></td>" & vbCrLf
Response.write "      <td height=25>" & vbCrLf
Response.write "        ÿ��ʾ<Input  maxLength=3 size=2 value=0 name=Rows>�к�,���Զ���ѭ���б��в��� <Input size=10 value="""" name=RowsHtml>&nbsp;&nbsp;<FONT color=#ff0000>֧��Html����</FONT></td>" & vbCrLf
Response.write "    </tr>" & vbCrLf
Response.write "    <tr class=tdbg>" & vbCrLf
Response.write "      <td width=""130"" class='tdbg5' align=right height=25><STRONG>ѭ����ǩ֧�ֱ�ǩ��</STRONG></td>" & vbCrLf
Response.write "      <td>" & vbCrLf
Response.write "        <table border='0' cellpadding='0' cellspacing='0' width='100%' height='100%' align='center'>" & vbCrLf
If ModuleType = 1 Then
    Response.write "         <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ArticleUrl}')"" title=""���µ����ӵ�ַ"">{$ArticleUrl}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ArticleID}')"" title=""���µ�ID"">{$ArticleID}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$UpdateTime}')"" title=""���¸���ʱ��"">{$UpdateTime}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Stars}')"" title=""�������ֵȼ�"">{$Stars}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Author}')"" title=""��������"">{$Author}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$CopyFrom}')"" title=""������Դ"">{$CopyFrom}</a></td>" & vbCrLf
    Response.write "         </tr>" & vbCrLf
    Response.write "         <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Hits}')"" title=""�������"">{$Hits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Inputer}')"" title=""����¼����"">{$Inputer}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Editor}')"" title=""���α༭"">{$Editor}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ReadPoint}')"" title=""�Ķ�����"">{$ReadPoint}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Property}')"" title=""��������"">{$Property}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Top}')"" title=""�̶�"">{$Top}</a></td>" & vbCrLf
    Response.write "         </tr>" & vbCrLf
    Response.write "         <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Elite}')"" title=""�Ƽ�"">{$Elite}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Hot}')"" title=""����"">{$Hot}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Title}')"" title=""���������⣬�����ɲ���TitleLen����"">{$Title}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Subheading}')"" title=""�Զ����б�����"">{$Subheading}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Intro}')"" title=""���¼��"">{$Intro}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Content}')"" title=""�����������ݣ������ɲ���ContentLen����"">{$Content}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ArticlePic(130,90)}')"" title=""��ʾͼƬ���£�widthΪͼƬ��ȣ�heightΪͼƬ�߶�"">{$ArticlePic(130,90)}</a></td>" & vbCrLf
    Response.write "           <td valign='top'></td>" & vbCrLf
    Response.write "           <td valign='top'></td>" & vbCrLf
    Response.write "           <td valign='top'></td>" & vbCrLf
    Response.write "           <td valign='top'></td>" & vbCrLf
    Response.write "           <td valign='top'></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
ElseIf ModuleType = 2 Then
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftUrl}')"" title=""�����ַ"">{$SoftUrl}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftID}')"" title=""���ID"">{$SoftID}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftName}')"" title=""�������"">{$SoftName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftVersion}')"" title=""����汾"">{$SoftVersion}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftProperty}')"" title=""������ԣ��̶����Ƽ��ȣ�"">{$SoftProperty}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftSize}')"" title=""�����С"">{$SoftSize}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$UpdateTime}')"" title=""����ʱ��"">{$UpdateTime}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$CopyrightType}')"" title=""��Ȩ����"">{$CopyrightType}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Stars}')"" title=""���ֵȼ�"">{$Stars}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftIntro}')"" title=""������"">{$SoftIntro}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$OperatingSystem}')"" title=""����ƽ̨"">{$OperatingSystem}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftType}')"" title=""�������"">{$SoftType}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Hits}')"" title=""�����"">{$Hits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$DayHits}')"" title=""��ʾÿ�յ����"">{$DayHits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$WeekHits}')"" title=""��ʾÿ�ܵ����"">{$WeekHits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$MonthHits}')"" title=""��ʾÿ�µ����"">{$MonthHits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftPoint}')"" title=""��ʾ�������ʱ����ĵ���"">{$SoftPoint}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftAuthor}')"" title=""�������"">{$SoftAuthor}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftLanguage}')"" title=""��������"">{$SoftLanguage}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$AuthorEmail}')"" title=""����Email"">{$AuthorEmail}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$DemoUrl}')"" title=""�����ʾ��ַ"">{$DemoUrl}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$RegUrl}')"" title=""���ע���ַ"">{$RegUrl}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$SoftPic(130,90)}')"" title=""��ʾͼƬ�����widthΪͼƬ��ȣ�heightΪͼƬ�߶�"">{$SoftPic(130,90)}</a></td>" & vbCrLf
    Response.write "           <td valign='top'></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
ElseIf ModuleType = 3 Then
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PhotoID}')"" title=""ͼƬID"">{$PhotoID}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PhotoUrl}')"" title=""ͼƬ��ַ"">{$PhotoUrl}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$UpdateTime}')"" title=""����ʱ��"">{$UpdateTime}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Stars}')"" title=""���ֵȼ�"">{$Stars}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Author}')"" title=""ͼƬ����"">{$Author}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$CopyFrom}')"" title=""ͼƬ��Դ"">{$CopyFrom}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Hits}')"" title=""�����"">{$Hits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Inputer}')"" title=""ͼƬ��¼������"">{$Inputer}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Editor}')"" title=""ͼƬ�ı༭��"">{$Editor}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$InfoPoint}')"" title=""�鿴����"">{$InfoPoint}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Keyword}')"" title=""�ؼ���"">{$Keyword}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Keyword}')"" title=""�ؼ���"">{$Keyword}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Property}')"" title=""ͼƬ���ԣ��̶������ţ��Ƽ��ȣ�"">{$Property}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Top}')"" title=""��ʾ�̶�"">{$Top}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Elite}')"" title=""��ʾ�Ƽ�"">{$Elite}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Hot}')"" title=""��ʾ����"">{$Hot}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PhotoName}')"" title=""��ʾͼƬ����"">{$PhotoName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PhotoName}')"" title=""��ʾͼƬ����"">{$PhotoName}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PhotoIntro}')"" title=""��ʾͼƬ����"">{$PhotoIntro}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PhotoThumb}')"" title=""��ʾͼƬ������ͼ"">{$PhotoThumb}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$DayHits}')"" title=""��ʾ���յ����"">{$DayHits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$WeekHits}')"" title=""��ʾ���ܵ����"">{$WeekHits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$MonthHits}}')"" title=""��ʾ���µ����"">{$MonthHits}</a></td>" & vbCrLf
    Response.write "           <td valign='top'></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
ElseIf ModuleType = 5 Then
    Response.write "         <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ClassUrl}')"" title=""��Ŀ����"">{$ClassUrl}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ClassID}')"" title=""��ĿID"">{$ClassID}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ClassName}')"" title=""��Ŀ����"">{$ClassName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ParentDir}')"" title=""��Ŀ¼"">{$ParentDir}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ClassDir}')"" title=""��ĿĿ¼"">{$ClassDir}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductUrl}')"" title=""��Ʒ����"">{$ProductUrl}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductID}')"" title=""��ƷID"">{$ProductID}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductNum}')"" title=""��Ʒ��"">{$ProductNum}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductModel}')"" title=""��Ʒ�ͺ�"">{$ProductModel}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductStandard}')"" title=""��Ʒ���"">{$ProductStandard}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Top}')"" title=""�̶�"">{$Top}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Elite}')"" title=""�Ƽ�"">{$Elite}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Hot}')"" title=""����"">{$Hot}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$UpdateTime}')"" title=""����ʱ��"">{$UpdateTime}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Stars}')"" title=""���ֵȼ�"">{$Stars}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductName}')"" title=""��Ʒ����"">{$ProductName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductTypeName}')"" title=""��Ʒ���"">{$ProductTypeName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductIntro}')"" title=""��Ʒ���"">{$ProductIntro}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductExplain}')"" title=""��ʾ��Ʒ˵��"">{$ProductExplain}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProductThumb(130,90)}')"" title=""��ʾ��ƷͼƬ��widthΪͼƬ��ȣ�heightΪͼƬ�߶�"">{$ProductThumb(130,90)}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$BeginDate}')"" title=""��ʾ�Żݿ�ʼ����"">{$BeginDate}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$EndDate}')"" title=""��ʾ�Żݽ�������"">{$EndDate}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Discount}')"" title=""�����ۿ�"">{$Discount}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$LimitNum}')"" title=""�޹�����"">{$LimitNum}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Price_Original}')"" title=""ԭʼ���ۼ�"">{$Price_Original}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Price_Market}')"" title=""��ʾ�г���"">{$Price_Market}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Price_Member}')"" title=""��ʾ��Ա��"">{$Price_Member}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$Price}')"" title=""��ʾ�̳Ǽ�"">{$Price}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$ProducerName}')"" title=""�� �� ��"">{$ProducerName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$TrademarkName}')"" title=""Ʒ���̱�"">{$TrademarkName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PresentExp}')"" title=""�������"">{$PresentExp}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PresentPoint}')"" title=""���͵ĵ���"">{$PresentPoint}</a></td>" & vbCrLf
    Response.write "          </tr>" & vbCrLf
    Response.write "          <tr>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PresentMoney}')"" title=""�������ֽ�ȯ"">{$PresentMoney}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PointName}')"" title=""��ȯ������"">{$PointName}</a></td>" & vbCrLf
    Response.write "           <td valign='top'><a href=""javascript:insertLabel('{$PointUnit}')"" title=""��ȯ�ĵ�λ"">{$PointUnit}</a></td>" & vbCrLf
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
Response.write "        <td width=""130""  class='tdbg5' align=right height=25><STRONG>������ѭ����ǩHtml���룺</STRONG>" & vbCrLf
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
    Response.write "<option value=""0"" selected>��û���б�ģ�壡</option> " & vbCrLf
Else
    Response.write "<option value=""0"" selected>ѡ����õ�ģ����ʽ</option>" & vbCrLf
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
Response.write "          <input TYPE='button' value=' ȷ �� ' onCLICK='objectTag()'>&nbsp;&nbsp;" & vbCrLf
Response.write "          <input name='EditorpreviewContent' type='button' id='EditorpreviewContent' value=' Ԥ �� ' onclick='previewContent();'>" & vbCrLf
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
            strOption = "<option value='0'>�������κ�ר��</option>"
        Else
            strOption = "<option value='0' selected>�������κ�ר��</option>"
        End If
    End If
    If rsSpecial.bof And rsSpecial.bof Then
    Else
        Do While Not rsSpecial.EOF
            If rsSpecial("ChannelID") > 0 Then
                strOptionTemp = rsSpecial("SpecialName") & "����Ƶ����"
            Else
                strOptionTemp = rsSpecial("SpecialName") & "��ȫվ��"
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
    strOption = strOption & ">��ǰƵ��</option>"

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
    strChannel = strChannel & ">����ͬ��Ƶ��</option>"
    strChannel = strChannel & "<option value='ChannelID'"
    If ChannelID = "ChannelID" Then strChannel = strChannel & " selected"
    strChannel = strChannel & ">��ǰƵ��</option>"
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
    strClass_Option = strClass_Option & "<option value='0'>���������Ŀ</option>"
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
                    strClass_Option = strClass_Option & "��&nbsp;"
                Else
                    strClass_Option = strClass_Option & "��&nbsp;"
                End If
                Else
                If arrShowLine(i) = True Then
                    strClass_Option = strClass_Option & "��"
                Else
                    strClass_Option = strClass_Option & "&nbsp;"
                End If
                End If
            Next
            End If
            strClass_Option = strClass_Option & rsClass("ClassName")
            If rsClass("ClassType") = 2 Then
                strClass_Option = strClass_Option & "���⣩"
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
        strClass_Option = strClass_Option & "<option value='rsClass_arrChildID' " & classcss & " selected >��Ŀѭ���е���Ŀ</option>"
    Else
        strClass_Option = strClass_Option & "<option value='rsClass_arrChildID' " & classcss & " >��Ŀѭ���е���Ŀ</option>"
    End If
    If Trim(ClassID) = "ClassID" Then
        strClass_Option = strClass_Option & "<option value='ClassID' " & classcss & " selected>��ǰ��Ŀ������������Ŀ��</option>"
    Else
        strClass_Option = strClass_Option & "<option value='ClassID' " & classcss & ">��ǰ��Ŀ������������Ŀ��</option>"
    End If
    If Trim(ClassID) = "arrChildID" Then
        strClass_Option = strClass_Option & "<option value='arrChildID' " & classcss & " selected>��ǰ��Ŀ������Ŀ</option>"
    Else
        strClass_Option = strClass_Option & "<option value='arrChildID' " & classcss & ">��ǰ��Ŀ������Ŀ</option>"
    End If
    If Trim(ClassID) = "0" Then
        strClass_Option = strClass_Option & "<option value='0' " & classcss & " selected>��ʾ������Ŀ</option>"
    Else
        strClass_Option = strClass_Option & "<option value='0' " & classcss & ">��ʾ������Ŀ</option>"
    End If

    GetClass_Channel = strClass_Option
End Function
%>

