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

Dim Title, ModuleType, ChannelShortName, ChannelShowType, imageproperty, rs
Dim editLabel, arrParameter
Dim ClassID, NClassID, IncludeChild, SpecialID, Num, ProductType, IsHot, IsElite, AuthorName, DateNum
Dim OrderType, ShowType, TitleLen, ContentLen, ShowClassName, ShowPropertyType, ShowIncludePic, ShowAuthor
Dim ShowDateType, ShowHits, ShowHotSign, ShowNewSign, ShowTips, ShowCommentLink, UsePage, OpenType, Cols
Dim ImgWidth, ImgHeight, iTimeOut, urltype, CssNameA, CssName1, CssName2, effectID, IntervalLines
'�̳�
Dim ShowTableTitle, TableTitleStr, ShowProductModel, ShowProductStandard, ShowUnit, ShowStocksType, ShowPriceType
Dim ShowWeight, ShowPrice_Market, ShowPrice_Original, ShowPrice, ShowPrice_Member, ShowDiscount, ShowButtonType, ButtonStyle
Dim CssNameTable, CssNameTitle
'�˲���Ƹ
Dim PositionNum, IsUrgent, WorkPlaceNameLen, SubCompanyNameLen, PShowPoints, WShowPoints, SShowPoints, ShowPositionID, ShowPositionName, ShowWorkPlaceName, ShowSubCompanyName, ShowPositionNum, ShowPositionStatus, ShowValidDate, ShowUrgentSign, ShowNum

'��ģ�廹���Ҽ�
Dim InsertTemplate
Dim ChannelID, iChannelID, dChannelID
Dim NChannelID

ChannelID = Trim(Request("ChannelID"))
dChannelID = ReplaceLabelBadChar(Trim(Request("dChannelID")))

Select Case ChannelID
Case "ChannelID"
    ChannelID = "ChannelID"
Case ""
    ChannelID = ""
Case else
    If IsValidID(ChannelID) = False Then
        ChannelID = 0
    Else
        ChannelID = ReplaceLabelBadChar(ChannelID)
    End If  
End Select	

If InStr(ChannelID, ",") > 0 Then
    NChannelID = True
Else
    NChannelID = False
End If	

NClassID = False

If dChannelID = "" Then
   dChannelID = ChannelID
End If

If ChannelID = "" And iChannelID = "" Then
    Response.Write "Ƶ��������ʧ��"
    Response.End
End If

If ChannelID = "ChannelID" Then
    iChannelID = dChannelID
Else
    iChannelID = ChannelID
End If
Dim LabelName
LabelName = Trim(Request("LabelName"))
Title = Trim(Request("Title"))
ModuleType = PE_CLng(Trim(Request("ModuleType")))
ChannelShowType = Trim(Request("ChannelShowType"))
InsertTemplate = PE_CLng(Trim(Request("InsertTemplate")))

If SpecialID = "" Then SpecialID = 0

If Trim(request.querystring("editLabel")) <> "" Then
    editLabel = True
End If
Title = Trim(Request("Title"))

If Action = "Modify" Then
    Call GetLabelData

    If ChannelID = "ChannelID" Then
        iChannelID = PE_CLng(Trim(dChannelID))
    Else
        iChannelID = ChannelID
    End If
Else
    ModuleType = PE_CLng(Trim(Request("ModuleType")))
    ChannelShowType = Trim(Request("ChannelShowType"))
    InsertTemplate = PE_CLng(Trim(Request("InsertTemplate")))
    If Trim(Request("SpecialID")) = "SpecialID" Then
        SpecialID = Trim(Request("SpecialID"))
    Else
        SpecialID = PE_CLng(Trim(Request("SpecialID")))
    End If
    editLabel = PE_HtmlDecode(Trim(Request.Form("editLabel")))
    If ModuleType = 1 Then
        ChannelShortName = "����"
        imageproperty = "article"
    ElseIf ModuleType = 2 Then
        ChannelShortName = "���"
        imageproperty = "Soft"
    ElseIf ModuleType = 3 Then
        ChannelShortName = "ͼƬ"
        imageproperty = "Photo"
    ElseIf ModuleType = 5 Then
        iChannelID = 1000
        ChannelShortName = "��Ʒ"
        imageproperty = "Product"
    ElseIf ModuleType = 8 Then
        ChannelShortName = "ְλ"
        imageproperty = "Job"
    End If
End If
Response.Write "<html><head><title>" & Title & "</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='../Images/Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "<base target='_self'>"
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin=0 topmargin=0>" & vbCrLf
Response.Write "<form action='editor_label.asp' method='post' name='myform' id='myform'>"
Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
Response.Write "    <tr class='title'>"
Response.Write "      <td height='22' colspan='2' align='center'><strong>" & Title & "</strong></td>"
Response.Write "    </tr>"
If ModuleType <> 8 Then
    If ModuleType <> 5 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>����Ƶ����</strong></td>" & vbCrLf
        Response.Write "      <td height='25'><input type='hidden' name='iChannelID' value='" & ChannelID & "'><select name='ChannelID' "
        If NChannelID = True Then
            Response.Write "size='2' multiple style='height:250px;width:400px;'"
        Else
            Response.Write "size='1'"
        End If	
			
    Response.Write " onchange='document.myform.submit();'>" & GetChannel_Option(ModuleType, ChannelID) & "</select>"
    If ModuleType = 1 Then
        Response.Write " <input type='checkbox' name='NChannelID' value='1' onClick=""javascript:NChannelIDChild()"" "
        If NChannelID = True Then
            Response.Write " checked "
        End If
        Response.Write " >�Ƿ�ѡ����Ƶ��&nbsp;&nbsp;<font color='red'></font>"   
    End If				
		Response.write "</td>"
        Response.Write "    </tr>"
    End If
    If PE_CLng(iChannelID) > 0 Or ModuleType = 5 Or Instr(iChannelID,",")>0 Or Instr(iChannelID,"|")>0 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>������Ŀ��</strong></td>" & vbCrLf
        Response.Write "      <td height='25'><select name='ClassID' "
        If NClassID = True Then
            Response.Write "size='2' multiple style='height:250px;width:400px;'"
        Else
            Response.Write "size='1'"
        End If
        Response.Write ">" & GetClass_Channel(iChannelID, Trim(ClassID), NClassID) & "</select>"
        Response.Write " <input type='checkbox' name='IncludeChild' value='1' "
        If LCase(Trim(IncludeChild)) = "true" Then
            Response.Write " checked "
        End If
        Response.Write " >��������Ŀ&nbsp;&nbsp;<font color='red'><b>ע�⣺</b></font>����ָ��Ϊ�ⲿ��Ŀ </font>"
        Response.Write "  <br><input type='checkbox' name='NClassChild' value='1' onClick=""javascript:NClassIDChild()"" "
        If NClassID = True Then
            Response.Write " checked "
        End If
        Response.Write " >�Ƿ�ѡ������Ŀ&nbsp;&nbsp;<font color='red'><b>ע�⣺</b></font>��ɫ����Ŀ����ѡ </font>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>����ר�⣺</strong></td>"
        Response.Write "      <td height='25' ><select name='SpecialID' id='SpecialID'>" & GetSpecial_Option(iChannelID, SpecialID) & "</select></td>"
        Response.Write "    </tr>"
    Else
        Response.Write "<INPUT TYPE='hidden' name='ClassID' value='0' >"
        Response.Write "<INPUT TYPE='hidden' name='NClassChild' value='0' >"
        Response.Write "<INPUT TYPE='hidden' name='IncludeChild' value='true' >"
        Response.Write "<INPUT TYPE='hidden' name='SpecialID' value='0' >"
    End If

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>��ǩ˵����</strong></td>" & vbCrLf
    Response.Write "      <td height='25'><INPUT TYPE='text' NAME='lableExplain' value='' id='id' size='15' maxlength='20'>&nbsp;&nbsp;<FONT style='font-size:12px' color='blue'>����������д��ǩ��ʹ��˵�������Ժ�Ĳ���</FONT> </td>"
    Response.Write "    </tr>"
End If

Select Case ChannelShowType
Case "GetList"
    Call GetList
Case "GetPic"
    Call GetPic
Case "GetSlide"
    Call GetSlide
Case "GetPositionList"
    Call GetPositionList
Case "GetSearchResult"
    Call GetSearchResult
Case Else
    Response.Write "����Ĳ������"
    Response.End
End Select

Response.Write "    <tr class='tdbg'>"
Response.Write "      <td height='40' colspan='2' align='center'>"
Response.Write "        <input name='Title' type='hidden' id='Title' value='" & Title & "'>"
Response.Write "        <input name='LabelName' type='hidden' id='LabelName' value='" & LabelName & "'>"
Response.Write "        <input name='editLabel' type='hidden' id='editLabel' value='" & editLabel & "'>"
Response.Write "        <input name='dChannelID' type='hidden' id='dChannelID' value='" & dChannelID & "'>"
Response.Write "        <input name='ModuleType' type='hidden' id='ModuleType' value='" & ModuleType & "'>"
Response.Write "        <input name='InsertTemplate' type='hidden' id='InsertTemplate' value='" & InsertTemplate & "'>"
Response.Write "        <input name='ChannelShowType' type='hidden' id='ChannelShowType' value='" & ChannelShowType & "'>"
Response.Write "        <input name='MakeJS' type='button' id='MakeJS' onclick=""makejs('" & LabelName & "','" & ChannelShowType & "');"" value=' ȷ �� '>"
Response.Write "      </td>"
Response.Write "    </tr>"
Response.Write "  </table>"
Response.Write "</form>"
Response.Write "<script language=""JavaScript"" type=""text/JavaScript"">" & vbCrLf
Response.Write "function makejs(LabelName,Type)" & vbCrLf
Response.Write "{" & vbCrLf
If ModuleType <> 8 Then
    Response.Write "    if (document.myform.ClassID.value==''){" & vbCrLf
    Response.Write "        alert('" & ChannelShortName & "������Ŀ����ָ��Ϊ�ⲿ��Ŀ��');" & vbCrLf
    Response.Write "        document.myform.ClassID.focus();" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
End If
Response.Write "    var strJS;" & vbCrLf
If editLabel = "" And InsertTemplate = 0 Then
    If ModuleType <> 8 Then
        Response.Write "    if (document.myform.lableExplain.value !=""""){" & vbCrLf
        Response.Write "        strJS=""{$--""+document.myform.lableExplain.value+""--}"";" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS="""";" & vbCrLf
        Response.Write "    }" & vbCrLf
    Else
        Response.Write "    strJS="""";" & vbCrLf
    End If
    Response.Write "    strJS+=""<IMG  SRC='editor/images/label.gif' BORDER='0' "";" & vbCrLf
    Response.Write "    strJS+=""zzz='{$""+LabelName+""("";" & vbCrLf
Else
    Response.Write "    strJS=""{$""+LabelName+""("";" & vbCrLf
End If
Response.Write "  switch(Type){" & vbCrLf
Response.Write "  case ""GetList"":" & vbCrLf
If ModuleType <> 5  Then
    If ModuleType = 1 Then 
        Call CellNchannel
    Else
        Response.Write "    strJS+=document.myform.ChannelID.value;" & vbCrLf
    End If
    Response.Write "    strJS+="",""" & vbCrLf
End If

Call CellNclass

Response.Write "    strJS+="",""+document.myform.IncludeChild.checked;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.SpecialID.value;   " & vbCrLf
If ModuleType <> 5 Then
    Response.Write "    strJS+="",0""" & vbCrLf
End If
Response.Write "    strJS+="",""+document.myform.Num.value;" & vbCrLf
If ModuleType = 5 Then
    Response.Write "    strJS+="",""+document.myform.ProductType.value;" & vbCrLf
End If
Response.Write "    strJS+="",""+document.myform.IsHot.checked;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.IsElite.checked;" & vbCrLf
If ModuleType <> 5 Then
    Response.Write "    strJS+="",""+""\""""+document.myform.AuthorName.value+""\"""";" & vbCrLf
End If
Response.Write "    strJS+="",""+document.myform.DateNum.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.OrderType.value;" & vbCrLf

Response.Write "    strJS+="",""+document.myform.ShowType.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.TitleLen.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.ContentLen.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.ShowClassName.checked;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.ShowPropertyType.value;" & vbCrLf
If ModuleType = 1 Then
    Response.Write "    strJS+="",""+document.myform.ShowIncludePic.checked; //A" & vbCrLf
End If
If ModuleType <> 5 Then
    Response.Write "    strJS+="",""+document.myform.ShowAuthor.checked;" & vbCrLf
End If
Response.Write "    strJS+="",""+document.myform.ShowDateType.value;" & vbCrLf
If ModuleType <> 5 Then
    Response.Write "    strJS+="",""+document.myform.ShowHits.checked;" & vbCrLf
End If
Response.Write "    strJS+="",""+document.myform.ShowHotSign.checked;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.ShowNewSign.checked;" & vbCrLf
If ModuleType <> 5 Then
    Response.Write "    strJS+="",""+document.myform.ShowTips.checked;" & vbCrLf
End If
If ModuleType = 1 Then
    Response.Write "    strJS+="",""+document.myform.ShowCommentLink.checked; //A" & vbCrLf
End If
Response.Write "    strJS+="",""+document.myform.UsePage.checked;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.OpenType.value;" & vbCrLf
If ModuleType <> 5 Then
    Response.Write "    strJS+="",""+document.myform.Cols.value;            //A" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.CssNameA.value;        //A" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.CssName1.value;        //A" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.CssName2.value;        //A" & vbCrLf
End If
If ModuleType = 5 Then
    Response.Write "    strJS+="",""+document.myform.IntervalLines.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.Cols.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowTableTitle.checked;" & vbCrLf
    Response.Write "    var TableTitleStr=""""" & vbCrLf
    Response.Write "    for(var i=1;i<14;i++){" & vbCrLf
    Response.Write "        if (i==13){" & vbCrLf
    Response.Write "            TableTitleStr+=eval(""document.myform.TableTitleStr""+i+"".value"")" & vbCrLf
    Response.Write "        }else{" & vbCrLf
    Response.Write "            TableTitleStr+=eval(""document.myform.TableTitleStr""+i+"".value"")+""|""" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    strJS+="",""+TableTitleStr" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowProductModel.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowProductStandard.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowUnit.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowStocksType.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowWeight.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowPrice_Market.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowPrice_Original.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowPrice.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowPrice_Member.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowDiscount.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowButtonType.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ButtonStyle.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.CssNameTable.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.CssNameTitle.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.CssNameA.value;        //A" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.CssName1.value;        //A" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.CssName2.value;        //A" & vbCrLf
End If
Response.Write "    break;" & vbCrLf

Response.Write "   case ""GetPic"":" & vbCrLf
If ModuleType <> 5  Then
    If ModuleType = 1 Then 
        Call CellNchannel
    Else
        Response.Write "    strJS+=document.myform.ChannelID.value;" & vbCrLf
    End If
    Response.Write "    strJS+="",""" & vbCrLf
End If
Call CellNclass
Response.Write "    strJS+="",""+document.myform.IncludeChild.checked;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.SpecialID.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.Num.value;" & vbCrLf
If ModuleType = 5 Then
    Response.Write "    strJS+="",""+document.myform.ProductType.value;" & vbCrLf
End If
Response.Write "    strJS+="",""+document.myform.IsHot.checked;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.IsElite.checked;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.DateNum.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.OrderType.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.ShowType.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.ImgWidth.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.ImgHeight.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.TitleLen.value;" & vbCrLf
If ModuleType <> 5 Then
    Response.Write "    strJS+="",""+document.myform.ContentLen.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowTips.checked;" & vbCrLf
End If
Response.Write "    strJS+="",""+document.myform.Cols.value;" & vbCrLf
If ModuleType = 5 Then
    Response.Write "    strJS+="",""+document.myform.ShowPriceType.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowDiscount.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowButtonType.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ButtonStyle.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.OpenType.value;" & vbCrLf
End If
Response.Write "    break;" & vbCrLf

Response.Write "   case ""GetSlide"":" & vbCrLf
If ModuleType <> 5  Then
    If ModuleType = 1 Then 
        Call CellNchannel
    Else
        Response.Write "    strJS+=document.myform.ChannelID.value;" & vbCrLf
    End If
    Response.Write "    strJS+="",""" & vbCrLf
End If
Call CellNclass
Response.Write "    strJS+="",""+document.myform.IncludeChild.checked;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.SpecialID.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.Num.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.IsHot.checked;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.IsElite.checked;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.DateNum.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.OrderType.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.ImgWidth.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.ImgHeight.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.TitleLen.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.iTimeOut.value;" & vbCrLf
Response.Write "    strJS+="",""+document.myform.effectID.value;" & vbCrLf
'If ModuleType = 5 Then
'    Response.Write "    strJS+="",""+document.myform.OpenType.value;" & vbCrLf
'End If
Response.Write "    break;" & vbCrLf

If ModuleType = 8 Then
    Response.Write "  case ""GetPositionList"":" & vbCrLf
    Response.Write "    strJS+=document.myform.PositionNum.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.IsUrgent.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.DateNum.value;   " & vbCrLf
    Response.Write "    strJS+="",""+document.myform.OrderType.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowType.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.TitleLen.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.WorkPlaceNameLen.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.SubCompanyNameLen.value;" & vbCrLf
    Response.Write "    if (document.myform.PShowPoints.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.PShowPoints.checked;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""false"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.WShowPoints.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.WShowPoints.checked;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""false"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.SShowPoints.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.SShowPoints.checked;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""false"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowDateType.value;" & vbCrLf
    Response.Write "    if (document.myform.ShowPositionID.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowPositionID.value;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""0"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.ShowPositionName.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowPositionName.value;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""0"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.ShowWorkPlaceName.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowWorkPlaceName.value;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""0"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.ShowSubCompanyName.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowSubCompanyName.value;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""0"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.ShowPositionNum.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowPositionNum.value;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""0"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.ShowPositionStatus.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowPositionStatus.value;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""0"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.ShowValidDate.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowValidDate.value;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""0"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.ShowUrgentSign.checked ==false||document.myform.ShowType.value==2||document.myform.ShowType.value==3){" & vbCrLf
    Response.Write "        strJS+="",""+""false"";" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowUrgentSign.value;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.ShowNewSign.checked ==false||document.myform.ShowType.value==1||document.myform.ShowType.value==3){" & vbCrLf
    Response.Write "        strJS+="",""+""false"";" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowNewSign.value;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.ShowType.value==1||document.myform.ShowType.value==2){" & vbCrLf
    Response.Write "        strJS+="",""+""false"";" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.UsePage.value;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.OpenType.value;" & vbCrLf
    Response.Write "    break;" & vbCrLf

    Response.Write "  case ""GetSearchResult"":" & vbCrLf
    Response.Write "    strJS+=document.myform.ShowNum.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.OrderType.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.TitleLen.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.WorkPlaceNameLen.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.SubCompanyNameLen.value;" & vbCrLf
    Response.Write "    if (document.myform.PShowPoints.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.PShowPoints.checked;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""false"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.WShowPoints.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.WShowPoints.checked;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""false"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.SShowPoints.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.SShowPoints.checked;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""false"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowDateType.value;" & vbCrLf
    Response.Write "    if (document.myform.ShowPositionID.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowPositionID.value;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""0"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.ShowPositionName.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowPositionName.value;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""0"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.ShowWorkPlaceName.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowWorkPlaceName.value;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""0"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.ShowSubCompanyName.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowSubCompanyName.value;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""0"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.ShowPositionNum.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowPositionNum.value;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""0"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.ShowPositionStatus.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowPositionStatus.value;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""0"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (document.myform.ShowValidDate.checked ==true){" & vbCrLf
    Response.Write "        strJS+="",""+document.myform.ShowValidDate.value;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+="",""+""0"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.UsePage.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.OpenType.value;" & vbCrLf
    Response.Write "    break;" & vbCrLf
End If
Response.Write "    default:" & vbCrLf
Response.Write "        alert(""����������ã�"");" & vbCrLf
Response.Write "        break;" & vbCrLf
Response.Write "   }" & vbCrLf
If editLabel = "" And InsertTemplate = 0 Then
    Response.Write "   strJS+="")}' >"";" & vbCrLf
Else
    Response.Write "   strJS+="")}"";" & vbCrLf
End If
Response.Write "   window.returnValue = strJS;" & vbCrLf
Response.Write "   window.close();" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "</script>" & vbCrLf

%>
<script Language="JavaScript">
function NClassIDChild(){
    if (document.myform.NClassChild.checked==true){
        document.myform.ClassID.size=2;
        document.myform.ClassID.style.height=250;
        document.myform.ClassID.style.width=400;
        document.myform.ClassID.multiple=true;
        for(var i=0;i<document.myform.ClassID.length;i++){
            if (document.myform.ClassID.options[i].value=="rsClass_arrChildID"||document.myform.ClassID.options[i].value=="ClassID"||document.myform.ClassID.options[i].value=="arrChildID"||document.myform.ClassID.options[i].value=="0"){
                document.myform.ClassID.options[i].style.background="red";
            }
        }
    }else{
        document.myform.ClassID.size=1;
        document.myform.ClassID.style.width=200;
        document.myform.ClassID.multiple=false;
        for(var i=0;i<document.myform.ClassID.length;i++){
            if (document.myform.ClassID.options[i].value=="rsClass_arrChildID"||document.myform.ClassID.options[i].value=="ClassID"||document.myform.ClassID.options[i].value=="arrChildID"||document.myform.ClassID.options[i].value=="0"){
                document.myform.ClassID.options[i].style.background="";
            }
        }
    }
}
function NChannelIDChild(){
    if (document.myform.NChannelID.checked==true){
        document.myform.ChannelID.size=2;
        document.myform.ChannelID.style.height=250;
        document.myform.ChannelID.style.width=400;
        document.myform.ChannelID.multiple=true;
    }else{
        document.myform.ChannelID.size=1;
        document.myform.ChannelID.style.width=150;
        document.myform.ChannelID.multiple=false;
    }
}
function change_item(element){
    if(element.selectedIndex!=-1)
    var selectednumber = element.options[element.selectedIndex].value;

    if(selectednumber==1){
        objFiles.style.display="";
        <%
        If ModuleType = 5 Then
        %>
            document.myform.common.src = "<%=InstallDir%>Shop/images/<%=imageproperty%>_common.gif"
            document.myform.Elite.src = "<%=InstallDir%>Shop/images/<%=imageproperty%>_Elite.gif"
            document.myform.OnTop.src = "<%=InstallDir%>Shop/images/<%=imageproperty%>_OnTop.gif"
        <%
        Else
        %>
            document.myform.common.src = "<%=InstallDir & imageproperty%>/images/<%=imageproperty%>_common.gif"
            document.myform.Elite.src = "<%=InstallDir & imageproperty%>/images/<%=imageproperty%>_Elite.gif"
            document.myform.OnTop.src = "<%=InstallDir & imageproperty%>/images/<%=imageproperty%>_OnTop.gif"
        <%
        End If
        %>
    }
    else if (selectednumber==0)
    {
        objFiles.style.display="none";
    }
    else if (selectednumber==2)
    {
        objFiles.style.display="none";
    }
    else if (selectednumber>=3)
    {
        selectednumber = selectednumber - 1
        objFiles.style.display="";
        <%
        If ModuleType = 5 Then
        %>
            document.myform.common.src = "<%=InstallDir%>Shop/images/<%=imageproperty%>_common" + selectednumber + ".gif"
            document.myform.Elite.src = "<%=InstallDir%>Shop/images/<%=imageproperty%>_Elite" + selectednumber + ".gif"
            document.myform.OnTop.src = "<%=InstallDir%>Shop/images/<%=imageproperty%>_OnTop" + selectednumber + ".gif"
        <%
        Else
        %>
            document.myform.common.src = "<%=InstallDir & imageproperty%>/images/<%=imageproperty%>_common" + selectednumber + ".gif"
            document.myform.Elite.src = "<%=InstallDir & imageproperty%>/images/<%=imageproperty%>_Elite" + selectednumber + ".gif"
            document.myform.OnTop.src = "<%=InstallDir & imageproperty%>/images/<%=imageproperty%>_OnTop" + selectednumber + ".gif"
        <%
        End If
        %>
    }
}
</script>

<%
Response.Write "</body>"
Response.Write "</html>"
Call CloseConn

Sub GetList()

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>��ʾ��ʽ��</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ShowType' id='ShowType'>"
    Response.Write "           <option value='1' "
    If Trim(ShowType) = "1" Then Response.Write "selected"
    Response.Write ">��ͨ�б�</option>"
    Response.Write "           <option value='2' "
    If Trim(ShowType) = "2" Then Response.Write "selected"
    Response.Write ">���ʽ</option>"
    Response.Write "           <option value='3' "
    If Trim(ShowType) = "3" Then Response.Write "selected"
    Response.Write ">�������ʽ</option>"
    If ModuleType = 1 Then
        Response.Write "           <option value='4' "
        If Trim(ShowType) = "4" Then Response.Write "selected"
        Response.Write ">���ܶ���ʽ</option>"
        Response.Write "           <option value='5' "
        If Trim(ShowType) = "5" Then Response.Write "selected"
        Response.Write ">���DIV��ʽ</option>"
        Response.Write "           <option value='6' "
        If Trim(ShowType) = "6" Then Response.Write "selected"
        Response.Write ">���RSS��ʽ</option>"
    Else
        Response.Write "           <option value='4' "
        If Trim(ShowType) = "4" Then Response.Write "selected"
        Response.Write ">���DIV��ʽ</option>"
        Response.Write "           <option value='5' "
        If Trim(ShowType) = "5" Then Response.Write "selected"
        Response.Write ">���RSS��ʽ</option>"
    End If
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "��Ŀ��</strong></td>"
    Response.Write "      <td height='25'><input name='Num' type='text' value='"
    If Trim(Num) = "" Then
        Response.Write "10"
    Else
        Response.Write Num
    End If
    Response.Write "' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>���Ϊ0������ʾ����" & ChannelShortName & "��</font></td>"
    Response.Write "    </tr>"
    If ModuleType = 5 Then
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong> ��Ʒ���ͣ�</strong></td>"
    Response.Write "      <td height='25'><select name='ProductType' id='ProductType'>"
    Response.Write "        <option value='1'"
    If Trim(ProductType) = "1" Then Response.Write "selected"
    Response.Write ">����������Ʒ</option>"
    Response.Write "        <option value='2'"
    If Trim(ProductType) = "2" Then Response.Write "selected"
    Response.Write ">�Ǽ���Ʒ</option>"
    Response.Write "        <option value='3'"
    If Trim(ProductType) = "3" Then Response.Write "selected"
    Response.Write ">�ؼ���Ʒ</option>"
    Response.Write "        <option value='4'"
    If Trim(ProductType) = "4" Then Response.Write "selected"
    Response.Write ">������Ʒ</option>"
    Response.Write "        <option value='5'"
    If Trim(ProductType) = "5" Then Response.Write "selected"
    Response.Write ">�������ۺ��Ǽ���Ʒ</option>"
    Response.Write "        <option value='6'"
    If Trim(ProductType) = "6" Then Response.Write "selected"
    Response.Write ">������Ʒ</option>"
    Response.Write "        <option value='0'"
    If Trim(ProductType) = "0" Then Response.Write "selected"
    Response.Write ">������Ʒ</option>"
    Response.Write "        </select> </td>"
    Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "���ԣ�</strong></td>"
    Response.Write "      <td height='25'><input name='IsHot' type='checkbox' id='IsHot' value='1'"
    If LCase(Trim(IsHot)) = "true" Then Response.Write "checked"
    Response.Write ">"
    Response.Write "        ����" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<input name='IsElite' type='checkbox' id='IsElite' value='1'"
    If LCase(Trim(IsElite)) = "true" Then Response.Write "checked"
    Response.Write ">"
    Response.Write "        �Ƽ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>�������ѡ������ʾ����" & ChannelShortName & "��</font></td>"
    Response.Write "    </tr>"
    If ModuleType <> 5 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>����������</strong></td>"
        Response.Write "      <td height='25'><input name='AuthorName' type='text' value='"
        If Trim(AuthorName) = """" Then
            Response.Write ""
        Else
            Response.Write AuthorName
        End If
        Response.Write "'  size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>�����Ϊ�գ���ֻ��ʾָ�����ߵ�" & ChannelShortName & "�����ڸ����ļ���</font></td>"
        Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "����ͼƬ��</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' height='100%' valign='top'>"
    Response.Write "        <tr>"
    Response.Write "          <td width='100'>"
    Response.Write "            <select name='ShowPropertyType' id='ShowPropertyType' onChange=""javascript:change_item(this)"">"
    Response.Write "           <option value='0' "
    If Trim(ShowPropertyType) = "0" Then Response.Write "selected"
    Response.Write ">����ʾ</option>"
    Response.Write "           <option value='2' "
    If Trim(ShowPropertyType) = "2" Then Response.Write "selected"
    Response.Write ">����</option>"
    Response.Write "           <option value='1' "
    If Trim(ShowPropertyType) = "1" Then Response.Write "selected"
    Response.Write ">СͼƬ����ʽ 1��</option>"
    Response.Write "           <option value='3' "
    If Trim(ShowPropertyType) = "3" Then Response.Write "selected"
    Response.Write ">СͼƬ����ʽ 2��</option>"
    Response.Write "           <option value='4' "
    If Trim(ShowPropertyType) = "4" Then Response.Write "selected"
    Response.Write ">СͼƬ����ʽ 3��</option>"
    Response.Write "           <option value='5' "
    If Trim(ShowPropertyType) = "5" Then Response.Write "selected"
    Response.Write ">СͼƬ����ʽ 4��</option>"
    Response.Write "           <option value='6' "
    If Trim(ShowPropertyType) = "6" Then Response.Write "selected"
    Response.Write ">СͼƬ����ʽ 5��</option>"
    If ModuleType = 1 Then
        Response.Write "           <option value='7' "
        If Trim(ShowPropertyType) = "7" Then Response.Write "selected"
        Response.Write ">СͼƬ����ʽ 6��</option>"
        Response.Write "           <option value='8' "
        If Trim(ShowPropertyType) = "8" Then Response.Write "selected"
        Response.Write ">СͼƬ����ʽ 7��</option>"
        Response.Write "           <option value='9' "
        If Trim(ShowPropertyType) = "9" Then Response.Write "selected"
        Response.Write ">СͼƬ����ʽ 8��</option>"
        Response.Write "           <option value='10' "
        If Trim(ShowPropertyType) = "10" Then Response.Write "selected"
        Response.Write ">СͼƬ����ʽ 9��</option>"
    End If
    Response.Write "        </select>"
    Response.Write "         </td>"
    Response.Write "          <td id=objFiles style='display:none'>"
    Response.Write "&nbsp;&nbsp;��ͨͼƬ&nbsp;&nbsp;<IMG id=common SRC='" & InstallDir & "images/" & imageproperty & "_common.gif' BORDER='0' ALT='��ͨͼƬ'>&nbsp;&nbsp;�Ƽ�ͼƬ&nbsp;&nbsp;<IMG SRC='" & InstallDir & "images/" & imageproperty & "_Elite.gif' id=Elite BORDER='0' ALT='�Ƽ�ͼƬ'>&nbsp;&nbsp;�̶�ͼƬ&nbsp;&nbsp;<IMG SRC='" & InstallDir & "images/" & imageproperty & "_OnTop.gif' id=OnTop BORDER='0' ALT='�̶�ͼƬ'>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>���ڷ�Χ��</strong></td>"
    Response.Write "      <td height='25'>ֻ��ʾ���"
    Response.Write "        <input name='DateNum' type='text' id='DateNum' value="
    If Trim(DateNum) = "" Then
        Response.Write "0"
    Else
        Response.Write DateNum
    End If
    Response.Write " size='5' maxlength='3'>"
    Response.Write "        ���ڸ��µ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>&nbsp;&nbsp;���Ϊ�ջ�0������ʾ����������" & ChannelShortName & "��</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>���򷽷���</strong></td>"
    Response.Write "      <td height='25'><select name='OrderType' id='OrderType'>"
    Response.Write "       <option value='1' "
    If Trim(OrderType) = "1" Then Response.Write "selected"
    Response.Write ">" & ChannelShortName & "ID������</option>"
    Response.Write "       <option value='2' "
    If Trim(OrderType) = "2" Then Response.Write "selected"
    Response.Write ">" & ChannelShortName & "ID������</option>"
    Response.Write "       <option value='3' "
    If Trim(OrderType) = "3" Then Response.Write "selected"
    Response.Write ">����ʱ�䣨����</option>"
    Response.Write "       <option value='4' "
    If Trim(OrderType) = "4" Then Response.Write "selected"
    Response.Write ">����ʱ�䣨����</option>"
    Response.Write "       <option value='5' "
    If Trim(OrderType) = "5" Then Response.Write "selected"
    Response.Write ">�������������</option>"
    Response.Write "       <option value='6' "
    If Trim(OrderType) = "6" Then Response.Write "selected"
    Response.Write ">�������������</option>"
    Response.Write "       <option value='7' "
    If Trim(OrderType) = "7" Then Response.Write "selected"
    Response.Write ">��������������</option>"
    Response.Write "       <option value='8' "
    If Trim(OrderType) = "8" Then Response.Write "selected"
    Response.Write ">��������������</option>"
    Response.Write "      </select></td>"
    Response.Write "    </tr>"
    Response.Write " <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>��������ַ�����</strong></td>"
    Response.Write "      <td height='25'><input name='TitleLen' type='text' id='TitleLen' value="
    If Trim(TitleLen) = "" Then
        Response.Write "30"
    Else
        Response.Write TitleLen
    End If
    Response.Write "  size='5' maxlength='3'>"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>���Ϊ0������ʾ�������⡣��ĸ��һ���ַ��������������ַ���</font></td>"
    Response.Write "    </tr>"

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "�����ַ�����</strong></td>"
    Response.Write "      <td height='25'><input name='ContentLen' type='text' id='ContentLen' value="
    If Trim(ContentLen) = "" Then
        Response.Write "0"
    Else
        Response.Write ContentLen
    End If
    Response.Write "  size='5' maxlength='3'>"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>�������0�����ڱ����·�����ʾָ��������" & ChannelShortName & "����</font></td>"
    Response.Write "    </tr>"
    'If ModuleType = 1 Or ModuleType = 5 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>ÿ�е�������</strong></td>"
        Response.Write "      <td height='25'><INPUT TYPE='text' NAME='Cols' value="
        If Trim(Cols) = "" Then
            Response.Write "1"
        Else
            Response.Write Cols
        End If
        Response.Write "  id='id' size='5' maxlength='3'> &nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>�����������ͻ���</font>"
        Response.Write "      <input type='hidden' name='urltype' value='0'></td>"
        Response.Write "    </tr>"
    'End If
    Response.Write " <tr class='tdbg'>"
    Response.Write "      <td height='50' align='right' class='tdbg5'><strong>��ʾ���ݣ�</strong></td>"
    Response.Write "      <td height='50'><table width='100%' border='0' cellpadding='1' cellspacing='2'>"
    Response.Write "        <tr>"
    Response.Write "          <td><input name='ShowClassName' type='checkbox' id='ShowClassName' value='1' "
    If LCase(Trim(ShowClassName)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">������Ŀ</td>"
    If ModuleType = 1 Then
        Response.Write "          <td><input name='ShowIncludePic' type='checkbox' id='ShowIncludePic' value='1' "
        If LCase(Trim(ShowIncludePic)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">��ͼ�ġ���־</td>"
    End If
    If ModuleType <> 5 Then
        Response.Write "          <td><input name='ShowAuthor' type='checkbox' id='ShowAuthor' value='1' "
        If LCase(Trim(ShowAuthor)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">����</td>"
    End If
    Response.Write "          <td>����ʱ��"
    Response.Write "              <select name='ShowDateType' id='ShowDateType'>"
    Response.Write "                <option value='0' "
    If Trim(ShowDateType) = "0" Then Response.Write "selected"
    Response.Write ">����ʾ</option>"
    Response.Write "                <option value='1' "
    If Trim(ShowDateType) = "1" Then Response.Write "selected"
    Response.Write ">������</option>"
    Response.Write "                <option value='2'"
    If Trim(ShowDateType) = "2" Then Response.Write "selected"
    Response.Write ">����</option>"
    Response.Write "                <option value='3' "
    If Trim(ShowDateType) = "3" Then Response.Write "selected"
    Response.Write ">��-��</option>"
    Response.Write "              </select>"
    Response.Write "          </td>"
    If ModuleType <> 5 Then
        Response.Write "          <td><input name='ShowHits' type='checkbox' id='ShowHits' value='1' "
        If LCase(Trim(ShowHits)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write " >�������</td>"
    End If
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td><input name='ShowHotSign' type='checkbox' id='ShowHotSign' value='1' "
    If LCase(Trim(ShowHotSign)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">�ȵ�" & ChannelShortName & "��־</td>"
    Response.Write "          <td><input name='ShowNewSign' type='checkbox' id='ShowNewSign' value='1' "
    If LCase(Trim(ShowNewSign)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">����" & ChannelShortName & "��־</td>"
    If ModuleType <> 5 Then
        Response.Write "          <td><input name='ShowTips' type='checkbox' id='ShowTips' value='1' "
        If LCase(Trim(ShowTips)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">��ʾ��ʾ��Ϣ</td>"
    End If
    If ModuleType = 1 Then
        Response.Write "          <td><input name='ShowCommentLink' type='checkbox' id='ShowCommentLink' value='1' "
        If LCase(Trim(ShowCommentLink)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">��ʾ��������</td>"
    End If
    Response.Write "          <td><input name='UsePage' type='checkbox' id='UsePage' value='1'"
    If LCase(Trim(UsePage)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">�Ƿ��ҳ��ʾ</td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "�򿪷�ʽ��</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='OpenType' id='OpenType'>"
    Response.Write "          <option value='0' "
    If Trim(OpenType) = "0" Then Response.Write "selected"
    Response.Write ">��ԭ���ڴ�</option>"
    Response.Write "          <option value='1' "
    If Trim(OpenType) = "1" Then Response.Write "selected"
    Response.Write ">���´��ڴ�</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    If ModuleType = 5 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>ÿ�������пհ�һ�У�</strong></td>"
        Response.Write "      <td height='25'><input name='IntervalLines' type='text' value='"
        If Trim(IntervalLines) = """" Then
            Response.Write ""
        Else
            Response.Write IntervalLines
        End If
        Response.Write "'  size='10' maxlength='10'>&nbsp;<font color=blue>Ϊ0ʱ������</font></td>"
        Response.Write "    </tr>"
        Response.Write "     <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>���ͷ�����֣�</strong></td>"
        Response.Write "      <td height='25' >"
        If Trim(TableTitleStr) = "" Or InStr(TableTitleStr, "|") <= 0 Or UBound(Split(TableTitleStr, "|")) > 12 Or UBound(Split(TableTitleStr, "|")) < 12 Then
            TableTitleStr = "��Ʒ����|�ͺ�|���|����ʱ��|��λ|�����|����|�г���|�̳Ǽ�|�Żݼ�|��Ա��|�ۿ���|����"
        End If
        TableTitleStr = Split(TableTitleStr, "|")
        Response.Write "<table border='0' cellpadding='0' cellspacing='0' width='100%' height='100%' align='center'>"
        Response.Write " <tr class='tdbg' align='center'>"
        Response.Write "    <td>��Ʒ����</td><td>�ͺ�</td><td>���</td><td>����ʱ��</td><td>��λ</td><td>�����</td><td>����</td>"
        Response.Write " </tr>"
        Response.Write " <tr class='tdbg' align='center'>"
        Response.Write "    <td><input name='TableTitleStr1' type='text' value='" & TableTitleStr(0) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr2' type='text' value='" & TableTitleStr(1) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr3' type='text' value='" & TableTitleStr(2) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr4' type='text' value='" & TableTitleStr(3) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr5' type='text' value='" & TableTitleStr(4) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr6' type='text' value='" & TableTitleStr(5) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr7' type='text' value='" & TableTitleStr(6) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write " </tr>"
        Response.Write "  <tr class='tdbg' align='center'>"
        Response.Write "    <td>�г���</td><td>�̳Ǽ�</td><td>�Żݼ�</td><td>��Ա��</td><td>�ۿ���</td><td>����</td>"
        Response.Write " </tr>"
        Response.Write "  <tr class='tdbg' align='center'>"
        Response.Write "    <td><input name='TableTitleStr8' type='text' value='" & TableTitleStr(7) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr9' type='text' value='" & TableTitleStr(8) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr10' type='text' value='" & TableTitleStr(9) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr11' type='text' value='" & TableTitleStr(10) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr12' type='text' value='" & TableTitleStr(11) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr13' type='text' value='" & TableTitleStr(12) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "  </tr>"
        Response.Write " </table>"
        Response.Write "     </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>��ʾ��Ʒ��淽ʽ��</strong></td>"
        Response.Write "      <td height='25'><select name='ShowStocksType' id='ShowStocksType'>"
        Response.Write "       <option value='0' "
        If Trim(ShowStocksType) = "0" Then Response.Write "selected"
        Response.Write ">����ʾ</option>"
        Response.Write "       <option value='1' "
        If Trim(ShowStocksType) = "1" Then Response.Write "selected"
        Response.Write ">��ʾ������</option>"
        Response.Write "       <option value='2' "
        If Trim(ShowStocksType) = "2" Then Response.Write "selected"
        Response.Write ">��ʾʵ�ʿ��</option>"
        Response.Write "      </select></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>��ť��ʾ��ʽ��</strong></td>"
        Response.Write "      <td height='25'><select name='ShowButtonType' id='ShowButtonType'>"
        Response.Write "       <option value='0' "
        If Trim(ShowButtonType) = "0" Then Response.Write "selected"
        Response.Write ">����ʾ</option>"
        Response.Write "       <option value='1' "
        If Trim(ShowButtonType) = "1" Then Response.Write "selected"
        Response.Write ">��ʾ����ť</option>"
        Response.Write "       <option value='2' "
        If Trim(ShowButtonType) = "2" Then Response.Write "selected"
        Response.Write ">��ʾ��ϸ��ť</option>"
        Response.Write "       <option value='3' "
        If Trim(ShowButtonType) = "3" Then Response.Write "selected"
        Response.Write ">��ʾ�ղذ�ť</option>"
        Response.Write "       <option value='4' "
        If Trim(ShowButtonType) = "4" Then Response.Write "selected"
        Response.Write ">��ʾ������ϸ��ť</option>"
        Response.Write "       <option value='5' "
        If Trim(ShowButtonType) = "5" Then Response.Write "selected"
        Response.Write ">��ʾ�����ղذ�ť</option>"
        Response.Write "       <option value='6' "
        If Trim(ShowButtonType) = "6" Then Response.Write "selected"
        Response.Write ">��ϸ���ղذ�ť</option>"
        Response.Write "       <option value='7' "
        If Trim(ShowButtonType) = "7" Then Response.Write "selected"
        Response.Write ">��������ʾ</option>"
        Response.Write "      </select></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='50' align='right' class='tdbg5'><strong>��ʾ��Ʒ��ϸ��Ϣ��</strong></td>"
        Response.Write "      <td height='50'><table width='100%' border='0' cellpadding='1' cellspacing='2'>"
        Response.Write "        <tr>"
        Response.Write "          <td><input name='ShowTableTitle' type='checkbox' id='ShowTableTitle' value='1' "
        If LCase(Trim(ShowTableTitle)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">���ͷ������</td>"
        Response.Write "          <td><input name='ShowProductModel' type='checkbox' id='ShowProductModel' value='1' "
        If LCase(Trim(ShowProductModel)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">�Ƿ���ʾ��Ʒ�ͺ�</td>"
        Response.Write "          <td><input name='ShowProductStandard' type='checkbox' id='ShowProductStandard' value='1' "
        If LCase(Trim(ShowProductStandard)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">�Ƿ���ʾ��Ʒ���</td>"
        Response.Write "        </tr>"
        Response.Write "        <tr>"
        Response.Write "          <td><input name='ShowUnit' type='checkbox' id='ShowUnit' value='1' "
        If LCase(Trim(ShowUnit)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">�Ƿ���ʾ��Ʒ��λ</td>"
        Response.Write "          <td><input name='ShowWeight' type='checkbox' id='ShowWeight' value='1' "
        If LCase(Trim(ShowWeight)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">�Ƿ���ʾ��Ʒ����</td>"
        Response.Write "          <td><input name='ShowPrice_Market' type='checkbox' id='ShowPrice_Market' value='1' "
        If LCase(Trim(ShowPrice_Market)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">�Ƿ���ʾ�г���</td>"
        Response.Write "        </tr>"
        Response.Write "      <tr>"
        Response.Write "          <td><input name='ShowPrice_Original' type='checkbox' id='ShowPrice_Original' value='1' "
        If LCase(Trim(ShowPrice_Original)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">�Ƿ���ʾԭ��</td>"

        Response.Write "          <td><input name='ShowPrice' type='checkbox' id='ShowPrice' value='1' "
        If LCase(Trim(ShowPrice)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">�Ƿ���ʾ��ǰ���ۼ�</td>"

        Response.Write "          <td><input name='ShowPrice_Member' type='checkbox' id='ShowPrice_Member' value='1' "
        If LCase(Trim(ShowPrice_Member)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">�Ƿ���ʾ��Ա��</td>"

        Response.Write "          <td><input name='ShowDiscount' type='checkbox' id='ShowDiscount' value='1' "
        If LCase(Trim(ShowDiscount)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">�Ƿ���ʾ�ۿ���</td>"
        Response.Write "        </tr>"
        Response.Write "      </table>"
        Response.Write "     </td>"
        Response.Write "   </tr>"
        Response.Write "   <tr class='tdbg'>"
        Response.Write "     <td height='25' align='right' class='tdbg5'><strong>��ť��ʽ��</strong></td>"
        Response.Write "     <td height='25' ><input name='ButtonStyle' type='text' value='"
        If Trim(ButtonStyle) = """" Then
            Response.Write ""
        Else
            Response.Write ButtonStyle
        End If
        Response.Write "'  size='10' maxlength='20'>&nbsp;&nbsp;<font color='blue'>����д����ͼƬ����</font><br>"
        Response.Write "������<br>"
        Response.Write "��" & InstallDir & "Shop/images/ProductBuy<FONT color='blue'>�����֡�</FONT>.gif<br>"
        Response.Write "��" & InstallDir & "Shop/images/ProductContent<FONT color='blue'>�����֡�</FONT>.gif<br>"
        Response.Write "��" & InstallDir & "Shop/images/ProductFav<FONT color='blue'>�����֡�</FONT>.gif<br>"
        Response.Write "&nbsp;&nbsp;<font color='blue'>�밴���Ϸ�ʽ�����ϴ��Զ��尴ťͼƬ</font></td>"
        Response.Write "   </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>���CSS��</strong></td>"
        Response.Write "      <td height='25'><input name='CssNameTable' type='text' value='"
        If Trim(CssNameTable) = """" Then
            Response.Write ""
        Else
            Response.Write CssNameTable
        End If
        Response.Write "'  size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>����CSS��������ѡ����(���ڱ��ʽ��Ч)</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>���ͷ��CSS��</strong></td>"
        Response.Write "      <td height='25'><input name='CssNameTitle' type='text' value='"
        If Trim(CssNameTitle) = """" Then
            Response.Write ""
        Else
            Response.Write CssNameTitle
        End If
        Response.Write "'  size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>���ͷ���е�CSS��������ѡ������(���ڱ��ʽ��Ч)</font></td>"
        Response.Write "    </tr>"
    End If
    'If ModuleType = 1 Or ModuleType = 5 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>CSS������</strong></td>"
        Response.Write "      <td height='25'><input name='CssNameA' type='text' value='"
        If Trim(CssNameA) = """" Then
            Response.Write ""
        Else
            Response.Write CssNameA
        End If
        Response.Write "'  size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>�б����������ӵ��õ�CSS��������ѡ����(���ڱ��ʽ��Ч)</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>�����ʽ1��</strong></td>"
        Response.Write "      <td height='25'><input name='CssName1' type='text' value='"
        If Trim(CssName1) = """" Then
            Response.Write ""
        Else
            Response.Write CssName1
        End If
        Response.Write "'  size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>�б��������е�CSSЧ������������ѡ����(���ڱ��ʽ��Ч)</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>�����ʽ2��</strong></td>"
        Response.Write "      <td height='25'><input name='CssName2' type='text' value='"
        If Trim(CssName2) = """" Then
            Response.Write ""
        Else
            Response.Write CssName2
        End If
        Response.Write "'  size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>�б���ż���е�CSSЧ������������ѡ����(���ڱ��ʽ��Ч)</font></td>"
        Response.Write "    </tr>"
   ' End If
End Sub

Sub GetPic()

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "��Ŀ��</strong></td>"
    Response.Write "      <td height='25'><input name='Num' type='text' value="
    If Trim(Num) = "" Then
        Response.Write "4"
    Else
        Response.Write Num
    End If
    Response.Write " size='5' maxlength='3'>"
    Response.Write "      <font color='#FF0000'>���Ϊ0������ʾ����" & ChannelShortName & "��</font></td>"
    Response.Write "    </tr>"
    If ModuleType = 5 Then
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong> ��Ʒ���ͣ�</strong></td>"
    Response.Write "      <td height='25'><select name='ProductType' id='ProductType'>"
    Response.Write "        <option value='1'"
    If Trim(ProductType) = "1" Then Response.Write "selected"
    Response.Write ">����������Ʒ</option>"
    Response.Write "        <option value='2'"
    If Trim(ProductType) = "2" Then Response.Write "selected"
    Response.Write ">�Ǽ���Ʒ</option>"
    Response.Write "        <option value='3'"
    If Trim(ProductType) = "3" Then Response.Write "selected"
    Response.Write ">�ؼ���Ʒ</option>"
    Response.Write "        <option value='4'"
    If Trim(ProductType) = "4" Then Response.Write "selected"
    Response.Write ">������Ʒ</option>"
    Response.Write "        <option value='5'"
    If Trim(ProductType) = "5" Then Response.Write "selected"
    Response.Write ">�������ۺ��Ǽ���Ʒ</option>"
    Response.Write "        <option value='6'"
    If Trim(ProductType) = "6" Then Response.Write "selected"
    Response.Write ">������Ʒ</option>"
    Response.Write "        <option value='0'"
    If Trim(ProductType) = "0" Then Response.Write "selected"
    Response.Write ">������Ʒ</option>"
    Response.Write "        </select> </td>"
    Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "���ԣ�</strong></td>"
    Response.Write "      <td height='25'> <input name='IsHot' type='checkbox' id='IsHot' value='1' "
    If LCase(Trim(IsHot)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">"
    Response.Write "        ����" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp; <input name='IsElite' type='checkbox' id='IsElite' value='1' "
    If LCase(Trim(IsElite)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">"
    Response.Write "        �Ƽ�" & ChannelShortName & " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>�������ѡ������ʾ����" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>���ڷ�Χ��</strong></td>"
    Response.Write "      <td height='25'>ֻ��ʾ���"
    Response.Write "        <input name='DateNum' type='text' id='DateNum' value="
    If Trim(DateNum) = "" Then
        Response.Write "0"
    Else
        Response.Write DateNum
    End If
    Response.Write " size='5' maxlength='3'>"
    Response.Write "        ���ڸ��µ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>&nbsp;&nbsp;���Ϊ�ջ�0������ʾ����������" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>���򷽷���</strong></td>"
    Response.Write "      <td height='25'><select name='OrderType' id='OrderType'>"
    Response.Write "       <option value='1' "
    If Trim(OrderType) = "1" Then Response.Write "selected"
    Response.Write ">" & ChannelShortName & "ID������</option>"
    Response.Write "       <option value='2' "
    If Trim(OrderType) = "2" Then Response.Write "selected"
    Response.Write ">" & ChannelShortName & "ID������</option>"
    Response.Write "       <option value='3' "
    If Trim(OrderType) = "3" Then Response.Write "selected"
    Response.Write ">����ʱ�䣨����</option>"
    Response.Write "       <option value='4' "
    If Trim(OrderType) = "4" Then Response.Write "selected"
    Response.Write ">����ʱ�䣨����</option>"
    Response.Write "       <option value='5' "
    If Trim(OrderType) = "5" Then Response.Write "selected"
    Response.Write ">�������������</option>"
    Response.Write "       <option value='6' "
    If Trim(OrderType) = "6" Then Response.Write "selected"
    Response.Write ">�������������</option>"
    Response.Write "      </select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>��ʾ��ʽ��</strong></td>"
    Response.Write "      <td height='25'><select name='ShowType' id='ShowType'>"
    If ModuleType = 5 Then
        Response.Write "        <option value='1' "
        If Trim(ShowType) = "1" Then Response.Write "selected"
        Response.Write " >ͼƬ+����+�۸�+��ť����������</option>"
        Response.Write "        <option value='2' "
        If Trim(ShowType) = "2" Then Response.Write "selected"
        Response.Write " >��ͼƬ+���ƣ��������У�+������+�۸�+��ť��</option>"
        Response.Write "        <option value='3' "
        If Trim(ShowType) = "3" Then Response.Write "selected"
        Response.Write " >ͼƬ+������+�۸�+��ť���������У�����������</option>"
        Response.Write "        <option value='4' "
        If Trim(ShowType) = "4" Then Response.Write "selected"
        Response.Write " >ͼƬ+����+�۸���������</option>"
        Response.Write "        <option value='5' "
        If Trim(ShowType) = "5" Then Response.Write "selected"
        Response.Write " >��ͼƬ+���ƣ��������У�+�۸���������</option>"
        Response.Write "        <option value='6' "
        If Trim(ShowType) = "6" Then Response.Write "selected"
        Response.Write " >ͼƬ+������+�۸��������У�����������</option>"
        Response.Write "        <option value='7' "
        If Trim(ShowType) = "7" Then Response.Write "selected"
        Response.Write " >ͼƬ+����+��ť����������</option>"
        Response.Write "        <option value='8' "
        If Trim(ShowType) = "8" Then Response.Write "selected"
        Response.Write " >ͼƬ+���ƣ���������</option>"
        Response.Write "        <option value='9' "
        If Trim(ShowType) = "9" Then Response.Write "selected"
        Response.Write " >ͼƬ+��ť����������</option>"
        Response.Write "        <option value='10' "
        If Trim(ShowType) = "10" Then Response.Write "selected"
        Response.Write " >ֻ��ʾͼƬ</option>"
        Response.Write "        <option value='11' "
        If Trim(ShowType) = "11" Then Response.Write "selected"
        Response.Write " >���DIV��ʽ</option>"
    Else
        Response.Write "        <option value='1' "
        If Trim(ShowType) = "1" Then Response.Write "selected"
        Response.Write " >ͼƬ+����+���ݼ�飺��������</option>"
        Response.Write "        <option value='2' "
        If Trim(ShowType) = "2" Then Response.Write "selected"
        Response.Write " >��ͼƬ+���⣺�������У�+���ݼ�飺��������</option>"
        Response.Write "        <option value='3' "
        If Trim(ShowType) = "3" Then Response.Write "selected"
        Response.Write " >ͼƬ+������+���ݼ�飺�������У�����������</option>"
        Response.Write "        <option value='4' "
        If Trim(ShowType) = "4" Then Response.Write "selected"
        Response.Write " >���DIV��ʽ</option>"
        Response.Write "        <option value='5' "
        If Trim(ShowType) = "5" Then Response.Write "selected"
        Response.Write " >���RSS��ʽ</option>"
    End If
    Response.Write "        </select>"
    Response.Write "     </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><b>��ҳͼƬ���ã�</b></td>"
    Response.Write "      <td height='25'>&nbsp;��ȣ�"
    Response.Write "        <input name='ImgWidth' type='text' id='ImgWidth' value="
    If Trim(ImgWidth) = "" Then
        Response.Write "130"
    Else
        Response.Write ImgWidth
    End If
    Response.Write " size='5' maxlength='3'>"
    Response.Write "        ����&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "  �߶ȣ� <input name='ImgHeight' type='text' id='ImgHeight' value="
    If Trim(ImgHeight) = "" Then
        Response.Write "90"
    Else
        Response.Write ImgHeight
    End If
    Response.Write "  size='5' maxlength='3'>"
    Response.Write "        ����</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>��������ַ�����</strong></td>"
    Response.Write "      <td height='25'><input name='TitleLen' type='text' id='TitleLen' value="
    If Trim(TitleLen) = "" Then
        Response.Write "30"
    Else
        Response.Write TitleLen
    End If
    Response.Write "   size='5' maxlength='3'>"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>��Ϊ0������ʾ���⣻��Ϊ-1������ʾ�������⡣��ĸ��һ���ַ��������������ַ���</font></td>"
    Response.Write "    </tr>"
    If ModuleType <> 5 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "�����ַ�����</strong></td>"
        Response.Write "      <td height='25'><input name='ContentLen' type='text' id='ContentLen' value="
        If Trim(ContentLen) = "" Then
            Response.Write "0"
        Else
            Response.Write ContentLen
        End If
        Response.Write "  size='5' maxlength='3'>"
        Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>�������0������ʾָ��������" & ChannelShortName & "����</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>��ʾ���ݣ�</strong></td>"
        Response.Write "      <td height='25'><input name='ShowTips' type='checkbox' id='ShowTips' value='1' "
        If LCase(Trim(ShowTips)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">"
        Response.Write "      ��ʾ���ߡ�����ʱ�䡢���������ʾ��Ϣ</td>"
        Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>ÿ����ʾ" & ChannelShortName & "����</strong></td>"
    Response.Write "      <td height='25'><select name='Cols' id='Cols'>"
    Response.Write "      <option value='1' "
    If Trim(Cols) = "1" Then Response.Write "selected"
    Response.Write ">1</option>"
    Response.Write "      <option value='2' "
    If Trim(Cols) = "2" Then Response.Write "selected"
    Response.Write ">2</option>"
    Response.Write "      <option value='3' "
    If Trim(Cols) = "3" Then Response.Write "selected"
    Response.Write ">3</option>"
    Response.Write "      <option value='4' "
    If Trim(Cols) = "4" Then Response.Write "selected"
    Response.Write ">4</option>"
    Response.Write "      <option value='5' "
    If Trim(Cols) = "5" Then Response.Write "selected"
    Response.Write ">5</option>"
    Response.Write "      <option value='6' "
    If Trim(Cols) = "6" Then Response.Write "selected"
    Response.Write ">6</option>"
    Response.Write "      <option value='7' "
    If Trim(Cols) = "7" Then Response.Write "selected"
    Response.Write ">7</option>"
    Response.Write "      <option value='8' "
    If Trim(Cols) = "8" Then Response.Write "selected"
    Response.Write ">8</option>"
    Response.Write "      <option value='9' "
    If Trim(Cols) = "9" Then Response.Write "selected"
    Response.Write ">9</option>"
    Response.Write "      <option value='10' "
    If Trim(Cols) = "10" Then Response.Write "selected"
    Response.Write ">10</option>"
    Response.Write "      <option value='11' "
    If Trim(Cols) = "11" Then Response.Write "selected"
    Response.Write ">11</option>"
    Response.Write "      <option value='12' "
    If Trim(Cols) = "12" Then Response.Write "selected"
    Response.Write ">12</option>"
    Response.Write "      </select>"
    Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;����ָ�������ͻỻ��</td>"
    Response.Write "    </tr>"
    If ModuleType = 5 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>��ʾ�۸�ʽ��</strong></td>"
        Response.Write "      <td height='25'><select name='ShowPriceType' id='ShowPriceType'>"
        Response.Write "      <option value='0' "
        If Trim(ShowPriceType) = "0" Then Response.Write "selected"
        Response.Write ">�Զ���ʾ</option>"
        Response.Write "      <option value='1' "
        If Trim(ShowPriceType) = "1" Then Response.Write "selected"
        Response.Write ">ֻ��ʾԭ��</option>"
        Response.Write "      <option value='2' "
        If Trim(ShowPriceType) = "2" Then Response.Write "selected"
        Response.Write ">ֻ��ʾ��ǰ��</option>"
        Response.Write "      <option value='3' "
        If Trim(ShowPriceType) = "3" Then Response.Write "selected"
        Response.Write ">ֻ��ʾ�г�����ԭ��</option>"
        Response.Write "      <option value='4' "
        If Trim(ShowPriceType) = "4" Then Response.Write "selected"
        Response.Write ">ֻ��ʾ�г����뵱ǰ��</option>"
        Response.Write "      <option value='5' "
        If Trim(ShowPriceType) = "5" Then Response.Write "selected"
        Response.Write ">ֻ��ʾԭ���뵱ǰ��</option>"
        Response.Write "      <option value='6' "
        If Trim(ShowPriceType) = "6" Then Response.Write "selected"
        Response.Write ">ֻ��ʾԭ�����Ա��</option>"
        Response.Write "      <option value='7' "
        If Trim(ShowPriceType) = "7" Then Response.Write "selected"
        Response.Write ">��ʾ�г��ۡ�ԭ�ۺ͵�ǰ��</option>"
        Response.Write "      <option value='8' "
        If Trim(ShowPriceType) = "8" Then Response.Write "selected"
        Response.Write ">��ʾ�г��ۡ�ԭ�ۺͻ�Ա��</option>"
        Response.Write "      <option value='9' "
        If Trim(ShowPriceType) = "9" Then Response.Write "selected"
        Response.Write ">��ʾ�г��ۡ���ǰ�ۺͻ�Ա��</option>"
        Response.Write "      <option value='10' "
        If Trim(ShowPriceType) = "10" Then Response.Write "selected"
        Response.Write ">��ʾ�г��ۡ�ԭ�ۡ���ǰ�ۺͻ�Ա��</option>"
        Response.Write "      </select>"
        Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>ֻ�е�ShowType������Ϊ���۸�ʽʱ����Ч</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>�Ƿ���ʾ�ۿ��ʣ�</strong></td>"
        Response.Write "          <td><input name='ShowDiscount' type='checkbox' id='ShowDiscount' value='1' "
        If LCase(Trim(ShowDiscount)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>ֻ�е�ShowType������Ϊ���۸�ʽʱ����Ч</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>��ť��ʾ��ʽ��</strong></td>"
        Response.Write "      <td height='25'><select name='ShowButtonType' id='ShowButtonType'>"
        Response.Write "       <option value='0' "
        If Trim(ShowButtonType) = "0" Then Response.Write "selected"
        Response.Write ">����ʾ</option>"
        Response.Write "       <option value='1' "
        If Trim(ShowButtonType) = "1" Then Response.Write "selected"
        Response.Write ">��ʾ����ť</option>"
        Response.Write "       <option value='2' "
        If Trim(ShowButtonType) = "2" Then Response.Write "selected"
        Response.Write ">��ʾ��ϸ��ť</option>"
        Response.Write "       <option value='3' "
        If Trim(ShowButtonType) = "3" Then Response.Write "selected"
        Response.Write ">��ʾ�ղذ�ť</option>"
        Response.Write "       <option value='4' "
        If Trim(ShowButtonType) = "4" Then Response.Write "selected"
        Response.Write ">��ʾ������ϸ��ť</option>"
        Response.Write "       <option value='5' "
        If Trim(ShowButtonType) = "5" Then Response.Write "selected"
        Response.Write ">��ʾ�����ղذ�ť</option>"
        Response.Write "       <option value='6' "
        If Trim(ShowButtonType) = "6" Then Response.Write "selected"
        Response.Write ">��ϸ���ղذ�ť</option>"
        Response.Write "       <option value='7' "
        If Trim(ShowButtonType) = "7" Then Response.Write "selected"
        Response.Write ">��������ʾ</option>"
        Response.Write "      </select></td>"
        Response.Write "    </tr>"
        Response.Write "   <tr class='tdbg'>"
        Response.Write "     <td height='25' align='right' class='tdbg5'><strong>��ť��ʽ��</strong></td>"
        Response.Write "     <td height='25' ><input name='ButtonStyle' type='text' value='"
        If Trim(ButtonStyle) = """" Then
            Response.Write ""
        Else
            Response.Write ButtonStyle
        End If
        Response.Write "'  size='10' maxlength='20'></td>"
        Response.Write "   </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "�򿪷�ʽ��</strong></td>"
        Response.Write "      <td height='25'>"
        Response.Write "        <select name='OpenType' id='OpenType'>"
        Response.Write "          <option value='0' "
        If Trim(OpenType) = "0" Then Response.Write "selected"
        Response.Write ">��ԭ���ڴ�</option>"
        Response.Write "          <option value='1' "
        If Trim(OpenType) = "1" Then Response.Write "selected"
        Response.Write ">���´��ڴ�</option>"
        Response.Write "        </select>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
    End If
End Sub

Sub GetSlide()

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "��Ŀ��</strong></td>"
    Response.Write "      <td height='25'><input name='Num' type='text' value="
    If Trim(Num) = "" Then
        Response.Write "4"
    Else
        Response.Write Num
    End If
    Response.Write " size='5' maxlength='3'>"
    Response.Write "      <font color='#FF0000'>���Ϊ0������ʾ����" & ChannelShortName & "��</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "���ԣ�</strong></td>"
    Response.Write "      <td height='25'> <input name='IsHot' type='checkbox' id='IsHot' value='1' "
    If LCase(Trim(IsHot)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">"
    Response.Write "        ����" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp; <input name='IsElite' type='checkbox' id='IsElite' value='1' "
    If LCase(Trim(IsElite)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">"
    Response.Write "        �Ƽ�" & ChannelShortName & " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>�������ѡ������ʾ����" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>���ڷ�Χ��</strong></td>"
    Response.Write "      <td height='25'>ֻ��ʾ���"
    Response.Write "        <input name='DateNum' type='text' id='DateNum' value="
    If Trim(DateNum) = "" Then
        Response.Write "0"
    Else
        Response.Write DateNum
    End If
    Response.Write " size='5' maxlength='3'>"
    Response.Write "        ���ڸ��µ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>&nbsp;&nbsp;���Ϊ�ջ�0������ʾ����������" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>���򷽷���</strong></td>"
    Response.Write "      <td height='25'><select name='OrderType' id='OrderType'>"
    Response.Write "       <option value='1' "
    If Trim(OrderType) = "1" Then Response.Write "selected"
    Response.Write ">" & ChannelShortName & "ID������</option>"
    Response.Write "       <option value='2' "
    If Trim(OrderType) = "2" Then Response.Write "selected"
    Response.Write ">" & ChannelShortName & "ID������</option>"
    Response.Write "       <option value='3' "
    If Trim(OrderType) = "3" Then Response.Write "selected"
    Response.Write ">����ʱ�䣨����</option>"
    Response.Write "       <option value='4' "
    If Trim(OrderType) = "4" Then Response.Write "selected"
    Response.Write ">����ʱ�䣨����</option>"
    Response.Write "       <option value='5' "
    If Trim(OrderType) = "5" Then Response.Write "selected"
    Response.Write ">�������������</option>"
    Response.Write "       <option value='6' "
    If Trim(OrderType) = "6" Then Response.Write "selected"
    Response.Write ">�������������</option>"
    Response.Write "      </select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><b>��ҳͼƬ���ã�</b></td>"
    Response.Write "      <td height='25'>&nbsp;��ȣ�"
    Response.Write "        <input name='ImgWidth' type='text' id='ImgWidth' value="
    If Trim(ImgWidth) = "" Then
        Response.Write "130"
    Else
        Response.Write ImgWidth
    End If
    Response.Write " size='5' maxlength='3'>"
    Response.Write "        ����&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "  �߶ȣ� <input name='ImgHeight' type='text' id='ImgHeight' value="
    If Trim(ImgHeight) = "" Then
        Response.Write "90"
    Else
        Response.Write ImgHeight
    End If
    Response.Write "  size='5' maxlength='3'>"
    Response.Write "        ����</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><b>����/���Ƴ���</b></td>"
    Response.Write "      <td height='25'> <input name='TitleLen' type='text' id='TitleLen' value="
    If Trim(TitleLen) = "" Then
        Response.Write "30"
    Else
        Response.Write TitleLen
    End If
    Response.Write "  size='5' maxlength='3'> ���ַ�</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><b>Ч���任���ʱ��</b></td>"
    
    Response.Write "      <td height='25'> <input name='iTimeOut' type='text' id='iTimeOut' value="
    If Trim(iTimeOut) = "" Then
        Response.Write "5000"
    Else
        Response.Write iTimeOut
    End If
    Response.Write "  size='5' maxlength='5'>&nbsp;&nbsp;<font color=blue><b>����Ϊ��λ</b></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><b>ͼƬת��Ч��</b></td>"
    Response.Write "      <td height='25'> <input name='effectID' type='text' id='effectID' value="
    If Trim(effectID) = "" Then
        Response.Write "-1"
    Else
        Response.Write effectID
    End If
    Response.Write "  size='5' maxlength='3'>&nbsp;&nbsp;<font color=blue><b>-1��ʾ���Ч����0��23ָ��ĳһ����Ч</b></td>"
    Response.Write "    </tr>"
    'Response.Write "    <tr class='tdbg'>"
    'Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "�򿪷�ʽ��</strong></td>"
    'Response.Write "      <td height='25'>"
    'Response.Write "        <select name='OpenType' id='OpenType'>"
    'Response.Write "          <option value='0' "
    'If Trim(OpenType) = "0" Then Response.Write "selected"
    'Response.Write ">��ԭ���ڴ�</option>"
    'Response.Write "          <option value='1' "
    'If Trim(OpenType) = "1" Then Response.Write "selected"
    'Response.Write ">���´��ڴ�</option>"
    'Response.Write "        </select>"
    'Response.Write "      </td>"
    'Response.Write "    </tr>"
End Sub

Sub GetPositionList()
    Response.Write "    <tr class=tdbg>"
    Response.Write "      <td align=left height=25>��ʾְλ����</td>"
    Response.Write "      <td colspan='1'><input name='PositionNum'  type='text' size='12' value='"
    If Trim(PositionNum) = "" Then
        Response.Write "0"
    Else
        Response.Write PositionNum
    End If
    Response.Write "'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class=tdbg>"
    Response.Write "       <td align=left height=25>�Ƿ������Ƹ��</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Select id=IsUrgent name=IsUrgent>"
    Response.Write "             <Option value='True'"
    If Trim(IsUrgent) = "True" Then Response.Write "selected"
    Response.Write ">������Ƹ</Option>"
    Response.Write "             <Option value='False'"
    If Trim(IsUrgent) = "False" Then Response.Write "selected"
    Response.Write ">������Ƹ</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>���ڷ�Χ��</td>"
    Response.Write "       <td><input name='DateNum'  type='text' size='12' value='"
    If Trim(DateNum) = "" Then
        Response.Write "0"
    Else
        Response.Write DateNum
    End If
    Response.Write "'>"
    Response.Write "       &nbsp;&nbsp;&nbsp;<font color='red'>�������0����ֻ��ʾ��������ڸ��µ�ְλ</font></td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>����ʽ</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Select id=OrderType name=OrderType>"
    Response.Write "             <Option value='1'"
    If Trim(OrderType) = "1" Then Response.Write "selected"
    Response.Write ">��ְλID����</Option>"
    Response.Write "             <Option value='2'"
    If Trim(OrderType) = "2" Then Response.Write "selected"
    Response.Write ">��ְλID����</Option>"
    Response.Write "             <Option value='3'"
    If Trim(OrderType) = "3" Then Response.Write "selected"
    Response.Write ">������ʱ�併��</Option>"
    Response.Write "             <Option value='4'"
    If Trim(OrderType) = "4" Then Response.Write "selected"
    Response.Write ">������ʱ������</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>ְλ��ʾ��ʽ:</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "            <Select id=ShowType name=ShowType>"
    Response.Write "               <Option value='1'"
    If Trim(ShowType) = "1" Then Response.Write "selected"
    Response.Write ">������Ƹ��ʽ</Option>"
    Response.Write "               <Option value='2'"
    If Trim(ShowType) = "2" Then Response.Write "selected"
    Response.Write ">������Ƹ��ʽ</Option>"
    Response.Write "               <Option value='3'"
    If Trim(ShowType) = "3" Then Response.Write "selected"
    Response.Write ">ְλ��Ϣ�б�</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>ְλ���Ƴ��ȣ�</td>"
    Response.Write "       <td><input name='TitleLen' type='text' size='12' value='"
    If Trim(TitleLen) = "" Then
        Response.Write "0"
    Else
        Response.Write TitleLen
    End If
    Response.Write "'>"
    Response.Write "       &nbsp;&nbsp;&nbsp;<font color='red'>һ������=����Ӣ���ַ�,��Ϊ0������ʾ����ְλ��</font></td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "      <td align=left height=25>�����ص����Ƴ��ȣ�</td>"
    Response.Write "      <td colspan='1'><input name='WorkPlaceNameLen' type='text' size='12' value='"
    If Trim(WorkPlaceNameLen) = "" Then
        Response.Write "0"
    Else
        Response.Write WorkPlaceNameLen
    End If
    Response.Write "'></td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "      <td align=left height=25>���˵�λ���Ƴ��ȣ�</td>"
    Response.Write "      <td colspan='1'><input name='SubCompanyNameLen' type='text' size='12' value='"
    If Trim(SubCompanyNameLen) = "" Then
        Response.Write "0"
    Else
        Response.Write SubCompanyNameLen
    End If
    Response.Write "'></td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>���ƹ���ʱ����ʾʡ�Ժ����ã�</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "        <Input id='PShowPoints' type='checkbox' value='True' name='PShowPoints' "
    If LCase(Trim(PShowPoints)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">ְλ����"
    Response.Write "         <Input id='WShowPoints' type='checkbox' value='True' name='WShowPoints' "
    If LCase(Trim(WShowPoints)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">�����ص�"
    Response.Write "          <Input id='SShowPoints' type='checkbox' value='True' name='SShowPoints'"
    If LCase(Trim(SShowPoints)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">���˵�λ"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>����������ʾ��ʽ��</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Select id=ShowDateType name=ShowDateType>"
    Response.Write "             <Option value='0'"
    If Trim(ShowDateType) = "0" Then Response.Write "selected"
    Response.Write ">����ʾ</Option>"
    Response.Write "             <Option value='1'"
    If Trim(ShowDateType) = "1" Then Response.Write "selected"
    Response.Write ">��ʾ������</Option>"
    Response.Write "             <Option value='2'"
    If Trim(ShowDateType) = "2" Then Response.Write "selected"
    Response.Write ">��ʾ����</Option>"
    Response.Write "             <Option value='3'"
    If Trim(ShowDateType) = "3" Then Response.Write "selected"
    Response.Write ">��ʾ���գ���-�գ�</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>�Ƿ���ʾ����<br>ְλ��Ϣѡ�</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Input id='ShowPositionID' type='checkbox' value='1' name='ShowPositionID'"
    If Trim(ShowPositionID) = "1" Then Response.Write "checked"
    Response.Write ">��ʾְλID"
    Response.Write "          <Input id='ShowPositionName' type='checkbox' value='1' name='ShowPositionName'"
    If Trim(ShowPositionName) = "1" Then Response.Write "checked"
    Response.Write ">��ʾְλ����"
    Response.Write "          <Input id='ShowWorkPlaceName' type='checkbox' value='1' name='ShowWorkPlaceName'"
    If Trim(ShowWorkPlaceName) = "1" Then Response.Write "checked"
    Response.Write ">��ʾ�����ص�<br>"
    Response.Write "          <Input id='ShowSubCompanyName' type='checkbox' value='1' name='ShowSubCompanyName'"
    If Trim(ShowSubCompanyName) = "1" Then Response.Write "checked"
    Response.Write ">��ʾ���˵�λ"
    Response.Write "          <Input id='ShowPositionNum' type='checkbox' value='1' name='ShowPositionNum'"
    If Trim(ShowPositionNum) = "1" Then Response.Write "checked"
    Response.Write ">��ʾ��Ƹ����"
    Response.Write "          <Input id='ShowPositionStatus' type='checkbox' value='1' name='ShowPositionStatus' "
    If Trim(ShowPositionStatus) = "1" Then Response.Write "checked"
    Response.Write ">��ʾְλ״̬<br>"
    Response.Write "          <Input id='ShowValidDate' type='checkbox' value='1' name='ShowValidDate' "
    If Trim(ShowValidDate) = "1" Then Response.Write "checked"
    Response.Write ">��ʾ��Ч��"
    Response.Write "          <Input id='ShowUrgentSign' type='checkbox' value='True' name='ShowUrgentSign'"
    If Trim(ShowUrgentSign) = "True" Then Response.Write "checked"
    Response.Write ">��ʾ������Ƹ��־"
    Response.Write "          <Input id='ShowNewSign' type='checkbox' value='True' name='ShowNewSign'"
    If Trim(ShowNewSign) = "True" Then Response.Write "checked"
    Response.Write ">��ʾ����Ƹ��־"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>�Ƿ��ҳ��ʾ��</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "           <Select id=UsePage name=UsePage>"
    Response.Write "              <Option value='True'"
    If Trim(UsePage) = "True" Then Response.Write "selected"
    Response.Write ">��ҳ��ʾ</Option>"
    Response.Write "              <Option value='False'"
    If Trim(UsePage) = "False" Then Response.Write "selected"
    Response.Write ">����ҳ��ʾ</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>����ְλҳ�򿪷�ʽ��</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Select id=OpenType name=OpenType>"
    Response.Write "             <Option value='0'"
    If Trim(OpenType) = "0" Then Response.Write "selected"
    Response.Write ">ԭ���ڴ�</Option>"
    Response.Write "             <Option value='1'"
    If Trim(OpenType) = "1" Then Response.Write "selected"
    Response.Write ">�´��ڴ�</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
End Sub

Sub GetSearchResult()
    Response.Write "    <tr class=tdbg>"
    Response.Write "      <td align=left height=25>��ʾ��¼����</td>"
    Response.Write "      <td colspan='1'><input name='ShowNum'  type='text' size='12' value='"
    If Trim(ShowNum) = "" Then
        Response.Write "0"
    Else
        Response.Write ShowNum
    End If
    Response.Write "'></td>"
    Response.Write "    </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>����ʽ</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Select id=OrderType name=OrderType>"
    Response.Write "             <Option value='1'"
    If Trim(OrderType) = "1" Then Response.Write "selected"
    Response.Write ">��ְλID����</Option>"
    Response.Write "             <Option value='2'"
    If Trim(OrderType) = "2" Then Response.Write "selected"
    Response.Write ">��ְλID����</Option>"
    Response.Write "             <Option value='3'"
    If Trim(OrderType) = "3" Then Response.Write "selected"
    Response.Write ">������ʱ�併��</Option>"
    Response.Write "             <Option value='4'"
    If Trim(OrderType) = "4" Then Response.Write "selected"
    Response.Write ">������ʱ������</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>ְλ���Ƴ��ȣ�</td>"
    Response.Write "       <td><input name='TitleLen' type='text' size='12' value='"
    If Trim(TitleLen) = "" Then
        Response.Write "0"
    Else
        Response.Write TitleLen
    End If
    Response.Write "'>"
    Response.Write "       &nbsp;&nbsp;&nbsp;<font color='red'>һ������=����Ӣ���ַ�,��Ϊ0������ʾ����ְλ��</font></td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "      <td align=left height=25>�����ص����Ƴ��ȣ�</td>"
    Response.Write "      <td colspan='1'><input name='WorkPlaceNameLen' type='text' size='12' value='"
    If Trim(WorkPlaceNameLen) = "" Then
        Response.Write "0"
    Else
        Response.Write WorkPlaceNameLen
    End If
    Response.Write "'></td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "      <td align=left height=25>���˵�λ���Ƴ��ȣ�</td>"
    Response.Write "      <td colspan='1'><input name='SubCompanyNameLen' type='text' size='12' value='"
    If Trim(SubCompanyNameLen) = "" Then
        Response.Write "0"
    Else
        Response.Write SubCompanyNameLen
    End If
    Response.Write "'></td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>���ƹ���ʱ����ʾʡ�Ժ����ã�</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "        <Input id='PShowPoints' type='checkbox' value='True' name='PShowPoints' "
    If LCase(Trim(PShowPoints)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">ְλ����"
    Response.Write "         <Input id='WShowPoints' type='checkbox' value='True' name='WShowPoints' "
    If LCase(Trim(WShowPoints)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">�����ص�"
    Response.Write "          <Input id='SShowPoints' type='checkbox' value='True' name='SShowPoints'"
    If LCase(Trim(SShowPoints)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">���˵�λ"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>����������ʾ��ʽ��</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Select id=ShowDateType name=ShowDateType>"
    Response.Write "             <Option value='0'"
    If Trim(ShowDateType) = "0" Then Response.Write "selected"
    Response.Write ">����ʾ</Option>"
    Response.Write "             <Option value='1'"
    If Trim(ShowDateType) = "1" Then Response.Write "selected"
    Response.Write ">��ʾ������</Option>"
    Response.Write "             <Option value='2'"
    If Trim(ShowDateType) = "2" Then Response.Write "selected"
    Response.Write ">��ʾ����</Option>"
    Response.Write "             <Option value='3'"
    If Trim(ShowDateType) = "3" Then Response.Write "selected"
    Response.Write ">��ʾ���գ���-�գ�</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>�Ƿ���ʾ����<br>ְλ��Ϣѡ�</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Input id='ShowPositionID' type='checkbox' value='1' name='ShowPositionID'"
    If Trim(ShowPositionID) = "1" Then Response.Write "checked"
    Response.Write ">��ʾְλID"
    Response.Write "          <Input id='ShowPositionName' type='checkbox' value='1' name='ShowPositionName'"
    If Trim(ShowPositionName) = "1" Then Response.Write "checked"
    Response.Write ">��ʾְλ����"
    Response.Write "          <Input id='ShowWorkPlaceName' type='checkbox' value='1' name='ShowWorkPlaceName'"
    If Trim(ShowWorkPlaceName) = "1" Then Response.Write "checked"
    Response.Write ">��ʾ�����ص�<br>"
    Response.Write "          <Input id='ShowSubCompanyName' type='checkbox' value='1' name='ShowSubCompanyName'"
    If Trim(ShowSubCompanyName) = "1" Then Response.Write "checked"
    Response.Write ">��ʾ���˵�λ"
    Response.Write "          <Input id='ShowPositionNum' type='checkbox' value='1' name='ShowPositionNum'"
    If Trim(ShowPositionNum) = "1" Then Response.Write "checked"
    Response.Write ">��ʾ��Ƹ����"
    Response.Write "          <Input id='ShowPositionStatus' type='checkbox' value='1' name='ShowPositionStatus' "
    If Trim(ShowPositionStatus) = "1" Then Response.Write "checked"
    Response.Write ">��ʾְλ״̬<br>"
    Response.Write "          <Input id='ShowValidDate' type='checkbox' value='1' name='ShowValidDate' "
    If Trim(ShowValidDate) = "1" Then Response.Write "checked"
    Response.Write ">��ʾ��Ч��"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>�Ƿ��ҳ��ʾ��</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "           <Select id=UsePage name=UsePage>"
    Response.Write "              <Option value='True'"
    If Trim(UsePage) = "True" Then Response.Write "selected"
    Response.Write ">��ҳ��ʾ</Option>"
    Response.Write "              <Option value='False'"
    If Trim(UsePage) = "False" Then Response.Write "selected"
    Response.Write ">����ҳ��ʾ</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>����ְλҳ�򿪷�ʽ��</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Select id=OpenType name=OpenType>"
    Response.Write "             <Option value='0'"
    If Trim(OpenType) = "0" Then Response.Write "selected"
    Response.Write ">ԭ���ڴ�</Option>"
    Response.Write "             <Option value='1'"
    If Trim(OpenType) = "1" Then Response.Write "selected"
    Response.Write ">�´��ڴ�</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
End Sub

Function GetSpecial_Option(iChannelID, SpecialID)
    Dim sqlSpecial, rsSpecial, strOption, strOptionTemp
	If Instr(iChannelID,",")>0 and IsValidID(iChannelID) = True Then
        sqlSpecial = "select ChannelID,SpecialID,SpecialName,OrderID from PE_Special where ChannelID=0 or ChannelID in (" & iChannelID & ")   order by ChannelID,OrderID"
	Else
        sqlSpecial = "select ChannelID,SpecialID,SpecialName,OrderID from PE_Special where ChannelID=0 or ChannelID=" & PE_Clng(iChannelID) & "   order by ChannelID,OrderID"	
	End If
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
        If FoundInArr(ChannelID,rsChannel(0),",")=True Or FoundInArr(ChannelID,rsChannel(0),"|")=True  Then
            strChannel = strChannel & "<option value='" & rsChannel(0) & "' selected>" & rsChannel(1) & "</option>"
        Else
            strChannel = strChannel & "<option value='" & rsChannel(0) & "' >" & rsChannel(1) & "</option>"
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
	ChannelID = Replace(ChannelID,"|",",")	
    Dim arrShowLine(20)
    For i = 0 To UBound(arrShowLine)
    arrShowLine(i) = False
    Next
    If InStr(ChannelID, ",") > 0 and IsValidID(ChannelID) = True Then
	    sqlClass = "Select * from PE_Class where ChannelID in (" & ChannelID & ") order by RootID,OrderID"		
    Else
        sqlClass = "Select * from PE_Class where ChannelID=" & PE_CLng(ChannelID) & " order by RootID,OrderID"
	End If
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

    If Trim(ClassID) = "-1" Then
        strClass_Option = strClass_Option & "<option value='-1' " & classcss & " selected>δָ���κ���Ŀ</option>"
    Else
        strClass_Option = strClass_Option & "<option value='-1' " & classcss & ">δָ���κ���Ŀ</option>"
    End If

    GetClass_Channel = strClass_Option
End Function

Sub GetLabelData()
    editLabel = PE_HtmlDecode(Trim(request.querystring("editLabel")))
    If InStr(editLabel, "{$") = 0 Then
        Response.Write "<center><br><font color=red>��ѡ��Ĳ��Ǳ�ǩ</font></center>"
        Response.End
    End If

    Dim editLabeltemp
    editLabeltemp = Trim(Replace(Replace(editLabel, "{$", ""), "}", ""))
    editLabeltemp = Replace(editLabeltemp, """", "")
    LabelName = Left(editLabeltemp, InStr(Trim(Replace(Replace(editLabeltemp, "{$", ""), "}", "")), "(") - 1)
    editLabeltemp = Trim(Replace(Replace(Replace(editLabeltemp, "(", ""), ")", ""), LabelName, ""))
    arrParameter = Split(editLabeltemp, ",")
    Select Case LabelName
    Case "GetArticleList"
        ChannelShortName = "����"
        ChannelShowType = "GetList"
        imageproperty = "article"
        ModuleType = 1
        ChannelID = arrParameter(0)
        If InStr(arrParameter(0), "|") > 0 Then
            NChannelID = True
        Else
            NChannelID = False
        End If		
        ClassID = arrParameter(1)
        If InStr(arrParameter(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = arrParameter(2)
        SpecialID = arrParameter(3)
        urltype = "0"
        Num = arrParameter(5)
        IsHot = arrParameter(6)
        IsElite = arrParameter(7)
        AuthorName = arrParameter(8)
        DateNum = arrParameter(9)
        OrderType = arrParameter(10)
        ShowType = arrParameter(11)
        TitleLen = arrParameter(12)
        ContentLen = arrParameter(13)
        ShowClassName = arrParameter(14)
        ShowPropertyType = arrParameter(15)
        ShowIncludePic = arrParameter(16)
        ShowAuthor = arrParameter(17)
        ShowDateType = arrParameter(18)
        ShowHits = arrParameter(19)
        ShowHotSign = arrParameter(20)
        ShowNewSign = arrParameter(21)
        ShowTips = arrParameter(22)
        ShowCommentLink = arrParameter(23)
        UsePage = arrParameter(24)
        OpenType = arrParameter(25)
        If UBound(arrParameter) = 26 Then
            Cols = arrParameter(26)
        End If
        If UBound(arrParameter) >= 29 Then
            Cols = arrParameter(26)
            CssNameA = arrParameter(27)
            CssName1 = arrParameter(28)
            CssName2 = arrParameter(29)
        End If

        If UBound(arrParameter) >= 30 Then
            IntervalLines = arrParameter(30)
        End If
     Case "GetPicArticle"
        ChannelShortName = "����"
        imageproperty = "article"
        ChannelID = arrParameter(0)
        If InStr(arrParameter(0), "|") > 0 Then
            NChannelID = True
        Else
            NChannelID = False
        End If			
        ClassID = arrParameter(1)
        If InStr(arrParameter(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = arrParameter(2)
        SpecialID = arrParameter(3)
        Num = arrParameter(4)
        IsHot = arrParameter(5)
        IsElite = arrParameter(6)
        DateNum = arrParameter(7)
        OrderType = arrParameter(8)
        ShowType = arrParameter(9)
        ImgWidth = arrParameter(10)
        ImgHeight = arrParameter(11)
        TitleLen = arrParameter(12)
        ContentLen = arrParameter(13)
        ShowTips = arrParameter(14)
        Cols = arrParameter(15)
        ChannelShowType = "GetPic"
        ModuleType = 1
     Case "GetSlidePicArticle"
        ChannelShortName = "����"
        imageproperty = "article"
        ChannelID = arrParameter(0)
        If InStr(arrParameter(0), "|") > 0 Then
            NChannelID = True
        Else
            NChannelID = False
        End If			
        ClassID = arrParameter(1)
        If InStr(arrParameter(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = arrParameter(2)
        SpecialID = arrParameter(3)
        Num = arrParameter(4)
        IsHot = arrParameter(5)
        IsElite = arrParameter(6)
        DateNum = arrParameter(7)
        OrderType = arrParameter(8)
        ImgWidth = arrParameter(9)
        ImgHeight = arrParameter(10)
        TitleLen = arrParameter(11)
        iTimeOut = arrParameter(12)
        effectID = arrParameter(13)
        ChannelShowType = "GetSlide"
        ModuleType = 1
     Case "GetSoftList"
        ChannelShortName = "���"
        imageproperty = "Soft"
        ChannelID = arrParameter(0)
        ClassID = arrParameter(1)
        If InStr(arrParameter(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = arrParameter(2)
        SpecialID = arrParameter(3)
        urltype = "0"
        Num = arrParameter(5)
        IsHot = arrParameter(6)
        IsElite = arrParameter(7)
        AuthorName = arrParameter(8)
        DateNum = arrParameter(9)
        OrderType = arrParameter(10)
        ShowType = arrParameter(11)
        TitleLen = arrParameter(12)
        ContentLen = arrParameter(13)
        ShowClassName = arrParameter(14)
        ShowPropertyType = arrParameter(15)
        ShowAuthor = arrParameter(16)
        ShowDateType = arrParameter(17)
        ShowHits = arrParameter(18)
        ShowHotSign = arrParameter(19)
        ShowNewSign = arrParameter(20)
        ShowTips = arrParameter(21)
        UsePage = arrParameter(22)
        OpenType = arrParameter(23)
        If UBound(arrParameter) >= 27 Then
            Cols = arrParameter(24)
            CssNameA = arrParameter(25)
            CssName1 = arrParameter(26)
            CssName2 = arrParameter(27)
        End If
        If UBound(arrParameter) >= 28 Then
            IntervalLines = arrParameter(28)
        End If
        ChannelShowType = "GetList"
        ModuleType = 2
     Case "GetPicSoft"
        ChannelShortName = "���"
        imageproperty = "Soft"
        ChannelID = arrParameter(0)
        ClassID = arrParameter(1)
        If InStr(arrParameter(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = arrParameter(2)
        SpecialID = arrParameter(3)
        Num = arrParameter(4)
        IsHot = arrParameter(5)
        IsElite = arrParameter(6)
        DateNum = arrParameter(7)
        OrderType = arrParameter(8)
        ShowType = arrParameter(9)
        ImgWidth = arrParameter(10)
        ImgHeight = arrParameter(11)
        TitleLen = arrParameter(12)
        ContentLen = arrParameter(13)
        ShowTips = arrParameter(14)
        Cols = arrParameter(15)
        ChannelShowType = "GetPic"
        ModuleType = 2
     Case "GetSlidePicSoft"
        ChannelShortName = "���"
        imageproperty = "Soft"
        ChannelID = arrParameter(0)
        ClassID = arrParameter(1)
        If InStr(arrParameter(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = arrParameter(2)
        SpecialID = arrParameter(3)
        Num = arrParameter(4)
        IsHot = arrParameter(5)
        IsElite = arrParameter(6)
        DateNum = arrParameter(7)
        OrderType = arrParameter(8)
        ImgWidth = arrParameter(9)
        ImgHeight = arrParameter(10)
        TitleLen = arrParameter(11)
        iTimeOut = arrParameter(12)
        effectID = arrParameter(13)
        ChannelShowType = "GetSlide"
        ModuleType = 2
     Case "GetPhotoList"
        ChannelShortName = "ͼƬ"
        imageproperty = "Photo"
        ChannelID = arrParameter(0)
        ClassID = arrParameter(1)
        If InStr(arrParameter(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = arrParameter(2)
        SpecialID = arrParameter(3)
        urltype = "0"
        Num = arrParameter(5)
        IsHot = arrParameter(6)
        IsElite = arrParameter(7)
        AuthorName = arrParameter(8)
        DateNum = arrParameter(9)
        OrderType = arrParameter(10)
        ShowType = arrParameter(11)
        TitleLen = arrParameter(12)
        ContentLen = arrParameter(13)
        ShowClassName = arrParameter(14)
        ShowPropertyType = arrParameter(15)
        ShowAuthor = arrParameter(16)
        ShowDateType = arrParameter(17)
        ShowHits = arrParameter(18)
        ShowHotSign = arrParameter(19)
        ShowNewSign = arrParameter(20)
        ShowTips = arrParameter(21)
        UsePage = arrParameter(22)
        OpenType = arrParameter(23)
        If UBound(arrParameter) >= 27 Then
            Cols = arrParameter(24)
            CssNameA = arrParameter(25)
            CssName1 = arrParameter(26)
            CssName2 = arrParameter(27)
        End If
        If UBound(arrParameter) >= 28 Then
            IntervalLines = arrParameter(28)
        End If
        ChannelShowType = "GetList"
        ModuleType = 3
     Case "GetPicPhoto"
        ChannelShortName = "ͼƬ"
        imageproperty = "Photo"
        ChannelID = arrParameter(0)
        ClassID = arrParameter(1)
        If InStr(arrParameter(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = arrParameter(2)
        SpecialID = arrParameter(3)
        Num = arrParameter(4)
        IsHot = arrParameter(5)
        IsElite = arrParameter(6)
        DateNum = arrParameter(7)
        OrderType = arrParameter(8)
        ShowType = arrParameter(9)
        ImgWidth = arrParameter(10)
        ImgHeight = arrParameter(11)
        TitleLen = arrParameter(12)
        ContentLen = arrParameter(13)
        ShowTips = arrParameter(14)
        Cols = arrParameter(15)
        ChannelShowType = "GetPic"
        ModuleType = 3

     Case "GetSlidePicPhoto"
        ChannelShortName = "ͼƬ"
        imageproperty = "Photo"
        ChannelID = arrParameter(0)
        ClassID = arrParameter(1)
        If InStr(arrParameter(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = arrParameter(2)
        SpecialID = arrParameter(3)
        Num = arrParameter(4)
        IsHot = arrParameter(5)
        IsElite = arrParameter(6)
        DateNum = arrParameter(7)
        OrderType = arrParameter(8)
        ImgWidth = arrParameter(9)
        ImgHeight = arrParameter(10)
        TitleLen = arrParameter(11)
        iTimeOut = arrParameter(12)
        effectID = arrParameter(13)
        ChannelShowType = "GetSlide"
        ModuleType = 3
     Case "GetProductList"
        ChannelShortName = "��Ʒ"
        imageproperty = "Product"
        ChannelID = 1000
        ClassID = arrParameter(0)
        If InStr(arrParameter(0), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = arrParameter(1)
        SpecialID = arrParameter(2)
        Num = arrParameter(3)
        ProductType = arrParameter(4)
        IsHot = arrParameter(5)
        IsElite = arrParameter(6)
        DateNum = arrParameter(7)
        OrderType = arrParameter(8)
        ShowType = arrParameter(9)
        TitleLen = arrParameter(10)
        ContentLen = arrParameter(11)
        ShowClassName = arrParameter(12)
        ShowPropertyType = arrParameter(13)
        ShowDateType = arrParameter(14)
        ShowHotSign = arrParameter(15)
        ShowNewSign = arrParameter(16)
        UsePage = arrParameter(17)
        OpenType = arrParameter(18)
        If UBound(arrParameter) >= 39 Then
            IntervalLines = arrParameter(19)
            Cols = arrParameter(20)
            ShowTableTitle = arrParameter(21)
            TableTitleStr = arrParameter(22)
            ShowProductModel = arrParameter(23)
            ShowProductStandard = arrParameter(24)
            ShowUnit = arrParameter(25)
            ShowStocksType = arrParameter(26)
            ShowWeight = arrParameter(27)
            ShowPrice_Market = arrParameter(28)
            ShowPrice_Original = arrParameter(29)
            ShowPrice = arrParameter(30)
            ShowPrice_Member = arrParameter(31)
            ShowDiscount = arrParameter(32)
            ShowButtonType = arrParameter(33)
            ButtonStyle = arrParameter(34)

            CssNameTable = arrParameter(35)
            CssNameTitle = arrParameter(36)
            CssNameA = arrParameter(37)
            CssName1 = arrParameter(38)
            CssName2 = arrParameter(39)
        End If
        urltype = "0"
        ChannelShowType = "GetList"
        ModuleType = 5
    Case "GetPicProduct"
        ChannelShortName = "��Ʒ"
        imageproperty = "Product"
        ChannelID = 1000
        ClassID = arrParameter(0)
        If InStr(arrParameter(0), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = arrParameter(1)
        SpecialID = arrParameter(2)
        Num = arrParameter(3)
        ProductType = arrParameter(4)
        IsHot = arrParameter(5)
        IsElite = arrParameter(6)
        DateNum = arrParameter(7)
        OrderType = arrParameter(8)
        ShowType = arrParameter(9)
        ImgWidth = arrParameter(10)
        ImgHeight = arrParameter(11)
        TitleLen = arrParameter(12)
        Cols = arrParameter(13)
        If UBound(arrParameter) >= 18 Then
            ShowPriceType = arrParameter(14)
            ShowDiscount = arrParameter(15)
            ShowButtonType = arrParameter(16)
            ButtonStyle = arrParameter(17)
            OpenType = arrParameter(18)
        End If
        ChannelShowType = "GetPic"
        ModuleType = 5
    Case "GetSlidePicProduct"
        ChannelID = 1000
        ChannelShortName = "��Ʒ"
        imageproperty = "Product"
        ClassID = arrParameter(0)
        If InStr(arrParameter(0), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = arrParameter(1)
        SpecialID = arrParameter(2)
        Num = arrParameter(3)
        IsHot = arrParameter(4)
        IsElite = arrParameter(5)
        DateNum = arrParameter(6)
        OrderType = arrParameter(7)
        ImgWidth = arrParameter(8)
        ImgHeight = arrParameter(9)
        TitleLen = arrParameter(10)
        iTimeOut = arrParameter(11)
        effectID = arrParameter(12)
        If UBound(arrParameter) >= 13 Then
            OpenType = arrParameter(13)
        End If
        ChannelShowType = "GetSlide"
        ModuleType = 5
    Case "GetPositionList"
        ChannelShortName = "ְλ"
        ChannelShowType = "GetPositionList"
        imageproperty = "Job"
        ModuleType = 8
        PositionNum = arrParameter(0)
        IsUrgent = arrParameter(1)
        DateNum = arrParameter(2)
        OrderType = arrParameter(3)
        ShowType = arrParameter(4)
        TitleLen = arrParameter(5)
        WorkPlaceNameLen = arrParameter(6)
        SubCompanyNameLen = arrParameter(7)
        PShowPoints = arrParameter(8)
        WShowPoints = arrParameter(9)
        SShowPoints = arrParameter(10)
        ShowDateType = arrParameter(11)
        ShowPositionID = arrParameter(12)
        ShowPositionName = arrParameter(13)
        ShowWorkPlaceName = arrParameter(14)
        ShowSubCompanyName = arrParameter(15)
        ShowPositionNum = arrParameter(16)
        ShowPositionStatus = arrParameter(17)
        ShowValidDate = arrParameter(18)
        If arrParameter(4) = 2 Or arrParameter(4) = 3 Then
            ShowUrgentSign = False
        Else
            ShowUrgentSign = arrParameter(19)
        End If
        If arrParameter(4) = 1 Or arrParameter(4) = 3 Then
            ShowNewSign = False
        Else
            ShowNewSign = arrParameter(20)
        End If
        If arrParameter(4) = 1 Or arrParameter(4) = 2 Then
            UsePage = False
        Else
            UsePage = arrParameter(21)
        End If
        OpenType = arrParameter(22)
    Case "GetSearchResult"
        ChannelShortName = "ְλ"
        ChannelShowType = "GetSearchResult"
        imageproperty = "Job"
        ModuleType = 8
        ShowNum = arrParameter(0)
        OrderType = arrParameter(1)
        TitleLen = arrParameter(2)
        WorkPlaceNameLen = arrParameter(3)
        SubCompanyNameLen = arrParameter(4)
        PShowPoints = arrParameter(5)
        WShowPoints = arrParameter(6)
        SShowPoints = arrParameter(7)
        ShowDateType = arrParameter(8)
        ShowPositionID = arrParameter(9)
        ShowPositionName = arrParameter(10)
        ShowWorkPlaceName = arrParameter(11)
        ShowSubCompanyName = arrParameter(12)
        ShowPositionNum = arrParameter(13)
        ShowPositionStatus = arrParameter(14)
        ShowValidDate = arrParameter(15)
        If arrParameter(4) = 1 Or arrParameter(4) = 2 Then
            UsePage = False
        Else
            UsePage = arrParameter(16)
        End If
        OpenType = arrParameter(17)
    Case Else
        Response.Write "<center><br><font color=red>��ѡ��Ĳ��Ǳ�ǩ</font></center>"
        Response.End
    End Select
End Sub

Sub CellNclass()
    Response.Write "    if (document.myform.NClassChild.checked==true){ " & vbCrLf
    Response.Write "        var Nclassidzhi=""""" & vbCrLf
    Response.Write "        for(var i=0;i<document.myform.ClassID.length;i++){" & vbCrLf
    Response.Write "            if (document.myform.ClassID.options[i].selected==true){" & vbCrLf
    Response.Write "                if (document.myform.ClassID.options[i].value==""rsClass_arrChildID""||document.myform.ClassID.options[i].value==""ClassID""||document.myform.ClassID.options[i].value==""arrChildID""||document.myform.ClassID.options[i].value==""0""){" & vbCrLf
    Response.Write "                    alert(""���ڶ�ѡ��ѡ���˺�ɫ���֣���ѡ��Ŀ���ǲ��ܰ����ǲ��ֵġ�"");" & vbCrLf
    Response.Write "                    return false" & vbCrLf
    Response.Write "                }else{" & vbCrLf
    Response.Write "                    if (Nclassidzhi==""""){" & vbCrLf
    Response.Write "                        Nclassidzhi+=document.myform.ClassID.options[i].value;" & vbCrLf
    Response.Write "                    }else{" & vbCrLf
    Response.Write "                        Nclassidzhi+=""|""+document.myform.ClassID.options[i].value;" & vbCrLf
    Response.Write "                    }" & vbCrLf
    Response.Write "                }" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        strJS+=Nclassidzhi;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+=document.myform.ClassID.value;" & vbCrLf
    Response.Write "    }" & vbCrLf
End Sub	
	
Sub CellNchannel()
    Response.Write "    if (document.myform.NChannelID.checked==true){ " & vbCrLf
    Response.Write "        var Nchannelidzhi=""""" & vbCrLf	
    Response.Write "        for(var i=0;i<document.myform.ChannelID.length;i++){" & vbCrLf
    Response.Write "            if (document.myform.ChannelID.options[i].selected==true){" & vbCrLf
    Response.Write "                    if (Nchannelidzhi==""""){" & vbCrLf
    Response.Write "                        Nchannelidzhi+=document.myform.ChannelID.options[i].value;" & vbCrLf
    Response.Write "                    }else{" & vbCrLf
    Response.Write "                        Nchannelidzhi+=""|""+document.myform.ChannelID.options[i].value;" & vbCrLf
    Response.Write "                    }" & vbCrLf
    Response.Write "                }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        strJS+=Nchannelidzhi;" & vbCrLf	
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+=document.myform.ChannelID.value;" & vbCrLf
    Response.Write "    }" & vbCrLf	
	
	
End Sub
%>
