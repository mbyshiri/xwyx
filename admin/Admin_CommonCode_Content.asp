<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Sub ShowTr_Class()
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>������Ŀ��</td>"
    Response.Write "            <td>"
    Response.Write "              <select name='ClassID'>" & GetClass_Option(3, ClassID) & "</select>"
    Response.Write "              &nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'><strong>ע�⣺</strong></font>"
    Response.Write "              <font color='#0000FF'>����ָ��Ϊ�ⲿ��Ŀ</font>"
    If AdminPurview = 2 And AdminPurview_Channel = 3 Then
        Response.Write "<font color='#0000FF'>��������ֻ����<font color='#FF0000'>��ɫ��Ŀ</font>��������Ŀ���������</font>"
    End If
    Response.Write "            </td>"
    Response.Write "          </tr>"
End Sub

Sub ShowTabs_Special(arrSpecialID, strDisabled)
    Response.Write "<SCRIPT language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.SpecialID.length;i++){" & vbCrLf
    Response.Write "    document.myform.SpecialID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.SpecialID.length;i++){" & vbCrLf
    Response.Write "    document.myform.SpecialID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>����ר�⣺</td>"
    Response.Write "            <td>"
    Response.Write "              <select name='SpecialID' size='2' multiple style='height:300px;width:260px;'>" & GetSpecial_Option(arrSpecialID) & "</select>"
    If strDisabled <> " disabled" Then
        Response.Write "              <br><input type='button' name='Submit' value='  ѡ������ר��  ' onclick='SelectAll()'>"
        Response.Write "              <br><input type='button' name='Submit' value='ȡ��ѡ������ר��' onclick='UnSelectAll()'>"
    End If
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf
End Sub


Sub ShowTabs_Property_Add()
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "���ԣ�</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='OnTop' type='checkbox' id='OnTop' value='yes'> �̶�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "              <input name='Hot' type='checkbox' id='Hot' value='yes' onclick=""javascript:document.myform.Hits.value='" & HitsOfHot & "'""> ����" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "              <input name='Elite' type='checkbox' id='Elite' value='yes'> �Ƽ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "              " & ChannelShortName & "���ֵȼ��� <select name='Stars' id='Stars'>" & GetStars(3) & "</select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    If ModuleType = 2 Or ModuleType = 3 Then
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'>"
        If ModuleType = 2 Then
            Response.Write "���ش�����"
        Else
            Response.Write "�鿴������"
        End If
        Response.Write "</td>"
        Response.Write "            <td>"
        Response.Write "              ���գ� <input name='DayHits' type='text' id='DayHits' value='0' size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "              ���ܣ� <input name='WeekHits' type='text' id='WeekHits' value='0' size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "              ���£� <input name='MonthHits' type='text' id='MonthHits' value='0' size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "              �ܼƣ� <input name='Hits' type='text' id='Hits' value='0' size='10' maxlength='10'>"
        Response.Write "            </td>"
        Response.Write "          </tr>"
    Else
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'>�������ʼֵ��</td>"
        Response.Write "            <td>"
        Response.Write "              <input name='Hits' type='text' id='Hits' value='0' size='10' maxlength='10' style='text-align:center'>&nbsp;&nbsp; <font color='#0000FF'>�⹦�����ṩ������Ա�����õġ�����������Ҫ��ѽ��^_^</font>"
        Response.Write "            </td>"
        Response.Write "          </tr>"
    End If
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>¼��ʱ�䣺</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='UpdateTime' type='text' id='UpdateTime' value='" & Now() & "' maxlength='50'> ʱ���ʽΪ����-��-�� ʱ:��:�롱���磺<font color='#0000FF'>2003-5-12 12:32:47</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>��ɫ���</td>"
    Response.Write "            <td><select Name='SkinID'>" & GetSkin_Option(Session("SkinID")) & "</select>&nbsp;���ģ���а���CSS����ɫ��ͼƬ����Ϣ</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�������ģ�壺</td>"
    Response.Write "            <td><select Name='TemplateID'>" & GetTemplate_Option(ChannelID, 3, Session("TemplateID")) & "</select>&nbsp;���ģ���а����˰�����Ƶİ�ʽ����Ϣ</td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf
End Sub

Sub ShowTabs_Property_Modify(rsInfo)
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "���ʣ�</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='OnTop' type='checkbox' id='OnTop' value='yes'"
    If rsInfo("OnTop") = True Then Response.Write " checked"
    Response.Write "> �̶�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "              <input name='Hot' type='checkbox' id='Hot' value='yes' onclick=""javascript:document.myform.Hits.value='" & HitsOfHot & "'"" disabled> ����" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "              <input name='Elite' type='checkbox' id='Elite' value='yes'"
    If rsInfo("Elite") = True Then Response.Write " checked"
    Response.Write "> �Ƽ�" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "              " & ChannelShortName & "���ֵȼ���<select name='Stars' id='Stars'>" & GetStars(rsInfo("Stars")) & "</select>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    If ModuleType = 2 Then
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'>���ش�����</td>"
        Response.Write "            <td>"
        Response.Write "              ���գ� <input name='DayHits' type='text' id='DayHits' value='" & rsInfo("DayHits") & "' size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "              ���ܣ� <input name='WeekHits' type='text' id='WeekHits' value='" & rsInfo("WeekHits") & "' size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "              ���£� <input name='MonthHits' type='text' id='MonthHits' value='" & rsInfo("MonthHits") & "' size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;"
        Response.Write "              �ܼƣ� <input name='Hits' type='text' id='Hits' value='" & rsInfo("Hits") & "' size='10' maxlength='10'>"
        Response.Write "            </td>"
        Response.Write "          </tr>"
    Else
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'>�������</td>"
        Response.Write "            <td>"
        Response.Write "              <input name='Hits' type='text' id='Hits' value='" & rsInfo("Hits") & "' size='10' maxlength='10'>&nbsp;&nbsp;<font color='#0000FF'>�⹦�����ṩ������Ա�����õġ�����������Ҫ��ѽ��^_^</font>"
        Response.Write "            </td>"
        Response.Write "          </tr>"
    End If
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>¼��ʱ�䣺</td>"
    Response.Write "            <td>"
    Response.Write "              <input name='UpdateTime' type='text' id='UpdateTime' value='" & rsInfo("UpdateTime") & "' maxlength='50'> ʱ���ʽΪ����-��-�� ʱ:��:�롱���磺<font color='#0000FF'>2003-5-12 12:32:47</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>��ɫ���</td>"
    Response.Write "            <td><select Name='SkinID'>" & GetSkin_Option(rsInfo("SkinID")) & "</select>&nbsp;���ģ���а���CSS����ɫ��ͼƬ����Ϣ</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�������ģ�壺</td>"
    Response.Write "            <td><select Name='TemplateID'>" & GetTemplate_Option(ChannelID, 3, rsInfo("TemplateID")) & "</select>&nbsp;���ģ���а����˰�����Ƶİ�ʽ����Ϣ</td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf
End Sub

Sub ShowTabs_Purview_Add(ViewString)
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ViewString & "Ȩ�ޣ�</td>"
    Response.Write "            <td><input name='InfoPurview' type='radio' value='0' checked>�̳���ĿȨ�ޣ���������ĿΪ��֤��Ŀʱ������ѡ����<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='1'>���л�Ա����������ĿΪ������Ŀ���뵥����ĳЩ" & ChannelShortName & "����" & ViewString & "Ȩ�����ã�����ѡ����<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='2'>ָ����Ա�飨��������ĿΪ������Ŀ���뵥����ĳЩ" & ChannelShortName & "����" & ViewString & "Ȩ�����ã�����ѡ����<br>"
    Response.Write GetUserGroup("", "")
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ViewString & "������</td>"
    Response.Write "            <td><input name='InfoPoint' type='text' id='InfoPoint' value='" & Session("InfoPoint") & "' size='5' maxlength='4' style='text-align:center'> "
    Response.Write "              &nbsp;&nbsp;<font color='#0000FF'>�����������0������Ȩ�޵Ļ�Ա" & ViewString & "��" & ChannelShortName & "ʱ��������Ӧ��������Ϊ9999ʱ���⣩���οͽ��޷�" & ViewString & "��" & ChannelShortName & "</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ظ��շѣ�</td>"
    Response.Write "            <td><input name='ChargeType' type='radio' value='0' checked>���ظ��շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='1'>�����ϴ��շ�ʱ�� <input name='PitchTime' type='text' value='24' size='8' maxlength='8' style='text-align:center'> Сʱ�������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='2'>��Ա�ظ�" & ViewString & "��" & ChannelShortName & " <input name='ReadTimes' type='text' value='10' size='8' maxlength='8' style='text-align:center'> �κ������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='3'>�������߶�����ʱ�����շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='4'>����������һ������ʱ�������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='5'>ÿ" & ViewString & "һ�ξ��ظ��շ�һ�Σ����鲻Ҫʹ�ã�"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ֳɱ�����</td>"
    Response.Write "            <td><input name='DividePercent' type='text' id='DividePercent' value='0' size='5' maxlength='4' style='text-align:center'> %"
    Response.Write "              &nbsp;&nbsp;<font color='#0000FF'>�����������0���򽫰���������" & ViewString & "����ȡ�ĵ���֧����¼����</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf
End Sub

Sub ShowTabs_Purview_Batch(ViewString)
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyInfoPurview' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>" & ViewString & "Ȩ�ޣ�</td>"
    Response.Write "            <td><input name='InfoPurview' type='radio' value='0' checked>�̳���ĿȨ�ޣ���������ĿΪ��֤��Ŀʱ������ѡ����<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='1'>���л�Ա����������ĿΪ������Ŀ���뵥����ĳЩ" & ChannelShortName & "����" & ViewString & "Ȩ�����ã�����ѡ����<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='2'>ָ����Ա�飨��������ĿΪ������Ŀ���뵥����ĳЩ" & ChannelShortName & "����" & ViewString & "Ȩ�����ã�����ѡ����<br>"
    Response.Write GetUserGroup("", "")
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyInfoPoint' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>" & ViewString & "������</td>"
    Response.Write "            <td><input name='InfoPoint' type='text' id='InfoPoint' value='" & Session("InfoPoint") & "' size='5' maxlength='4' style='text-align:center'> "
    Response.Write "              &nbsp;&nbsp;<font color='#0000FF'>�����������0������Ȩ�޵Ļ�Ա" & ViewString & "��" & ChannelShortName & "ʱ��������Ӧ��������Ϊ9999ʱ���⣩���οͽ��޷�" & ViewString & "��" & ChannelShortName & "</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyChargeType' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>�ظ��շѣ�</td>"
    Response.Write "            <td><input name='ChargeType' type='radio' value='0' checked>���ظ��շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='1'>�����ϴ��շ�ʱ�� <input name='PitchTime' type='text' value='24' size='8' maxlength='8' style='text-align:center'> Сʱ�������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='2'>��Ա�ظ�" & ViewString & "��" & ChannelShortName & " <input name='ReadTimes' type='text' value='10' size='8' maxlength='8' style='text-align:center'> �κ������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='3'>�������߶�����ʱ�����շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='4'>����������һ������ʱ�������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='5'>ÿ" & ViewString & "һ�ξ��ظ��շ�һ�Σ����鲻Ҫʹ�ã�"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyDividePercent' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>�ֳɱ�����</td>"
    Response.Write "            <td><input name='DividePercent' type='text' id='DividePercent' value='0' size='5' maxlength='4' style='text-align:center'> %"
    Response.Write "              &nbsp;&nbsp;<font color='#0000FF'>�����������0���򽫰���������" & ViewString & "����ȡ�ĵ���֧����¼����</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf
End Sub

Sub ShowTabs_MyField_Batch()
    Dim tempModuleType
    tempModuleType = 0 - ModuleType
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & tempModuleType & " or ChannelID=" & ChannelID & " Order by FieldID")
    If rsField.BOF And rsField.EOF Then
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='600' colspan='3'></td>"
    Response.Write "          </tr>"
    Else
        Do While Not rsField.EOF
            Call WriteBatchFieldHTML(rsField("FieldName"), rsField("Title"), rsField("Tips"), rsField("FieldType"), rsField("DefaultValue"), rsField("Options"), rsField("EnableNull"))
            rsField.MoveNext
        Loop
    End If
    Set rsField = Nothing
    Response.Write "        </tbody>" & vbCrLf
End Sub

Sub ShowTabs_Purview_Modify(ViewString, rsInfo, strDisabled)
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ViewString & "Ȩ�ޣ�</td>"
    Response.Write "            <td><input name='InfoPurview' type='radio' value='0'" & strDisabled
    If rsInfo("InfoPurview") = 0 Then Response.Write " checked"
    Response.Write ">�̳���ĿȨ�ޣ���������ĿΪ��֤��Ŀʱ������ѡ����<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='1'" & strDisabled
    If rsInfo("InfoPurview") = 1 Then Response.Write " checked"
    Response.Write ">���л�Ա����������ĿΪ������Ŀ���뵥����ĳЩ" & ChannelShortName & "����" & ViewString & "Ȩ�����ã�����ѡ����<br>"
    Response.Write "            <input name='InfoPurview' type='radio' value='2'" & strDisabled
    If rsInfo("InfoPurview") = 2 Then Response.Write " checked"
    Response.Write ">ָ����Ա�飨��������ĿΪ������Ŀ���뵥����ĳЩ" & ChannelShortName & "����" & ViewString & "Ȩ�����ã�����ѡ����<br>"
    Response.Write GetUserGroup(rsInfo("arrGroupID") & "", strDisabled)
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "" & ViewString & "������</td>"
    Response.Write "            <td><input name='InfoPoint' type='text' id='InfoPoint' value='" & rsInfo("InfoPoint") & "' size='5' maxlength='4' style='text-align:center'" & strDisabled & ">&nbsp;&nbsp;&nbsp;&nbsp; <font color='#0000FF'>�������0�����Ա" & ViewString & "��" & ChannelShortName & "ʱ��������Ӧ��������Ϊ9999ʱ���⣩���οͽ��޷�" & ViewString & "��" & ChannelShortName & "��</font></td>"
    Response.Write "          </tr>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ظ��շѣ�</td>"
    Response.Write "            <td><input name='ChargeType' type='radio' value='0'" & strDisabled
    If rsInfo("ChargeType") = 0 Then Response.Write " checked"
    Response.Write ">���ظ��շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='1'" & strDisabled
    If rsInfo("ChargeType") = 1 Then Response.Write " checked"
    Response.Write ">�����ϴ��շ�ʱ�� <input name='PitchTime' type='text' value='" & rsInfo("PitchTime") & "' size='8' maxlength='8' style='text-align:center'" & strDisabled & "> Сʱ�������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='2'" & strDisabled
    If rsInfo("ChargeType") = 2 Then Response.Write " checked"
    Response.Write ">��Ա�ظ�" & ViewString & "��" & ChannelShortName & " <input name='ReadTimes' type='text' value='" & rsInfo("ReadTimes") & "' size='8' maxlength='8' style='text-align:center'" & strDisabled & "> �κ������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='3'" & strDisabled
    If rsInfo("ChargeType") = 3 Then Response.Write " checked"
    Response.Write ">�������߶�����ʱ�����շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='4'" & strDisabled
    If rsInfo("ChargeType") = 4 Then Response.Write " checked"
    Response.Write ">����������һ������ʱ�������շ�<br>"
    Response.Write "            <input name='ChargeType' type='radio' value='5'" & strDisabled
    If rsInfo("ChargeType") = 5 Then Response.Write " checked"
    Response.Write ">ÿ" & ViewString & "һ�ξ��ظ��շ�һ�Σ����鲻Ҫʹ�ã�"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�ֳɱ�����</td>"
    Response.Write "            <td><input name='DividePercent' type='text' id='DividePercent' value='" & rsInfo("DividePercent") & "' size='5' maxlength='4' style='text-align:center'" & strDisabled & "> %"
    Response.Write "              &nbsp;&nbsp;<font color='#0000FF'>�����������0���򽫰���������" & ViewString & "����ȡ�ĵ���֧����¼����</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf
End Sub

Sub ShowTabs_Vote_Add()
    Dim i
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>���õ��飺</td><td><input name='ShowVote' type='checkbox' id='ShowVote' value='yes'></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�������⣺</td><td><textarea name='VoteTitle' cols='50' rows='4'></textarea></td>"
    Response.Write "          </tr>"
    For i = 1 To 8
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'>ѡ��" & i & "��</td><td><input type='text' name='select" & i & "' size='36'>&nbsp;Ʊ����<input type='text' name='answer" & i & "' size='10'></td>"
        Response.Write "          </tr>"
    Next

    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�������ͣ�</td><td><select name='VoteType' id='VoteType'><option value='Single' selected>��ѡ</option><option value='Multi'>��ѡ</option></select></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>��ʼʱ�䣺</td><td><input type='text' name='BeginTime' size='20' value='" & Now() & "'>&nbsp;���鿪ʼ��ʱ��</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>��ֹʱ�䣺</td><td><input type='text' name='EndTime' size='20' value='" & Now() + 30 & "'>&nbsp;���������ʱ��</td>"
    Response.Write "          </tr>"
    Response.Write "        </tbody>" & vbCrLf
End Sub

Sub ShowTabs_Vote_Modify(rsInfo)
    Dim UseVote, i, rsVote
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "          <tr class='tdbg'>"
    If rsInfo("VoteID") > 0 Then
        Response.Write "  <td width='120' align='right' class='tdbg5'>���õ��飺</td><td><input name='ShowVote' type='checkbox' id='ShowVote' value='yes' checked></td>"
        Set rsVote = Conn.Execute("select * from PE_Vote where ID=" & rsInfo("VoteID"))
        If Not (rsVote.BOF And rsVote.EOF) Then
            UseVote = True
        End If
    Else
        Response.Write "  <td width='120' align='right' class='tdbg5'>���õ��飺</td><td><input name='ShowVote' type='checkbox' id='ShowVote' value='yes'></td>"
    End If

    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�������⣺</td><td>"
    If UseVote = True Then
        Response.Write "<textarea name='VoteTitle' cols='50' rows='4'>" & rsVote("Title") & "</textarea>"
    Else
        Response.Write "<textarea name='VoteTitle' cols='50' rows='4'></textarea>"
    End If
    Response.Write "          </td></tr>"
    For i = 1 To 8
        Response.Write "          <tr class='tdbg'>"
        Response.Write "            <td width='120' align='right' class='tdbg5'>ѡ��" & i & "��</td><td>"
        If UseVote = True Then
            Response.Write "<input type='text' name='select" & i & "' size='36' value=" & rsVote("Select" & i) & ">&nbsp;Ʊ����<input type='text' name='answer" & i & "' size='10' value=" & rsVote("Answer" & i) & ">"
        Else
            Response.Write "<input type='text' name='select" & i & "' size='36'>&nbsp;Ʊ����<input type='text' name='answer" & i & "' size='10'>"
        End If
        Response.Write "          </td></tr>"
    Next
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>�������ͣ�</td><td><select name='VoteType' id='VoteType'>"
    If UseVote = True Then
        If rsVote("VoteType") = "Single" Then
            Response.Write "<option value='Single' selected>��ѡ</option>"
        Else
            Response.Write "<option value='Single'>��ѡ</option>"
        End If
        If rsVote("VoteType") = "Multi" Then
            Response.Write "<option value='Multi' selected>��ѡ</option>"
        Else
            Response.Write "<option value='Multi'>��ѡ</option>"
        End If
    Else
        Response.Write "<option value='Single' selected>��ѡ</option><option value='Multi'>��ѡ</option>"
    End If
    Response.Write "          </select></td></tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>��ʼʱ�䣺</td><td>"
    If UseVote = True Then
        Response.Write "<input type='text' name='BeginTime' size='20' value='" & rsVote("VoteTime") & "'>"
    Else
        Response.Write "<input type='text' name='BeginTime' size='20' value='" & Now() & "'>"
    End If
    Response.Write "&nbsp;���鿪ʼ��ʱ��</td></tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>��ֹʱ�䣺</td><td>"
    If UseVote = True Then
        Response.Write "<input type='text' name='EndTime' size='20' value='" & rsVote("EndTime") & "'>"
    Else
        Response.Write "<input type='text' name='EndTime' size='20' value='" & Now() + 30 & "'>"
    End If
    Response.Write "&nbsp;���������ʱ��</td></tr>"
    Response.Write "        </tbody>" & vbCrLf
    If UseVote = True Then
        Set rsVote = Nothing
    End If
End Sub

Sub ShowTabs_MyField_Add()
    Dim tempModuleType
    tempModuleType = 0 - ModuleType
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & tempModuleType & " or ChannelID=" & ChannelID & " Order by FieldID")
    Do While Not rsField.EOF
        Call WriteFieldHTML(rsField("FieldName"), rsField("Title"), rsField("Tips"), rsField("FieldType"), rsField("DefaultValue"), rsField("Options"), rsField("EnableNull"))
        rsField.MoveNext
    Loop
    Set rsField = Nothing
    Response.Write "        </tbody>" & vbCrLf
End Sub

Sub ShowTabs_MyField_Modify(rsInfo)
    Dim tempModuleType
    tempModuleType = 0 - ModuleType
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & tempModuleType & " or ChannelID=" & ChannelID & "")
    Do While Not rsField.EOF
        Call WriteFieldHTML(rsField("FieldName"), rsField("Title"), rsField("Tips"), rsField("FieldType"), rsInfo(Trim(rsField("FieldName"))), rsField("Options"), rsField("EnableNull"))
        rsField.MoveNext
    Loop
    Set rsField = Nothing
    
    Response.Write "        </tbody>" & vbCrLf
End Sub

Sub ShowTabs_MyField_View(rsInfo)
    Dim tempModuleType
    tempModuleType = 0 - ModuleType
    Response.Write "        <tbody id='Tabs' style='display:none'>" & vbCrLf
    Dim rsField
    Set rsField = Conn.Execute("select * from PE_Field where ChannelID=" & tempModuleType & " or ChannelID=" & ChannelID & "")
    Do While Not rsField.EOF
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td width='120' align='right' class='tdbg5'>" & rsField("Title") & "</td>"
        Response.Write "    <td>" & PE_HTMLEncode(rsInfo(Trim(rsField("FieldName")))) & "</td>"
        Response.Write "  </tr>"
        rsField.MoveNext
    Loop
    Set rsField = Nothing
    Response.Write "        </tbody>" & vbCrLf
End Sub

Sub ShowTabs_Status_Add()
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "״̬��</td>"
    Response.Write "            <td><input name='Status' type='radio' id='Status' value='-1'>�ݸ�&nbsp;&nbsp;"
    Response.Write "                <input Name='Status' Type='Radio' Id='Status' Value='0' "
    If MyStatus < 3 Then
        Response.Write " checked>�����&nbsp;&nbsp;"
    Else
        Response.Write " >�����&nbsp;&nbsp;"
        Response.Write "<input Name='Status' Type='Radio' Id='Status' Value='" & MyStatus & "' checked>" & arrStatus(MyStatus) & ""
        If UseCreateHTML > 0 And AutoCreateType > 0 Then
            Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='CreateImmediate' type='checkbox' value='Yes' checked>��������"
        End If
    End If
    
    Response.Write "            </td>"
    Response.Write "          </tr>"
End Sub

Sub ShowTabs_Status_Modify(rsInfo)
    If MyStatus = 1 then 
        Dim rsClassCheck,CheckParentID,CheckParentPath
        Set rsClassCheck = Conn.execute("select * from PE_Class where ClassID = "&ClassID)
        CheckParentID = rsClassCheck("ParentID")
        CheckParentPath = rsClassCheck("ParentPath")
        rsClassCheck.close
        set rsClassCheck = nothing
        If CheckParentID > 0 Then
            PurviewChecked = CheckPurview_Class(arrClass_Check, CheckParentPath & "," & ClassID)
        Else
            PurviewChecked = CheckPurview_Class(arrClass_Check, ClassID)
        End If
    ElseIf MyStatus > 1 Then
        PurviewChecked = True
    Else
        PurviewChecked = False							
    End If	
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='120' align='right' class='tdbg5'>" & ChannelShortName & "״̬��</td>"
    Response.Write "            <td>"
    If rsInfo("Inputer") = UserName And rsInfo("Status") <= MyStatus Then
        If PurviewChecked = True Or AdminPurview = 1 then 	
            Response.Write "<Input Name='Status' Type='radio' Id='Status' Value='-1'"
            If rsInfo("Status") = -1 Then
                Response.Write " checked"
            End If
            Response.Write "> �ݸ�&nbsp;&nbsp;"
            Response.Write "<Input Name='Status' Type='radio' Id='Status' Value='0'"
            If rsInfo("Status") < MyStatus Then
                Response.Write "checked"
            End If
            Response.Write "> δ���&nbsp;&nbsp;"
            Response.Write "<Input Name='Status' Type='radio' Id='Status' Value='1'"
            If rsInfo("Status") = MyStatus Then
                Response.Write "Checked"
            End If
            Response.Write ">"&arrStatus(MyStatus)	
        Else
            Response.Write "<Input style='display:none'  Name='Status' Type='radio' Id='Status' Value='"& rsInfo("Status") &"' Checked>"&arrStatus(rsInfo("Status"))       		
        End If	
    ElseIf rsInfo("Inputer") <> UserName Then
        If rsInfo("Status") = -1 Then
            Response.Write "�ݸ�"
        ElseIf rsInfo("Status") = -2 Then
            Response.Write "�˸�"
        Else
            Response.Write arrStatus(rsInfo("Status"))
        End If
    Else
        If rsInfo("Status") = -1 Then
            Response.Write "<Input style='display:none'  Name='Status' Type='radio' Id='Status' Value='"& rsInfo("Status") &"' Checked>�ݸ�"
        ElseIf rsInfo("Status") = -2 Then
            Response.Write "<Input style='display:none'  Name='Status' Type='radio' Id='Status' Value='"& rsInfo("Status") &"' Checked>�˸�"
        Else
            Response.Write "<Input style='display:none'  Name='Status' Type='radio' Id='Status' Value='"& rsInfo("Status") &"' Checked>"& arrStatus(rsInfo("Status"))
        End If
    End If
    If UseCreateHTML > 0 And AutoCreateType > 0  Then
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='CreateImmediate' type='checkbox' value='Yes' checked>��������"
    End If
    Response.Write "            </td>"
    Response.Write "          </tr>"
    If rsInfo("Inputer") <> UserName Then
        If PurviewChecked = True Or AdminPurview = 1 then 
            Response.Write "          <tr class='tdbg'>"
            Response.Write "            <td width='120' align='right' class='tdbg5'>��˲�����</td>"
            Response.Write "            <td>"
            Response.Write "<input name='Status' type='radio' value='" & rsInfo("Status") & "' checked onClick=""tabMsg.style.display='none';""> ���ı䵱ǰ״̬&nbsp;&nbsp;&nbsp;&nbsp;"
            If rsInfo("Status") >= 0 Then
                Response.Write "<input name='Status' type='radio' value='-2' onClick=""tabMsg.style.display='';document.myform.MsgTitle.value='�˸�֪ͨ';document.myform.MsgContent.value='" & EmailOfReject & "';""> �˸�&nbsp;&nbsp;&nbsp;&nbsp;"
            End If
            If rsInfo("Status") < MyStatus Then
                Response.Write "<input name='Status' type='radio' value='1'"
                If MyStatus = 3 Then
                    Response.Write " onClick=""tabMsg.style.display='';document.myform.MsgTitle.value='���¼��֪ͨ';document.myform.MsgContent.value='" & EmailOfPassed & "';"""
                Else
                    Response.Write " onClick=""tabMsg.style.display='none';"""
                End If
                Response.Write "> " & arrStatus(MyStatus)
            End If
            Response.Write "            </td>"
            Response.Write "          </tr>"


            Response.Write "        <tbody id='tabMsg' style='display:none'>" & vbCrLf
            Response.Write "          <tr class='tdbg'>"
            Response.Write "            <td width='120' align='right' class='tdbg5'>֪ͨ¼���ߣ�</td><td><input type='checkbox' name='SendMessageToInputer' value='Yes' checked>����վ�ڶ���֪ͨ¼����&nbsp;&nbsp;&nbsp;&nbsp;<input type='checkbox' name='SendEmailToInputer' value='Yes' checked>����Email֪ͨ¼����</td> "
            Response.Write "          </tr>"
            Response.Write "          <tr class='tdbg'>"
            Response.Write "            <td width='120' align='right' class='tdbg5'>֪ͨ���⣺</td><td><input type='text' name='MsgTitle' MaxLength='100' size='70' value=''></td>"
            Response.Write "          </tr>"
            Response.Write "          <tr class='tdbg'>"
            Response.Write "            <td width='120' align='right' class='tdbg5'>֪ͨ���ݣ�</td><td><Textarea name='MsgContent'cols='70' rows='5'></textarea></td>"
            Response.Write "          </tr>"
            Response.Write "        </tbody>" & vbCrLf
        Else
            Response.Write "          <tr class='tdbg'>"
            Response.Write "            <td width='120' align='right' class='tdbg5'>��˲�����</td>"
            Response.Write "            <td>"
            Response.Write "<input name='Status' type='radio' value='" & rsInfo("Status") & "' checked onClick=""tabMsg.style.display='none';""> ���ı䵱ǰ״̬&nbsp;&nbsp;&nbsp;&nbsp;"		
            Response.Write "            </td>"
            Response.Write "          </tr>"		
        End IF
    End If
End Sub

Sub ShowBatchCommon()
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyKeyword' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>�ؼ��֣�</td>"
    Response.Write "            <td><input name='Keyword' type='text' id='Keyword' value='' size='30' maxlength='255'> <font color='#FF0000'>*</font> " & GetKeywordList("Admin", ChannelID)
    Response.Write "              <br><font color='#0000FF'>�����������" & ChannelShortName & "�����������ؼ��֣��м���<font color='#FF0000'>��|��</font>���������ܳ���&quot;'&?;:()���ַ���</font>"
    Response.Write "            </td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyOnTop' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "���ʣ�</td>"
    Response.Write "            <td><input name='OnTop' type='checkbox' id='OnTop' value='Yes'> �̶�" & ChannelShortName & ""
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyElite' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>" & ChannelShortName & "���ʣ�</td>"
    Response.Write "            <td><input name='Elite' type='checkbox' id='Elite' value='Yes'> �Ƽ�" & ChannelShortName & ""
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyStars' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>���ֵȼ���</td>"
    Response.Write "            <td><select name='Stars' id='Stars'>" & GetStars(3) & "</select></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyHits' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>�������</td>"
    Response.Write "            <td><input name='Hits' type='text' id='Hits' value='0' size='10' maxlength='10'></td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyUpdateTime' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>����ʱ�䣺</td>"
    Response.Write "            <td><input name='UpdateTime' type='text' id='UpdateTime' value='" & Date & "' size='10' maxlength='10'> ֻ�޸ĸ���ʱ������ڲ��֣�ʱ�䱣����</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifySkin' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>��ɫ���</td>"
    Response.Write "            <td><select Name='SkinID'>" & GetSkin_Option(0) & "</select>&nbsp;���ģ���а���CSS����ɫ��ͼƬ����Ϣ</td>"
    Response.Write "          </tr>"
    Response.Write "          <tr class='tdbg'>"
    Response.Write "            <td width='30' align='center' class='tdbg5'><input type='checkbox' name='ModifyTemplate' value='Yes'></td>"
    Response.Write "            <td width='100' align='right' class='tdbg5'>�������ģ�壺</td>"
    Response.Write "            <td><select Name='TemplateID'>" & GetTemplate_Option(ChannelID, 3, 0) & "</select>&nbsp;���ģ���а����˰�����Ƶİ�ʽ����Ϣ</td>"
    Response.Write "          </tr>"
End Sub

Sub ShowComment(InfoID)
    Dim MaxPerPage
    MaxPerPage = 10
    Response.Write "<table width='100%' border='0' cellspacing='1' cellpadding='2' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='30' height='22'><strong>ID</strong></td>"
    Response.Write "    <td height='22'><strong>����</strong></td>"
    Response.Write "    <td width='60' height='22'><strong>������</strong></td>"
    Response.Write "    <td width='120' height='22'><strong>������IP</strong></td>"
    Response.Write "    <td width='120' height='22'><strong>����ʱ��</strong></td>"
    Response.Write "    <td width='100' height='22'><strong>����</strong></td>"
    Response.Write "  </tr>"
    Dim rsComment, sql, TotalPut
    sql = "select * from PE_Comment where ModuleType=" & ModuleType & " and InfoID=" & InfoID
    Set rsComment = Conn.Execute(sql)
    If rsComment.EOF Then
        Response.Write "<tr class='tdbg' align='center' height='50'><td colspan='20'>��ʱû���κ��˶Ա�" & ChannelShortName & "��������</td></tr>"
    Else
        TotalPut = rsComment.RecordCount
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > TotalPut Then
            If (TotalPut Mod MaxPerPage) = 0 Then
                CurrentPage = TotalPut \ MaxPerPage
            Else
                CurrentPage = TotalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < TotalPut Then
                rsComment.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
    
        Dim i
        i = 0
        Do While Not rsComment.EOF
            Response.Write "  <tr class='tdbg'>"
            Response.Write "    <td width='30' align='center'>" & rsComment("CommentID") & "</td>"
            Response.Write "    <td><a href=# title='" & rsComment("Content") & "'>" & Left(rsComment("Content"), 25) & "</a>" & "</td>"
            Response.Write "    <td width='60' align='center'>" & rsComment("UserName") & "</td>"
            Response.Write "    <td width='120' align='center'>" & rsComment("IP") & "</td>"
            Response.Write "    <td width='120' align='center'>" & rsComment("WriteTime") & "</td>"
            Response.Write "    <td width='100' align='center'>"
            If AdminPurview = 1 Or AdminPurview_Channel = 1 Then
                If rsComment("ReplyName") <> "" Then
                    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
                Else
                    Response.Write "<a href='Admin_Comment.asp?ChannelID=" & ChannelID & "&Action=Reply&CommentID=" & rsComment("Commentid") & "'>�ظ�</a>&nbsp;&nbsp;"
                End If
                Response.Write "<a href='Admin_Comment.asp?ChannelID=" & ChannelID & "&Action=Modify&CommentID=" & rsComment("Commentid") & "'>�޸�</a>&nbsp;&nbsp;"
                Response.Write "<a href='Admin_Comment.asp?ChannelID=" & ChannelID & "&Action=Del&CommentID=" & rsComment("CommentID") & "'>ɾ��</a>"
            End If
            Response.Write "</td></tr>"
            If rsComment("ReplyName") <> "" Then
                Response.Write "<tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'""> "
                Response.Write "<td align='center'>&nbsp;</td>"
                Response.Write "<td colspan='4'>����Ա��" & rsComment("ReplyName") & "���� " & rsComment("ReplyTime") & " �ظ���<br><div style='padding:0px 20px'>" & rsComment("ReplyContent") & "</div></td>"
                Response.Write "<td align='center'><a href='Admin_Comment.asp?ChannelID=" & ChannelID & "&Action=Reply&CommentID=" & rsComment("CommentID") & "'>�޸Ļظ�����</a></td>"
                Response.Write "</tr>"
            End If
    
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsComment.MoveNext
        Loop
    End If
    rsComment.Close
    Set rsComment = Nothing
    Response.Write "</table><br>"

End Sub

Sub ShowConsumeLog(InfoID)
    Dim MaxPerPage
    MaxPerPage = 10
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "  <tr align='center' class='title'>"
    Response.Write "    <td width='120'>����ʱ��</td>"
    Response.Write "    <td width='80'>������</td>"
    Response.Write "    <td width='100'>IP��ַ</td>"
    Response.Write "    <td width='60'>���ѵ���</td>"
    Response.Write "    <td width='50'>�ظ�����</td>"
    Response.Write "    <td>��ע/˵��</td>"
    Response.Write "  </tr>"
    
    Dim rsConsumeLog, sqlConsumeLog
    Dim TotalPoint, TotalPut
    TotalPoint = 0
    
    sqlConsumeLog = "select * from PE_ConsumeLog where ModuleType=" & ModuleType & " and InfoID=" & InfoID & " order by LogID desc"
    Set rsConsumeLog = Server.CreateObject("Adodb.RecordSet")
    rsConsumeLog.Open sqlConsumeLog, Conn, 1, 1
    If rsConsumeLog.BOF And rsConsumeLog.EOF Then
        TotalPut = 0
        Response.Write "<tr class='tdbg' height='50'><td colspan='20' align='center'>û���κ�������Ѽ�¼��</td></tr>"
    Else
        TotalPut = rsConsumeLog.RecordCount
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > TotalPut Then
            If (TotalPut Mod MaxPerPage) = 0 Then
                CurrentPage = TotalPut \ MaxPerPage
            Else
                CurrentPage = TotalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < TotalPut Then
                rsConsumeLog.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
    
        Dim i
        i = 0
        Do While Not rsConsumeLog.EOF
            TotalPoint = TotalPoint + rsConsumeLog("Point")
    
            Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
            Response.Write "    <td width='120' align='center'>" & rsConsumeLog("LogTime") & "</td>"
            Response.Write "    <td width='80' align='center'><a href='Admin_User.asp?Action=Show&UserName=" & rsConsumeLog("UserName") & "&InfoType=2'>" & rsConsumeLog("UserName") & "</a></td>"
            Response.Write "    <td width='100' align='center'>" & rsConsumeLog("IP") & "</td>"
            Response.Write "    <td width='60' align='right'>" & rsConsumeLog("Point") & "</td>"
            Response.Write "    <td width='50' align='center'>" & rsConsumeLog("Times") & "</td>"
            Response.Write "    <td align='left'>" & rsConsumeLog("Remark") & "</td>"
            Response.Write "  </tr>"
    
            i = i + 1
            If i >= MaxPerPage Then Exit Do
            rsConsumeLog.MoveNext
        Loop
    End If
    rsConsumeLog.Close
    Set rsConsumeLog = Nothing
    
    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='3' align='right'>��ҳ�ϼƣ�</td>"
    Response.Write "    <td align='right'>" & TotalPoint & "</td>"
    Response.Write "    <td colspan='3'>&nbsp;</td>"
    Response.Write "  </tr>"

    Dim trs, TotalPointAll
    Set trs = Conn.Execute("select sum(Point) from PE_ConsumeLog where ModuleType=" & ModuleType & " and InfoID=" & InfoID & "")
    If IsNull(trs(0)) Then
        TotalPointAll = 0
    Else
        TotalPointAll = trs(0)
    End If
    Set trs = Nothing
    Response.Write "  <tr class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">"
    Response.Write "    <td colspan='3' align='right'>�ܼƵ�����</td>"
    Response.Write "    <td align='right'>" & TotalPointAll & "</td>"
    Response.Write "    <td colspan='3'> </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write ShowPage(strFileName, TotalPut, MaxPerPage, CurrentPage, True, True, "��������ϸ��¼", True)
End Sub

Sub ShowForm_MoveToClass()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    
    Dim tChannelID, BatchInfoID
    tChannelID = Trim(Request("tChannelID"))
    If tChannelID = "" Then
        tChannelID = ChannelID
    Else
        tChannelID = PE_CLng(tChannelID)
    End If
    BatchInfoID = ReplaceBadChar(Request("Batch" & ModuleName & "ID"))
    If BatchInfoID = "" Then
        BatchInfoID = ReplaceBadChar(Request(ModuleName & "ID"))
    End If
        
    Response.Write "<form method='POST' name='myform' action='Admin_" & ModuleName & ".asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='4' align='center'><b>�����ƶ�" & ChannelShortName & "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='left' class='tdbg'>"
    Response.Write "      <td valign='top' width='300'>"
    Response.Write "        <input type='radio' name='" & ModuleName & "Type' value='1' checked>ָ��" & ChannelShortName & "ID��<input type='text' name='Batch" & ModuleName & "ID' value='" & BatchInfoID & "' size='30'><br>"
    Response.Write "        <input type='radio' name='" & ModuleName & "Type' value='2'>ָ����Ŀ��" & ChannelShortName & "��<br><select name='BatchClassID' size='2' multiple style='height:360px;width:300px;'>" & GetClass_Channel(ChannelID) & "</select><br>"
    Response.Write "        <input type='button' name='Submit' value='  ѡ��������Ŀ  ' onclick='SelectAll()'>"
    Response.Write "        <input type='button' name='Submit' value='ȡ��ѡ��������Ŀ' onclick='UnSelectAll()'>"
    Response.Write "      </td>"
    Response.Write "      <td align='center' >�ƶ���&gt;&gt;</td>"
    Response.Write "      <td valign='top'>"
    Response.Write "        Ŀ��Ƶ����<select name='tChannelID' onChange='document.myform.submit();'>" & GetChannel_Option(ModuleType, tChannelID) & "</select><br>"
    Response.Write "        Ŀ����Ŀ��<font color=red>������ָ��Ϊ�ⲿ��Ŀ��</font><br><select name='tClassID' size='2' style='height:360px;width:300px;'>" & GetClass_Channel(tChannelID) & "</select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "  <p align='center'>"
    Response.Write "    <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "    <input name='Action' type='hidden' id='Action' value='MoveToClass'>"
    Response.Write "    <input name='add' type='submit'  id='Add' value=' ִ�������� ' style='cursor:hand;' onClick=""document.myform.Action.value='DoMoveToClass';"">&nbsp; "
    Response.Write "    <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p>"
    Response.Write "</form>"
    Response.Write "<center><b>ע�⣺</b>��Ƶ��" & ChannelShortName & "�ƶ���Ƶ�����Զ����ֶ����ݲ��ᱻ�ƶ���</center>"
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.BatchClassID.length;i++){" & vbCrLf
    Response.Write "    document.myform.BatchClassID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.BatchClassID.length;i++){" & vbCrLf
    Response.Write "    document.myform.BatchClassID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub ShowForm_AddToSpecial()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    
    Response.Write "<form method='POST' name='myform' action='Admin_" & ModuleName & ".asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='4' align='center'><b>��" & ChannelShortName & "��ӵ�ר����</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='left' class='tdbg'>"
    Response.Write "      <td width='100' class='tdbg5'>ѡ����" & ChannelShortName & "ID��</td><td><input type='text' name='BatchInfoID' value='" & ReplaceBadChar(Request("InfoID")) & "' size='50'></td></tr>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='left' class='tdbg'>"
    Response.Write "      <td width='100' class='tdbg5' valign='top'>��ӵ�Ŀ��ר�⣺</td><td><select name='tSpecialID' size='2' multiple style='height:300px;width:300px;'>" & GetSpecial_Option(0) & "</select></td></tr>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "  <p align='center'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='DoAddToSpecial'>"
    Response.Write "   <input name='add' type='submit'  id='Add' value=' ִ�������� ' style='cursor:hand;'>&nbsp; "
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Special';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
End Sub

Sub ShowForm_MoveToSpecial()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
            
    Response.Write "<form method='POST' name='myform' action='Admin_" & ModuleName & ".asp' target='_self'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='4' align='center'><b>�����ƶ�" & ChannelShortName & "</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr align='left' class='tdbg'>"
    Response.Write "      <td valign='top' width='300'>"
    Response.Write "        <input type='radio' name='InfoType' value='1' checked>ָ��" & ChannelShortName & "ID��<input type='text' name='BatchInfoID' value='" & ReplaceBadChar(Request("InfoID")) & "' size='30'><br>"
    Response.Write "        <input type='radio' name='InfoType' value='2'>ָ��ר���" & ChannelShortName & "��<br><select name='BatchSpecialID' size='2' multiple style='height:360px;width:300px;'>" & GetSpecial_Option(0) & "</select><br>"
    Response.Write "        <input type='button' name='Submit' value='  ѡ������ר��  ' onclick='SelectAll()'>"
    Response.Write "        <input type='button' name='Submit' value='ȡ��ѡ������ר��' onclick='UnSelectAll()'>"
    Response.Write "      </td>"
    Response.Write "      <td align='center' >�ƶ���&gt;&gt;</td>"
    Response.Write "      <td valign='top'>Ŀ��ר�⣺</font><br><select name='tSpecialID' size='2' style='height:360px;width:300px;'>" & GetSpecial_Option(0) & "</select></td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "  <p align='center'>"
    Response.Write "   <input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "   <input name='Action' type='hidden' id='Action' value='DoMoveToSpecial'>"
    Response.Write "   <input name='add' type='submit'  id='Add' value=' ִ�������� ' style='cursor:hand;'>&nbsp; "
    Response.Write "   <input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage';"" style='cursor:hand;'>"
    Response.Write "  </p><br>"
    Response.Write "</form>"
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "function SelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.BatchSpecialID.length;i++){" & vbCrLf
    Response.Write "    document.myform.BatchSpecialID.options[i].selected=true;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function UnSelectAll(){" & vbCrLf
    Response.Write "  for(var i=0;i<document.myform.BatchSpecialID.length;i++){" & vbCrLf
    Response.Write "    document.myform.BatchSpecialID.options[i].selected=false;}" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Sub DoMoveToSpecial()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    
    Dim InfoType, BatchInfoID, BatchSpecialID
    Dim tSpecialID, tChannelDir, tUploadDir
    Dim rsBatchMove, sqlBatchMove
    
    InfoType = PE_CLng(Trim(Request("InfoType")))
    BatchInfoID = Trim(Request.Form("BatchInfoID"))
    BatchSpecialID = Trim(Request.Form("BatchSpecialID"))
    tSpecialID = PE_CLng(Trim(Request("tSpecialID")))
    
    If InfoType = 1 Then
        If IsValidID(BatchInfoID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����ƶ���" & ChannelShortName & "��ID</li>"
        End If
    Else
        If IsValidID(BatchSpecialID) = False Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��Ҫ�����ƶ���" & ChannelShortName & "��ר��</li>"
        End If
    End If
    If tSpecialID = 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ŀ��ר�⣡</li>"
    End If
    If FoundErr = True Then Exit Sub
    
    If InfoType = 1 Then
        sqlBatchMove = "select * from PE_InfoS where ModuleType=" & ModuleType & " and InfoID in (" & BatchInfoID & ")"
    Else
        sqlBatchMove = "select * from PE_InfoS where ModuleType=" & ModuleType & " and SpecialID in (" & BatchSpecialID & ")"
    End If
    Set rsBatchMove = Conn.Execute(sqlBatchMove)
    Do While Not rsBatchMove.EOF
        If PE_CLng(Conn.Execute("select count(InfoID) from PE_InfoS where ModuleType=" & ModuleType & " and SpecialID=" & tSpecialID & " and ItemID=" & rsBatchMove("ItemID") & "")(0)) > 0 Then
            Conn.Execute ("delete from PE_InfoS where InfoID=" & rsBatchMove("InfoID") & "")
        Else
            Conn.Execute ("update PE_InfoS set SpecialID=" & tSpecialID & " where InfoID=" & rsBatchMove("InfoID") & "")
        End If
        rsBatchMove.MoveNext
    Loop
    rsBatchMove.Close
    Set rsBatchMove = Nothing

    Call WriteSuccessMsg("�ɹ���ѡ����" & ChannelShortName & "�ƶ���Ŀ��ר���У�", "Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Special")
    Call ClearSiteCache(0)
End Sub

Sub DoAddToSpecial()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    Dim BatchInfoID, tSpecialID, rsInfo
    tSpecialID = Trim(Request("tSpecialID"))
    BatchInfoID = Trim(Request("BatchInfoID"))
    If IsValidID(BatchInfoID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��InfoID��</li>"
    End If
    If IsValidID(tSpecialID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��Ŀ��ר�⣡</li>"
    End If
    If FoundErr = True Then Exit Sub
    
    Dim arrSpecialID, i
    arrSpecialID = Split(tSpecialID, ",")
    Set rsInfo = Conn.Execute("select * from PE_InfoS where ModuleType=" & ModuleType & " and InfoID in (" & BatchInfoID & ") order by InfoID desc")
    If Not (rsInfo.BOF And rsInfo.EOF) Then
        For i = 0 To UBound(arrSpecialID)
            tSpecialID = PE_CLng(arrSpecialID(i))
            If tSpecialID > 0 Then
                rsInfo.movefirst
                Do While Not rsInfo.EOF
                    If rsInfo("SpecialID") = 0 Then
                        Conn.Execute ("update PE_InfoS set SpecialID=" & tSpecialID & " where InfoID=" & rsInfo("InfoID") & "")
                    Else
                        If PE_CLng(Conn.Execute("select count(InfoID) from PE_InfoS where ModuleType=" & ModuleType & " and SpecialID=" & tSpecialID & " and ItemID=" & rsInfo("ItemID") & "")(0)) = 0 Then
                            Conn.Execute ("insert into PE_InfoS (ModuleType,SpecialID,ItemID) values (" & ModuleType & "," & tSpecialID & "," & rsInfo("ItemID") & ")")
                        End If
                    End If
                    rsInfo.MoveNext
                Loop
            End If
        Next
    End If
    rsInfo.Close
    Set rsInfo = Nothing

    Call WriteSuccessMsg("�ɹ���ѡ����" & ChannelShortName & "�ƶ���Ŀ��ר���У�", "Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Special")
    Call ClearSiteCache(0)
End Sub

Sub DelFromSpecial()
    If AdminPurview = 2 And AdminPurview_Channel > 2 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ������Ȩ�޲�����</li>"
        Exit Sub
    End If
    Dim InfoID, rsInfo
    InfoID = Trim(Request("InfoID"))
    If IsValidID(InfoID) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��InfoID��</li>"
        Exit Sub
    End If
    Set rsInfo = Conn.Execute("select * from PE_InfoS where ModuleType=" & ModuleType & " and InfoID in (" & InfoID & ") order by InfoID desc")
    Do While Not rsInfo.EOF
        If PE_CLng(Conn.Execute("select count(InfoID) from PE_InfoS where ModuleType=" & ModuleType & " and ItemID=" & rsInfo("ItemID") & "")(0)) > 1 Then
            Conn.Execute ("delete from PE_InfoS where InfoID=" & rsInfo("InfoID") & "")
        Else
            Conn.Execute ("update PE_InfoS set SpecialID=0 where InfoID=" & rsInfo("InfoID") & "")
        End If
        rsInfo.MoveNext
    Loop
    rsInfo.Close
    Set rsInfo = Nothing

    Call WriteSuccessMsg("�ɹ���ѡ����" & ChannelShortName & "������ר�����Ƴ���", "Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Special")
    Call ClearSiteCache(0)
End Sub


'**************************************************
'��������GetSpecialList
'��  �ã�Ƶ��������ר�⵼��
'��  ����ChannelID ---- Ƶ��ID
'        SpecialID ---- ר��ID
'        FileName ---- ר������
'����ֵ��ר�⵼��
'**************************************************
Function GetSpecialList()
    Dim rsSpecial, sqlSpecial, strSpecial, i
    sqlSpecial = "select * from PE_Special where ChannelID=" & ChannelID & " order by OrderID"
    Set rsSpecial = Conn.Execute(sqlSpecial)
    If rsSpecial.BOF And rsSpecial.EOF Then
        strSpecial = strSpecial & "û���κ�ר��"
    Else
        i = 1
        strSpecial = "| "
        Do While Not rsSpecial.EOF
            If rsSpecial("SpecialID") = SpecialID Then
                strSpecial = strSpecial & "<a href='" & FileName & "&SpecialID=" & rsSpecial("SpecialID") & "'><font color=red>" & rsSpecial("SpecialName") & "</font></a>"
            Else
                strSpecial = strSpecial & "<a href='" & FileName & "&SpecialID=" & rsSpecial("SpecialID") & "'>" & rsSpecial("SpecialName") & "</a>"
            End If
            strSpecial = strSpecial & " | "
            i = i + 1
            If i Mod 10 = 0 Then
                strSpecial = strSpecial & "<br>"
            End If
            rsSpecial.MoveNext
        Loop
    End If
    rsSpecial.Close
    Set rsSpecial = Nothing
    GetSpecialList = strSpecial
End Function

'**************************************************
'��������GetClass_Option
'��  �ã���Ŀ�����˵�
'��  ����ShowType ---- ��ʾ����
'        CurrentID ---- ��ǰ��ĿID
'����ֵ����Ŀ�����˵�
'**************************************************
Function GetClass_Option(ShowType, CurrentID)
    Dim rsClass, sqlClass, strClass_Option, tmpDepth, i, ClassNum
    Dim arrShowLine(20)
    ClassNum = 1
    'CurrentID = PE_CLng(CurrentID)
    
    For i = 0 To UBound(arrShowLine)
        arrShowLine(i) = False
    Next
    sqlClass = "Select * from PE_Class where ChannelID=" & ChannelID & " order by RootID,OrderID"
    Set rsClass = Conn.Execute(sqlClass)
    If rsClass.BOF And rsClass.EOF Then
        strClass_Option = strClass_Option & "<option value=''>���������Ŀ</option>"
    Else
        Do While Not rsClass.EOF
            ClassNum = ClassNum + 1
            tmpDepth = rsClass("Depth")
            If rsClass("NextID") > 0 Then
                arrShowLine(tmpDepth) = True
            Else
                arrShowLine(tmpDepth) = False
            End If
            If ShowType = 1 Then
                If rsClass("ClassType") = 2 Then
                    strClass_Option = strClass_Option & "<option value=''"
                Else
                    strClass_Option = strClass_Option & "<option value='" & rsClass("ClassID") & "'"
                End If
                If AdminPurview = 2 And AdminPurview_Channel = 3 Then
                    If CheckPurview_Class(arrClass_Check, rsClass("ClassID")) = True Then
                        strClass_Option = strClass_Option & "style='background-color:#ff0000'"
                    End If
                End If
            ElseIf ShowType = 2 Then
                If rsClass("ClassType") = 2 Then
                    strClass_Option = strClass_Option & "<option value=''"
                Else
                    strClass_Option = strClass_Option & "<option value='" & rsClass("ClassID") & "'"
                End If
                If AdminPurview = 2 And AdminPurview_Channel = 3 Then
                    If CheckPurview_Class(arrClass_Manage, rsClass("ClassID")) = True Then
                        strClass_Option = strClass_Option & "style='background-color:#ff0000'"
                    End If
                End If
            ElseIf ShowType = 3 Then
                If rsClass("ClassType") = 2 Then
                    strClass_Option = strClass_Option & "<option value=''"
                Else
                    If rsClass("Child") > 0 And rsClass("EnableAdd") = False And FoundInArr(CurrentID, rsClass("ClassID"), ",") = False Then
                        strClass_Option = strClass_Option & "<option value='0'"
                    Else
                        strClass_Option = strClass_Option & "<option value='" & rsClass("ClassID") & "'"
                    End If
                End If
                If AdminPurview = 2 And AdminPurview_Channel = 3 Then
                    If CheckPurview_Class(arrClass_Input, rsClass("ClassID")) = True Then
                        strClass_Option = strClass_Option & "style='background-color:#ff0000'"
                    End If
                End If
            Else
                If rsClass("ClassType") = 2 Then
                    strClass_Option = strClass_Option & "<option value=''"
                Else
                    strClass_Option = strClass_Option & "<option value='" & rsClass("ClassID") & "'"
                End If
            End If
            If FoundInArr(CurrentID, rsClass("ClassID"), ",") Then
                strClass_Option = strClass_Option & " selected"
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
                strClass_Option = strClass_Option & "(��)"
            End If
            strClass_Option = strClass_Option & "</option>"
            ClassNum = ClassNum + 1
            rsClass.MoveNext
        Loop
    End If
    rsClass.Close
    Set rsClass = Nothing
    If ShowType = 3 And (AdminPurview = 1 Or (AdminPurview = 2 And AdminPurview_Channel < 3)) Then
        strClass_Option = strClass_Option & "<option value='-1'"
        If PE_CLng(CurrentID) = -1 Then strClass_Option = strClass_Option & " selected"
        strClass_Option = strClass_Option & ">��ָ���κ���Ŀ</option>"
    End If
    If ShowType = 0 And (AdminPurview = 1 Or (AdminPurview = 2 And AdminPurview_Channel < 3)) Then
        strClass_Option = strClass_Option & "<option value='-1'"
        If PE_CLng(CurrentID) = -1 Then strClass_Option = strClass_Option & " selected"
        strClass_Option = strClass_Option & ">δָ���κ���Ŀ</option>"
    End If
    GetClass_Option = strClass_Option
End Function

'**************************************************
'��������GetSkin_Option
'��  �ã���������˵�
'��  ����iSkinID ---- ���ID
'����ֵ����������˵�
'**************************************************
Function GetSkin_Option(iSkinID)
    Dim sqlSkin, rsSkin, strSkin
    If IsNull(iSkinID) Then iSkinID = 0
    strSkin = ""
    sqlSkin = "select * from PE_Skin"
    Set rsSkin = Conn.Execute(sqlSkin)
    If rsSkin.BOF And rsSkin.EOF Then
        strSkin = strSkin & "<option value=''>������ӷ��</option>"
    Else
        If iSkinID = 0 Then
            strSkin = strSkin & "<option value='0' selected>ϵͳĬ�Ϸ��</option>"
        Else
            strSkin = strSkin & "<option value='0'>ϵͳĬ�Ϸ��</option>"
        End If
        Do While Not rsSkin.EOF
            strSkin = strSkin & "<option value='" & rsSkin("SkinID") & "'"
            If rsSkin("SkinID") = iSkinID Then
                strSkin = strSkin & " selected"
            End If
            strSkin = strSkin & ">" & rsSkin("SkinName")
            If rsSkin("IsDefault") = True Then
                strSkin = strSkin & "��Ĭ�ϣ�"
            End If
            strSkin = strSkin & "</option>"
            rsSkin.MoveNext
        Loop
    End If
    rsSkin.Close
    Set rsSkin = Nothing
    GetSkin_Option = strSkin
End Function

'**************************************************
'��������GetTemplate_Option
'��  �ã�ģ�������˵�
'��  ����ChannelID ---- Ƶ��ID
'        TemplateType ---- ģ������
'        TemplateID---- ģ��ID
'����ֵ��ģ�������˵�
'**************************************************
Function GetTemplate_Option(ChannelID, TemplateType, TemplateID)
    Dim sqlTemplate, rsTemplate, strTemplate, strTemplateName
    strTemplate = ""
    If IsNull(TemplateID) Then TemplateID = 0
    Select Case TemplateType
        Case 1
            strTemplateName = "��ҳģ��"
        Case 2
            strTemplateName = "��Ŀģ��"
        Case 3
            strTemplateName = "����ҳģ��"
        Case 4
            strTemplateName = "ר��ҳģ��"
        Case Else
            strTemplateName = "ģ��"
    End Select
    If ChannelID = 0 And TemplateType = 4 Then TemplateType = 30
    sqlTemplate = "select * from PE_Template where ChannelID=" & ChannelID & " and TemplateType=" & TemplateType
    Set rsTemplate = Conn.Execute(sqlTemplate)
    If rsTemplate.BOF And rsTemplate.EOF Then
        strTemplate = strTemplate & "<option value=''>�������" & strTemplateName & "</option>"
    Else
        If TemplateID = 0 Then
            strTemplate = strTemplate & "<option value='0' selected>ϵͳĬ��" & strTemplateName & "</option>"
        Else
            strTemplate = strTemplate & "<option value='0'>ϵͳĬ��" & strTemplateName & "</option>"
        End If
        Do While Not rsTemplate.EOF
            strTemplate = strTemplate & "<option value='" & rsTemplate("TemplateID") & "'"
            If rsTemplate("TemplateID") = TemplateID Then
                strTemplate = strTemplate & " selected"
            End If
            strTemplate = strTemplate & ">" & rsTemplate("TemplateName")
            If rsTemplate("IsDefault") = True Then
                strTemplate = strTemplate & "��Ĭ�ϣ�"
            End If
            strTemplate = strTemplate & "</option>"
            rsTemplate.MoveNext
        Loop
    End If
    rsTemplate.Close
    Set rsTemplate = Nothing
    GetTemplate_Option = strTemplate
End Function

'**************************************************
'��������GetChannel_Option
'��  �ã�Ƶ�������˵�
'��  ����iModuleType ---- Ƶ������
'        iChannelID ---- Ƶ��ID
'����ֵ��Ƶ�������˵�Ŀ
'**************************************************
Function GetChannel_Option(iModuleType, iChannelID)
    Dim rsGetAdmin, rsChannel
    Dim strChannel
    Set rsGetAdmin = Conn.Execute("select * from PE_Admin where AdminName='" & AdminName & "'")
    Set rsChannel = Conn.Execute("select ChannelID,ChannelName,ChannelDir from PE_Channel  where ModuleType=" & iModuleType & " and Disabled=" & PE_False & " and ChannelType<=1 order by OrderID")
    Do While Not rsChannel.EOF
        If AdminPurview = 1 Or rsGetAdmin("AdminPurview_" & rsChannel("ChannelDir")) = 1 Then
            If rsChannel(0) = iChannelID Then
                strChannel = strChannel & "<option value='" & rsChannel(0) & "' selected>" & rsChannel(1) & "</option>"
            Else
                strChannel = strChannel & "<option value='" & rsChannel(0) & "'>" & rsChannel(1) & "</option>"
            End If
        End If
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    rsGetAdmin.Close
    Set rsGetAdmin = Nothing
    GetChannel_Option = strChannel
End Function

'=================================================
'��������EditCustom_Content
'��  �ã��༭��������
'=================================================
Sub EditCustom_Content(ByVal Action, ByVal Custom_Content, ByVal CustomType)

    Response.Write "<script language=""JavaScript"">" & vbCrLf
    Response.Write "<!--" & vbCrLf
    Response.Write "function setFileFileds(num){    " & vbCrLf
    Response.Write "    for(var i=1,str="""";i<=20;i++){" & vbCrLf
    Response.Write "        eval(""objFiles"" + i +"".style.display='none';"")" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    for(var i=1,str="""";i<=num;i++){" & vbCrLf
    Response.Write "        eval(""objFiles"" + i +"".style.display='';"")" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</script>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td width='150' class='tdbg5' align='center'><strong>����������Ŀ����</strong></td>"
    Response.Write "      <td>"
    If IsNull(Custom_Content) = True Then Custom_Content = ""

    Dim arrCustom, i, n, Custom_Num
    arrCustom = Split(Custom_Content, "{#$$$#}")
    Custom_Num = UBound(arrCustom) + 1

    Response.Write "      <select name=""Custom_Num"" onChange=""setFileFileds(this.value)"">" & vbCrLf
    For n = 1 To 20
        Response.Write "         <option value=""" & n & """ " & OptionValue(Custom_Num, n) & ">" & n & "</option>" & vbCrLf
    Next
    Response.Write "      </select>" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write " </td>"
    Response.Write "</tr>" & vbCrLf
    If Action = "Add" Then
        For i = 1 To 20
            Response.Write "    <tr class='tdbg' id=""objFiles" & i & """"
            If i > 1 Then
                Response.Write " style=""display:'none'"""
            End If
            Response.Write ">"
            Response.Write "      <td width='150' class='tdbg5'  ><center><strong>��������" & i & "��</strong></center><br>&nbsp;&nbsp;"
            Call EditCustomContentType(CustomType, i)
            Response.Write "      <td><TEXTAREA Name='Custom_Content" & i & "' ROWS='' COLS='' style='width:500px;height:100px'></TEXTAREA>"
            Response.Write "      </td></tr>" & vbCrLf
        Next
    Else
        For i = 0 To UBound(arrCustom)
            Response.Write "    <tr class='tdbg' id=""objFiles" & i + 1 & """ style=""display:''"">"
            Response.Write "      <td width='150' class='tdbg5' align='center'><center><strong>��������" & i + 1 & "��</strong></center><br>&nbsp;&nbsp;"
            Call EditCustomContentType(CustomType, i + 1)
            Response.Write "</td>"
            Response.Write "      <td><TEXTAREA Name='Custom_Content" & i + 1 & "' ROWS='' COLS='' style='width:500px;height:100px'>" & arrCustom(i) & "</TEXTAREA></td>"
            Response.Write "    </tr>" & vbCrLf
        Next
        Custom_Num = Custom_Num + 1
        For i = Custom_Num To 20
            Response.Write "    <tr class='tdbg' id=""objFiles" & i & """"
            If i > 1 Then
                Response.Write " style=""display:'none'"""
            End If
            Response.Write ">"
            Response.Write "      <td width='150' class='tdbg5' align='center'><center><strong>��������" & i & "��</strong></center><br>&nbsp;&nbsp;"
            Call EditCustomContentType(CustomType, i)
            Response.Write "</td>"
            Response.Write "      <td><TEXTAREA Name='Custom_Content" & i & "' ROWS='' COLS='' style='width:500px;height:100px'></TEXTAREA></td>"
            Response.Write "    </tr>" & vbCrLf
        Next
    End If
    Response.Write "  </tbody>" & vbCrLf
End Sub

'=================================================
'��������EditCustomContentType
'��  �ã�������������
'=================================================
Sub EditCustomContentType(ByVal CustomType, ByVal CustomNum)
    Select Case CustomType
    Case "Channel"
        Response.Write "��Ƶ��ģ��ҳ�����<Font color='blue'>{$Channel_Custom_Content" & CustomNum & "}" & vbCrLf
    Case "Class"
        Response.Write "����Ŀģ��ҳ�����<Font color='blue'>{$Class_Custom_Content" & CustomNum & "}" & vbCrLf
    Case "Special"
        Response.Write "��ר��ģ��ҳ�����<Font color='blue'>{$Special_Custom_Content" & CustomNum & "}" & vbCrLf
    Case Else
    End Select
    Response.Write "</font>����</td>"
End Sub

'**************************************************
'��������GetClass_Channel
'��  �ã���Ŀ�����˵�(�����Ȩ��)
'��  ����iChannelID ---- Ƶ��ID
'����ֵ����Ŀ�����˵�
'**************************************************
Function GetClass_Channel(iChannelID)
    Dim rsClass, sqlClass, strClass_Option, tmpDepth, i
    Dim arrShowLine(20)
    For i = 0 To UBound(arrShowLine)
        arrShowLine(i) = False
    Next
    sqlClass = "Select * from PE_Class where ChannelID=" & iChannelID & " order by RootID,OrderID"
    Set rsClass = Conn.Execute(sqlClass)
    If rsClass.BOF And rsClass.EOF Then
        strClass_Option = strClass_Option & "<option value=''>���������Ŀ</option>"
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
                If rsClass("Child") > 0 And rsClass("EnableAdd") = False Then
                    strClass_Option = strClass_Option & "<option value='0'"
                Else
                    strClass_Option = strClass_Option & "<option value='" & rsClass("ClassID") & "'"
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
                strClass_Option = strClass_Option & "(��)"
            End If
            strClass_Option = strClass_Option & "</option>"
            rsClass.MoveNext
        Loop
    End If
    rsClass.Close
    Set rsClass = Nothing
    strClass_Option = strClass_Option & "<option value='-1'>δָ���κ���Ŀ</option>"
    GetClass_Channel = strClass_Option
End Function

Function GetSpecialIDArr(ModuleType, ItemID)
    Dim rsInfo, arrSpecialID
    arrSpecialID = ""
    Set rsInfo = Conn.Execute("select SpecialID from PE_InfoS where ModuleType=" & ModuleType & " and ItemID=" & ItemID & "")
    Do While Not rsInfo.EOF
        If arrSpecialID = "" Then
            arrSpecialID = rsInfo(0)
        Else
            arrSpecialID = arrSpecialID & "," & rsInfo(0)
        End If
        rsInfo.MoveNext
    Loop
    rsInfo.Close
    Set rsInfo = Nothing
    GetSpecialIDArr = arrSpecialID
End Function

Sub SendEmailOfCheck(tUserName, rsItem)
    Dim SendMessageToInputer, SendEmailToInputer, MsgTitle, MsgContent
    SendMessageToInputer = Trim(Request.Form("SendMessageToInputer"))   '�Ƿ������Ͷ���
    SendEmailToInputer = Trim(Request.Form("SendEmailToInputer"))   '�Ƿ������Ͷ���
    MsgTitle = Trim(Request.Form("MsgTitle"))         '���ű���
    MsgContent = Trim(Request.Form("MsgContent"))     '��������

    MsgContent = ReplaceItemInfo(MsgContent, rsItem)
    If SendMessageToInputer = "Yes" Then
        Call SendMessage(tUserName, MsgTitle, MsgContent, UserName)
    End If
    If SendEmailToInputer = "Yes" Then
        Call SendCheckEmail(tUserName, MsgTitle, MsgContent, UserName)
    End If

End Sub


'**************************************************
'��������SendCheckEmail
'��  �ã��˸�ʱ������Email֪ͨ¼����
'��  ����InceptUser ----�û�����
'        Title ---- ����
'        Content ---- ����
'        SendUser ---- ��Email�Ĺ���Ա
'**************************************************
Sub SendCheckEmail(InceptUser, Title, Content, SendUser)
    If Content = "" Then
        Exit Sub
    End If
    Dim PE_Mail, ErrMsg, rsEmail, rsmaster
    Set rsEmail = Conn.Execute("select Email from PE_User where UserName='" & InceptUser & "'")
    If rsEmail.BOF And rsEmail.EOF Then

    Else
        Set PE_Mail = New SendMail
        ErrMsg = PE_Mail.Send(rsEmail(0), InceptUser, Title, Content, SendUser, WebmasterEmail, 3)
        Set PE_Mail = Nothing
    End If
    Set rsEmail = Nothing
End Sub

'**************************************************
'��������ReplaceItemInfo ��libinqq��
'��  �ã��Գ��ñ�ǩ�˸�˵������
'��  ����strContent ----�˸�˵��
'        ChannelShortName ---- ����Ƶ������
'����ֵ����������˸�˵��
'**************************************************
Function ReplaceItemInfo(strContent, rsItem)
    Dim strTemp
    strTemp = Replace(strContent, "{$ChannelShortName}", ChannelShortName)
    strTemp = Replace(strTemp, "{$Author}", rsItem("Author"))
    strTemp = Replace(strTemp, "{$CopyFrom}", rsItem("CopyFrom"))
    strTemp = Replace(strTemp, "{$Editor}", rsItem("Editor"))
    strTemp = Replace(strTemp, "{$Inputer}", rsItem("Inputer"))
    Select Case ModuleType
    Case 1
        strTemp = Replace(strTemp, "{$Title}", rsItem("Title"))
    Case 2
        strTemp = Replace(strTemp, "{$Title}", rsItem("SoftName"))
        strTemp = Replace(strTemp, "{$SoftName}", rsItem("SoftName"))
    Case 3
        strTemp = Replace(strTemp, "{$Title}", rsItem("PhotoName"))
        strTemp = Replace(strTemp, "{$PhotoName}", rsItem("PhotoName"))
    End Select
    ReplaceItemInfo = strTemp
End Function

'**************************************************
'��������WriteFieldHTML
'��  �ã���ʾ�Զ����ֶα�
'��  ����FieldName ----�Զ����ֶ�����
'        Title ---- ����
'        Tips ---- ������ʾ
'        FieldType ---- �ֶ�����  1--�����ı�  2--�����ı�  3--�����б�  4--ͼƬ  5--�ļ�  6--����  7--����
'        strValue ---- Ĭ��ֵ
'        Options ---- �б���Ŀ
'        EnableNull ---- �Ƿ����Ϊ��
'����ֵ���Զ����ֶα�
'**************************************************
Sub WriteFieldHTML(FieldName, Title, Tips, FieldType, strValue, Options, EnableNull)
    Dim FieldUpload, ChannelUpload, UserUpload,rsFieldUpload,sqlFieldUpload   
    Select Case FieldType
    Case 4,5
        FieldUpload = True		
        ChannelUpload = Conn.Execute("Select EnableUploadFile from PE_Channel where ChannelID="&ChannelID)(0) 
        If  ChannelUpload = False Then FieldUpload = False
        If UserName<>"" Then   
            sqlFieldUpload = "SELECT U.UserID,U.SpecialPermission,U.UserSetting,G.GroupSetting FROM PE_User U inner join PE_UserGroup G on U.GroupID=G.GroupID WHERE"
            sqlFieldUpload = sqlFieldUpload & " UserName='" & UserName & "'" 
            Set rsFieldUpload = Conn.Execute(sqlFieldUpload)
            If rsFieldUpload.BOF And rsFieldUpload.EOF Then
                FieldUpload = False
            Else
                If rsFieldUpload("SpecialPermission") = True Then
                    UserSetting = Split(Trim(rsFieldUpload("UserSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
                Else
                    UserSetting = Split(Trim(rsFieldUpload("GroupSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
                End If
                If CBool(PE_CLng(UserSetting(9))) = False Then
                    FieldUpload = False
                End If
            End If
            Set rsFieldUpload = nothing			 
        End If			               			
    End Select
    Dim strEnableNull
    If EnableNull = False Then
        strEnableNull = " <font color='#FF0000'>*</font>"
    End If
    Response.Write "<tr class='tdbg'><td width='120' align='right' class='tdbg5'>" & Title & "��</b><td colspan='5'>"
    Select Case FieldType
    Case 1,8    '�����ı���
        Response.Write "<input type='text' name='" & FieldName & "' size='80' maxlength='255' value='" & strValue & "'>" & strEnableNull
    Case 2,9    '�����ı���
        Response.Write "<textarea name='" & FieldName & "' cols='80' rows='10'>" & strValue & "</textarea>" & strEnableNull
    Case 3    '�����б�
        Response.Write "<select name='" & FieldName & "'>"
        Dim arrOptions, i
        arrOptions = Split(Options, vbCrLf)
        For i = 0 To UBound(arrOptions)
            Response.Write "<option value='" & arrOptions(i) & "'"
            If arrOptions(i) = strValue Then Response.Write " selected"
            Response.Write ">" & arrOptions(i) & "</option>"
        Next
        Response.Write "</select>" & strEnableNull
    Case 4   'ͼƬ  					
        If strValue = "" Then
            Response.Write "<input type='text' id='"&FieldName&"' name='"&FieldName&"'  size='45' maxlength='255' value='http://'><br>" & strEnableNull
        Else
            Response.Write "<input type='text' name='" & FieldName & "' id='" & FieldName & "' size='45' maxlength='255' value='" & strValue & "'><br>" & strEnableNull
        End If
        If PE_CBool(FieldUpload) = True Then		
            Response.Write "<iframe style='top:2px;' id='uploadPhoto' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=fieldpic&FieldName="& FieldName &"' frameborder=0 scrolling=no width='650' height='25'></iframe>"
        End If				
    Case 5   '�ļ�
        If strValue = "" Then
            Response.Write "<input type='text' id='"&FieldName&"' name='"&FieldName&"'  size='45' maxlength='255' value='http://'><br>" & strEnableNull
        Else
            Response.Write "<input type='text' name='" & FieldName & "' id='" & FieldName & "' size='45' maxlength='255' value='" & strValue & "'><br>" & strEnableNull
        End If
        If PE_CBool(FieldUpload) = True Then			
            Response.Write "<iframe style='top:2px' id='uploadsoft' src='upload.asp?ChannelID=" & ChannelID & "&dialogtype=fieldsoft&FieldName="& FieldName &"' frameborder=0 scrolling=no width='650' height='25'></iframe>"	
        End If
				
    Case 6    '����
        If strValue = "" Then
            Response.Write "<input type='text' name='" & FieldName & "' size='20' maxlength='20' value='" & Now() & "'>" & strEnableNull
        Else
            Response.Write "<input type='text' name='" & FieldName & "' size='20' maxlength='20' value='" & strValue & "'>" & strEnableNull
        End If
    Case 7    '����
        If strValue = "" Then
            Response.Write "<input type='text' name='" & FieldName & "'  onkeyup=""value=value.replace(/[^\d]/g,'')"" size='20' maxlength='20' value='0'>" & strEnableNull
        Else
            Response.Write "<input type='text' name='" & FieldName & "'  onkeyup=""value=value.replace(/[^\d]/g,'')""  size='20' maxlength='20' value='" & PE_Clng(strValue) & "'>" & strEnableNull
        End If		
    End Select
    If IsNull(Tips) = False And Tips <> ""  and (FieldType <> 4 and FieldType <> 5) Then
        Response.Write "<br>" & PE_HTMLEncode(Tips)
    End If
    Response.Write "</td></tr>"
End Sub

Sub WriteBatchFieldHTML(FieldName, Title, Tips, FieldType, strValue, Options, EnableNull)
    Response.Write "<tr class='tdbg'>"
    Response.Write "<td width='30' align='center' class='tdbg5'><input type='checkbox' name='Modify" & FieldName & "' value='Yes'></td>"
    Response.Write "<td width='120' align='right' class='tdbg5'>" & Title & "��</td>"
    Response.Write "<td width='450'>"
    Dim strEnableNull
    If EnableNull = False Then
        strEnableNull = " <font color='#FF0000'>*</font>"
    End If
    Select Case FieldType
    Case 1,8    '�����ı���
        Response.Write "<input type='text' name='" & FieldName & "' size='65' maxlength='255' value='" & strValue & "'>" & strEnableNull
    Case 2,9    '�����ı���
        Response.Write "<textarea name='" & FieldName & "' cols='55' rows='10'>" & strValue & "</textarea>" & strEnableNull
    Case 3    '�����б�
        Response.Write "<select name='" & FieldName & "'>"
        Dim arrOptions, i
        arrOptions = Split(Options, vbCrLf)
        For i = 0 To UBound(arrOptions)
            Response.Write "<option value='" & arrOptions(i) & "'"
            If arrOptions(i) = strValue Then Response.Write " selected"
            Response.Write ">" & arrOptions(i) & "</option>"
        Next
        Response.Write "</select>" & strEnableNull
    Case 4, 5   'ͼƬ���ļ�
        If strValue = "" Then
            Response.Write "<input type='text' name='" & FieldName & "' size='40' maxlength='255' value='http://'>" & strEnableNull
        Else
            Response.Write "<input type='text' name='" & FieldName & "' size='40' maxlength='255' value='" & strValue & "'>" & strEnableNull
        End If
    Case 6    '����
        If strValue = "" Then
            Response.Write "<input type='text' name='" & FieldName & "' size='20' maxlength='20' value='" & Now() & "'>" & strEnableNull
        Else
            Response.Write "<input type='text' name='" & FieldName & "' size='20' maxlength='20' value='" & strValue & "'>" & strEnableNull
        End If
    End Select
    If IsNull(Tips) = False And Tips <> "" Then
        Response.Write "<br>" & PE_HTMLEncode(Tips)
    End If
    Response.Write "</td></tr>"
End Sub

Sub SaveVote()
    Dim UseVote, VoteTitle, VoteType, VoteTime, EndTime
    Dim sql, rsVote, rsVote2, i
    UseVote = Trim(Request.Form("ShowVote"))
    If UseVote = "yes" Then
        VoteTitle = Trim(Request.Form("VoteTitle"))
        VoteType = Trim(Request.Form("VoteType"))
        VoteTime = Trim(Request.Form("BeginTime"))
        EndTime = Trim(Request.Form("EndTime"))
        Set rsVote = Server.CreateObject("adodb.recordset")
        If Action = "SaveAdd" Or Action = "SaveModifyAsAdd" Then
            sql = "select top 1 * from PE_Vote"
            rsVote.Open sql, Conn, 1, 3
            rsVote.addnew
            rsVote("Title") = VoteTitle
            For i = 1 To 8
                rsVote("select" & i) = Trim(Request("select" & i))
                If Request("answer" & i) = "" Then
                    rsVote("answer" & i) = 0
                Else
                    rsVote("answer" & i) = PE_CLng(Request("answer" & i))
                End If
            Next
            rsVote("VoteTime") = VoteTime
            rsVote("EndTime") = EndTime
            rsVote("VoteType") = VoteType
            rsVote("IsSelected") = True
            rsVote("ChannelID") = ChannelID
            rsVote("IsItem") = True
            rsVote("VoteNum") = 0			
            rsVote.Update
            rsVote.Close
            Set rsVote2 = Conn.Execute("select max(ID) from PE_Vote")
            VoteID = rsVote2(0)
        Else
            Select Case ModuleType
            Case 1
                Set rsVote2 = Conn.Execute("select VoteID from PE_Article where ArticleID=" & ArticleID)
            Case 2
                Set rsVote2 = Conn.Execute("select VoteID from PE_Soft where SoftID=" & SoftID)
            Case 3
                Set rsVote2 = Conn.Execute("select VoteID from PE_Photo where PhotoID=" & PhotoID)
            Case 5
                Set rsVote2 = Conn.Execute("select VoteID from PE_Product where ProductID=" & ProductID)
            End Select
            If rsVote2("VoteID") = 0 Then
                sql = "select top 1 * from PE_Vote"
                rsVote.Open sql, Conn, 1, 3
                rsVote.addnew
                rsVote("Title") = VoteTitle
                For i = 1 To 8
                    rsVote("select" & i) = Trim(Request("select" & i))
                    If Request("answer" & i) = "" Then
                        rsVote("answer" & i) = 0
                    Else
                        rsVote("answer" & i) = PE_CLng(Request("answer" & i))
                    End If
                Next
                rsVote("VoteTime") = VoteTime
                rsVote("EndTime") = EndTime
                rsVote("VoteType") = VoteType
                rsVote("IsSelected") = True
                rsVote("ChannelID") = ChannelID
                rsVote("IsItem") = True
                rsVote.Update
                rsVote.Close
                Set rsVote2 = Conn.Execute("select max(ID) from PE_Vote")
                VoteID = rsVote2(0)
            Else
                sql = "select top 1 * from PE_Vote where ID=" & rsVote2("VoteID")
                rsVote.Open sql, Conn, 1, 3
                rsVote("Title") = VoteTitle
                For i = 1 To 8
                    rsVote("select" & i) = Trim(Request("select" & i))
                    If Request("answer" & i) = "" Then
                        rsVote("answer" & i) = 0
                    Else
                        rsVote("answer" & i) = PE_CLng(Request("answer" & i))
                    End If
                Next
                rsVote("VoteTime") = VoteTime
                rsVote("EndTime") = EndTime
                rsVote("VoteType") = VoteType
                rsVote("IsSelected") = True
                rsVote("ChannelID") = ChannelID
                rsVote("IsItem") = True
                rsVote.Update
                rsVote.Close
                VoteID = rsVote2(0)
            End If
        End If
    Else
        If Action = "SaveModify" Then
            Select Case ModuleType
            Case 1
                Set rsVote2 = Conn.Execute("select VoteID from PE_Article where ArticleID=" & ArticleID)
            Case 2
                Set rsVote2 = Conn.Execute("select VoteID from PE_Soft where SoftID=" & SoftID)
            Case 3
                Set rsVote2 = Conn.Execute("select VoteID from PE_Photo where PhotoID=" & PhotoID)
            Case 5
                Set rsVote2 = Conn.Execute("select VoteID from PE_Product where ProductID=" & ProductID)
            End Select
            If rsVote2(0) > 0 Then
               Set rsVote = Conn.Execute("delete from PE_Vote where ID=" & rsVote2(0))
            End If
        End If
        VoteID = 0
    End If
    Set rsVote = Nothing
    Set rsVote2 = Nothing
End Sub

%>
