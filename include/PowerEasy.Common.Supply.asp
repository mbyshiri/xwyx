<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

'************************************
'��������getCommandType
'��  �ã�������Ϣ�趨�Ƽ�����ʾ�Ϳ۳���׼
'��  ����iType ----- ��ֵ���ͣ����Ǹ���ѡ��
'        CommandChannelPoint -------- Ƶ���Ƽ�Ҫ�۳��ĵ���
'����ֵ: �Ƽ�������Ҫ��ʾ������
'************************************
Function getCommandType(ByVal iType, ByVal CommandChannelPoint, ByVal PointName, ByVal PointUnit)
    Select Case iType
        Case 1
            getCommandType = "<Table>" & _
                             "<tr><td><INPUT TYPE='radio' Value='0' NAME='CommandType' >���Ƽ�</td></tr>" & _
                             "<tr><td><INPUT TYPE='radio' NAME='CommandType' Value='1' Checked>Ƶ���Ƽ�&nbsp;<INPUT TYPE=text' NAME='CommandChanneldays' Maxlength='4' size='4'>&nbsp;��&nbsp;<font color=red>ע�⣺</font><font color=#0000FF>Ƶ���Ƽ��۳���" & PointName & "��׼�ǣ�<font color=red>" & CommandChannelPoint & "</font>&nbsp;" & PointUnit & "/��</font></td></tr>" & _
                             "<tr><td><INPUT TYPE='radio' Value='2' NAME='CommandType'>��Ŀ�Ƽ�&nbsp;<INPUT TYPE=text' NAME='CommandClassdays' Maxlength='4' size='4'>&nbsp;��&nbsp;<font color=red>ע�⣺</font><font color=#0000FF>��ǰ��Ŀ�Ƽ��۳�" & PointName & "�ı�׼�ǣ�<font color=red><span id='CommandClassPoint'></Span></font>" & PointUnit & "/��</font></td></tr></Table>"
        Case 2
            getCommandType = "<Table>" & _
                             "<tr><td><INPUT TYPE='radio' Value='0' NAME='CommandType' >���Ƽ�</td></tr>" & _
                             "<tr><td><INPUT TYPE='radio' NAME='CommandType' Value='1' >Ƶ���Ƽ�&nbsp;<INPUT TYPE=text' NAME='CommandChanneldays' Maxlength='4' size='4'>&nbsp;��&nbsp;<font color=red>ע�⣺</font><font color=#0000FF>Ƶ���Ƽ��۳���" & PointName & "��׼�ǣ�<font color=red>" & CommandChannelPoint & "</font>&nbsp;" & PointUnit & "/��</font></td></tr>" & _
                             "<tr><td><INPUT TYPE='radio' Value='2' NAME='CommandType' Checked>��Ŀ�Ƽ�&nbsp;<INPUT TYPE=text' NAME='CommandClassdays' Maxlength='4' size='4'>&nbsp;��&nbsp;<font color=red>ע�⣺</font><font color=#0000FF>��ǰ��Ŀ�Ƽ��۳�" & PointName & "�ı�׼�ǣ�<font color=red><span id='CommandClassPoint'></Span></font>" & PointUnit & "/��</font></td></tr></Table>"
        Case Else
            getCommandType = "<Table>" & _
                             "<tr><td><INPUT TYPE='radio' Value='0' NAME='CommandType' Checked>���Ƽ�</td></tr>" & _
                             "<tr><td><INPUT TYPE='radio' NAME='CommandType' Value='1' >Ƶ���Ƽ�&nbsp;<INPUT TYPE=text' NAME='CommandChanneldays' Maxlength='4' size='4'>&nbsp;��&nbsp;<font color=red>ע�⣺</font><font color=#0000FF>Ƶ���Ƽ��۳���" & PointName & "��׼�ǣ�<font color=red>" & CommandChannelPoint & " </font>&nbsp;" & PointUnit & "/��</font></td></tr>" & _
                             "<tr><td><INPUT TYPE='radio' Value='2' NAME='CommandType'>��Ŀ�Ƽ�&nbsp;<INPUT TYPE=text' NAME='CommandClassdays' Maxlength='4' size='4'>&nbsp;��&nbsp;<font color=red>ע�⣺</font><font color=#0000FF>��ǰ��Ŀ�Ƽ��۳�" & PointName & "�ı�׼�ǣ�<font color=red><span id='CommandClassPoint'></Span></font>" & PointUnit & "/��</font></td></tr></Table>"
    End Select
End Function

'***************************
'��������SetAjax
'��  �ã������Ŀ�Ƽ��ĵ���
'��  ������
'����ֵ����
'***************************
Sub SetAjax()
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "    var http_request = false;" & vbCrLf
    Response.Write "    function InitRequest() {//��ʼ����ָ������������������ĺ���" & vbCrLf
    Response.Write "        http_request = false;" & vbCrLf
    Response.Write "        //��ʼ��ʼ��XMLHttpRequest����" & vbCrLf
    Response.Write "        if(window.XMLHttpRequest) { //Mozilla �����" & vbCrLf
    Response.Write "            http_request = new XMLHttpRequest();" & vbCrLf
    Response.Write "            if (http_request.overrideMimeType) {//����MiME���" & vbCrLf
    Response.Write "                http_request.overrideMimeType('text/xml');" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        else if (window.ActiveXObject) { // IE�����" & vbCrLf
    Response.Write "            try {" & vbCrLf
    Response.Write "                http_request = new ActiveXObject('Msxml2.XMLHTTP');" & vbCrLf
    Response.Write "            } catch (e) {" & vbCrLf
    Response.Write "                try {" & vbCrLf
    Response.Write "                    http_request = new ActiveXObject('Microsoft.XMLHTTP');" & vbCrLf
    Response.Write "                } catch (e) {}" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        if (!http_request) { // �쳣����������ʵ��ʧ��" & vbCrLf
    Response.Write "            window.alert('���ܴ���XMLHttpRequest����ʵ��.');" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        " & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    //�趨��ʼֵ" & vbCrLf
    Response.Write "    function setBackValue(url)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        InitRequest();" & vbCrLf
    Response.Write "        http_request.onreadystatechange = function()" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            if (http_request.readyState == 4) " & vbCrLf
    Response.Write "            { // �ж϶���״̬" & vbCrLf
    Response.Write "                if (http_request.status == 200) " & vbCrLf
    Response.Write "                { // ��Ϣ�Ѿ��ɹ����أ���ʼ������Ϣ " & vbCrLf
    Response.Write "                    document.getElementById('CommandClassPoint').innerHTML=http_request.responseText;" & vbCrLf
    Response.Write "                } else { //ҳ�治����" & vbCrLf
    Response.Write "                    alert('���������ҳ�����쳣��');" & vbCrLf
    Response.Write "                }" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }       " & vbCrLf
    Response.Write "        http_request.open('GET',url,true);" & vbCrLf
    Response.Write "        http_request.send(null);" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    " & vbCrLf
    Response.Write "setBackValue('getClassCommandValue.asp?ClassID='+document.myform.ClassID.value);"
    Response.Write "</script>" & vbCrLf
End Sub

'********************************************
'�� �� ����GetSupplyInfo_Radio
'��    �ã������صı��������԰��ڶ�ȡ��ص�ѡ�Ȼ�����ɵ�ѡ��
'��    ����iType    --- ��ѡ�еĵ�ѡ����Ϊ��ʱ��ѡ�е�һ��
'        FormName --- �����ɱ�������
'        NodeName --- ���԰�����ؽڵ������
'******************************************
Function GetSupplyInfo_Radio(ByVal iType, ByVal FormName, ByVal NodeName, ByVal SupplyTypeNum)
    Dim LangRoot, i, strTemp, ShowLength
    Set LangRoot = XmlDoc.selectNodes(NodeName)
    iType = PE_CLng(iType)
    SupplyTypeNum = PE_CLng(SupplyTypeNum)
    If SupplyTypeNum >= LangRoot.Length Or SupplyTypeNum <= 0 Then
        ShowLength = LangRoot.Length - 1
    Else
        ShowLength = SupplyTypeNum
    End If
    For i = 0 To ShowLength
        strTemp = strTemp & "<INPUT TYPE='radio' NAME='" & FormName & "' Value='" & i & "'"
        If iType = i Then strTemp = strTemp & " Checked "
        strTemp = strTemp & ">" & LangRoot(i).text & "&nbsp;&nbsp;"
    Next
    Set LangRoot = Nothing
    GetSupplyInfo_Radio = strTemp
End Function
'***************************************
'�� �� ����GetSupplyInfoType
'��    �ã����ص���Ϣ����
'��    ����iType    ---- ��ѡ�е�Ҫ��ʾ������
'          NodeName ---- ���԰��нڵ������
'************************************
Function GetSupplyInfoType(ByVal iType, ByVal NodeName)
    Dim LangRoot
    iType = PE_CLng(iType)
    Set LangRoot = XmlDoc.selectNodes(NodeName)
    GetSupplyInfoType = LangRoot(iType).text
    Set LangRoot = Nothing
End Function

Function GetTypePoint(ByVal iType, ByVal NodeName, ByVal strAttribute)
    Dim LangRoot, i, strTemp
    Set LangRoot = XmlDoc.selectNodes(NodeName)
    'strTemp = LangRoot(iType).getAttribute(strAttribute)
    Set LangRoot = Nothing
    GetTypePoint = strTemp
End Function
'****************************************
'�����Ϣ����Ч��
'***************************************
Function GetSupplyPeriod_Select(ByVal iType)
    If iType < 0 Or iType > 90 Or iType = "" Then
        iType = 10
    End If
    Select Case iType
        Case -1
            GetSupplyPeriod_Select = "<Select Name='SupplyPeriod'><option Value='10' >10��</option><option Value='20'>20��</option><option Value='30'>һ����</option><option Value='90'>������</option><option Value='-1' selected>������Ч</option></Select>"
        Case 10
            GetSupplyPeriod_Select = "<Select Name='SupplyPeriod'><option Value='10' selected>10��</option><option Value='20'>20��</option><option Value='30'>һ����</option><option Value='90'>������</option><option Value='-1' >������Ч</option></Select>"
        Case 20
            GetSupplyPeriod_Select = "<Select Name='SupplyPeriod'><option Value='10' >10��</option><option Value='20' selected >20��</option><option Value='30'>һ����</option><option Value='90'>������</option><option Value='-1' >������Ч</option></Select>"
        Case 30
            GetSupplyPeriod_Select = "<Select Name='SupplyPeriod'><option Value='10' >10��</option><option Value='20'>20��</option><option Value='30' selected >һ����</option><option Value='90'>������</option><option Value='-1'>������Ч</option></Select>"
        Case 90
            GetSupplyPeriod_Select = "<Select Name='SupplyPeriod'><option Value='10' >10��</option><option Value='20'>20��</option><option Value='30'>һ����</option><option Value='90' selected >������</option><option Value='-1' >������Ч</option></Select>"
    End Select
End Function

'******************************************
'�����Ϣ״̬
'*******************************************
Function GetSupplyPeriod(ByVal UpdateTime, ByVal Period)
    If PE_CLng(Period) = -1 Then
        GetSupplyPeriod = "������Ч"
    Else
        If DateDiff("d", PE_CDate(UpdateTime), Now()) > Period Then
            GetSupplyPeriod = "����"
        End If
    End If
End Function

Function GetSupplyStatus(ByVal iType)
    Select Case PE_CLng(iType)
        Case 0
            GetSupplyStatus = "δ���"
        Case 1
            GetSupplyStatus = "�����"
    End Select
End Function
%>
