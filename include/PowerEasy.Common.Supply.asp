<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'************************************
'函数名：getCommandType
'作  用：供求信息设定推荐的提示和扣除标准
'参  数：iType ----- 数值类型，看那个被选中
'        CommandChannelPoint -------- 频道推荐要扣除的点数
'返回值: 推荐部分所要显示的内容
'************************************
Function getCommandType(ByVal iType, ByVal CommandChannelPoint, ByVal PointName, ByVal PointUnit)
    Select Case iType
        Case 1
            getCommandType = "<Table>" & _
                             "<tr><td><INPUT TYPE='radio' Value='0' NAME='CommandType' >不推荐</td></tr>" & _
                             "<tr><td><INPUT TYPE='radio' NAME='CommandType' Value='1' Checked>频道推荐&nbsp;<INPUT TYPE=text' NAME='CommandChanneldays' Maxlength='4' size='4'>&nbsp;天&nbsp;<font color=red>注意：</font><font color=#0000FF>频道推荐扣除的" & PointName & "标准是：<font color=red>" & CommandChannelPoint & "</font>&nbsp;" & PointUnit & "/天</font></td></tr>" & _
                             "<tr><td><INPUT TYPE='radio' Value='2' NAME='CommandType'>栏目推荐&nbsp;<INPUT TYPE=text' NAME='CommandClassdays' Maxlength='4' size='4'>&nbsp;天&nbsp;<font color=red>注意：</font><font color=#0000FF>当前栏目推荐扣除" & PointName & "的标准是：<font color=red><span id='CommandClassPoint'></Span></font>" & PointUnit & "/天</font></td></tr></Table>"
        Case 2
            getCommandType = "<Table>" & _
                             "<tr><td><INPUT TYPE='radio' Value='0' NAME='CommandType' >不推荐</td></tr>" & _
                             "<tr><td><INPUT TYPE='radio' NAME='CommandType' Value='1' >频道推荐&nbsp;<INPUT TYPE=text' NAME='CommandChanneldays' Maxlength='4' size='4'>&nbsp;天&nbsp;<font color=red>注意：</font><font color=#0000FF>频道推荐扣除的" & PointName & "标准是：<font color=red>" & CommandChannelPoint & "</font>&nbsp;" & PointUnit & "/天</font></td></tr>" & _
                             "<tr><td><INPUT TYPE='radio' Value='2' NAME='CommandType' Checked>栏目推荐&nbsp;<INPUT TYPE=text' NAME='CommandClassdays' Maxlength='4' size='4'>&nbsp;天&nbsp;<font color=red>注意：</font><font color=#0000FF>当前栏目推荐扣除" & PointName & "的标准是：<font color=red><span id='CommandClassPoint'></Span></font>" & PointUnit & "/天</font></td></tr></Table>"
        Case Else
            getCommandType = "<Table>" & _
                             "<tr><td><INPUT TYPE='radio' Value='0' NAME='CommandType' Checked>不推荐</td></tr>" & _
                             "<tr><td><INPUT TYPE='radio' NAME='CommandType' Value='1' >频道推荐&nbsp;<INPUT TYPE=text' NAME='CommandChanneldays' Maxlength='4' size='4'>&nbsp;天&nbsp;<font color=red>注意：</font><font color=#0000FF>频道推荐扣除的" & PointName & "标准是：<font color=red>" & CommandChannelPoint & " </font>&nbsp;" & PointUnit & "/天</font></td></tr>" & _
                             "<tr><td><INPUT TYPE='radio' Value='2' NAME='CommandType'>栏目推荐&nbsp;<INPUT TYPE=text' NAME='CommandClassdays' Maxlength='4' size='4'>&nbsp;天&nbsp;<font color=red>注意：</font><font color=#0000FF>当前栏目推荐扣除" & PointName & "的标准是：<font color=red><span id='CommandClassPoint'></Span></font>" & PointUnit & "/天</font></td></tr></Table>"
    End Select
End Function

'***************************
'函数名：SetAjax
'作  用：获得栏目推荐的点数
'参  数：无
'返回值：无
'***************************
Sub SetAjax()
    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "    var http_request = false;" & vbCrLf
    Response.Write "    function InitRequest() {//初始化、指定处理函数、发送请求的函数" & vbCrLf
    Response.Write "        http_request = false;" & vbCrLf
    Response.Write "        //开始初始化XMLHttpRequest对象" & vbCrLf
    Response.Write "        if(window.XMLHttpRequest) { //Mozilla 浏览器" & vbCrLf
    Response.Write "            http_request = new XMLHttpRequest();" & vbCrLf
    Response.Write "            if (http_request.overrideMimeType) {//设置MiME类别" & vbCrLf
    Response.Write "                http_request.overrideMimeType('text/xml');" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        else if (window.ActiveXObject) { // IE浏览器" & vbCrLf
    Response.Write "            try {" & vbCrLf
    Response.Write "                http_request = new ActiveXObject('Msxml2.XMLHTTP');" & vbCrLf
    Response.Write "            } catch (e) {" & vbCrLf
    Response.Write "                try {" & vbCrLf
    Response.Write "                    http_request = new ActiveXObject('Microsoft.XMLHTTP');" & vbCrLf
    Response.Write "                } catch (e) {}" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        if (!http_request) { // 异常，创建对象实例失败" & vbCrLf
    Response.Write "            window.alert('不能创建XMLHttpRequest对象实例.');" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        " & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    //设定初始值" & vbCrLf
    Response.Write "    function setBackValue(url)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        InitRequest();" & vbCrLf
    Response.Write "        http_request.onreadystatechange = function()" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            if (http_request.readyState == 4) " & vbCrLf
    Response.Write "            { // 判断对象状态" & vbCrLf
    Response.Write "                if (http_request.status == 200) " & vbCrLf
    Response.Write "                { // 信息已经成功返回，开始处理信息 " & vbCrLf
    Response.Write "                    document.getElementById('CommandClassPoint').innerHTML=http_request.responseText;" & vbCrLf
    Response.Write "                } else { //页面不正常" & vbCrLf
    Response.Write "                    alert('您所请求的页面有异常。');" & vbCrLf
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
'函 数 名：GetSupplyInfo_Radio
'作    用：获得相关的表单，从语言包内读取相关的选项，然后生成单选项
'参    数：iType    --- 被选中的单选表单，为空时则选中第一个
'        FormName --- 被生成表单的名字
'        NodeName --- 语言包中相关节点的名字
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
'函 数 名：GetSupplyInfoType
'作    用：返回的信息类型
'参    数：iType    ---- 被选中的要显示的类型
'          NodeName ---- 语言包中节点的名字
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
'获得信息的有效期
'***************************************
Function GetSupplyPeriod_Select(ByVal iType)
    If iType < 0 Or iType > 90 Or iType = "" Then
        iType = 10
    End If
    Select Case iType
        Case -1
            GetSupplyPeriod_Select = "<Select Name='SupplyPeriod'><option Value='10' >10天</option><option Value='20'>20天</option><option Value='30'>一个月</option><option Value='90'>三个月</option><option Value='-1' selected>长期有效</option></Select>"
        Case 10
            GetSupplyPeriod_Select = "<Select Name='SupplyPeriod'><option Value='10' selected>10天</option><option Value='20'>20天</option><option Value='30'>一个月</option><option Value='90'>三个月</option><option Value='-1' >长期有效</option></Select>"
        Case 20
            GetSupplyPeriod_Select = "<Select Name='SupplyPeriod'><option Value='10' >10天</option><option Value='20' selected >20天</option><option Value='30'>一个月</option><option Value='90'>三个月</option><option Value='-1' >长期有效</option></Select>"
        Case 30
            GetSupplyPeriod_Select = "<Select Name='SupplyPeriod'><option Value='10' >10天</option><option Value='20'>20天</option><option Value='30' selected >一个月</option><option Value='90'>三个月</option><option Value='-1'>长期有效</option></Select>"
        Case 90
            GetSupplyPeriod_Select = "<Select Name='SupplyPeriod'><option Value='10' >10天</option><option Value='20'>20天</option><option Value='30'>一个月</option><option Value='90' selected >三个月</option><option Value='-1' >长期有效</option></Select>"
    End Select
End Function

'******************************************
'获得信息状态
'*******************************************
Function GetSupplyPeriod(ByVal UpdateTime, ByVal Period)
    If PE_CLng(Period) = -1 Then
        GetSupplyPeriod = "长期有效"
    Else
        If DateDiff("d", PE_CDate(UpdateTime), Now()) > Period Then
            GetSupplyPeriod = "过期"
        End If
    End If
End Function

Function GetSupplyStatus(ByVal iType)
    Select Case PE_CLng(iType)
        Case 0
            GetSupplyStatus = "未审核"
        Case 1
            GetSupplyStatus = "已审核"
    End Select
End Function
%>
