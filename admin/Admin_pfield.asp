<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

response.Expires = -1
response.ExpiresAbsolute = Now() - 1
response.Expires = 0
response.CacheControl = "no-cache"

Dim fieldname, num, dbname, dbtype, isknow

fieldname = Trim(Request("fieldname"))
num = Trim(Request("num"))
dbname = Trim(Request("dbname"))
If dbname = "" Then dbname = 0
dbtype = Trim(Request("dbtype"))
If dbtype = "" Then dbtype = 0
isknow = False
%>
<html>
<head>
<title>�ֶ���������</title>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<link href='Admin_Style.css' rel='stylesheet' type='text/css'>
<script language = 'JavaScript'>
function changemode(){
    var dbname=document.myform.ftype.value;
    if(dbname=='Text'){
    input1.style.display='';
    }else{
    input1.style.display='none';
    }
    if(dbname=='Num'){
    input2.style.display='';
    }else{
    input2.style.display='none';
    }
    if(dbname=='Time'){
    input3.style.display='';
    }else{
    input3.style.display='none';
    }
    if(dbname=='yn'){
    input4.style.display='';
    }else{
    input4.style.display='none';
    }
    if(dbname=='GetUrl'|dbname=='GetClass'|dbname=='GetSpecil'|dbname=='GetChannel'){
    input5.style.display='';
    }else{
    input5.style.display='none';
    }
}
function changetime(){
    var dbname=document.myform.Timetype.value;
    if(dbname=='3'){
    document.myform.Timemb.value="2";
    }else{
        document.myform.Timemb.value="{year}��{month}��{day}��";
    }
}
function submitdate(){
    var dbname=document.myform.ftype.value;
    if(dbname=='Text'){
        for (var i=0;i<document.myform.CatType.length;i++){
            if (document.myform.CatType[i].checked){
                var cattype=document.myform.CatType[i].value;
        }
        }
        dbname="{$Field(" + document.myform.Fieldnum.value + "," + dbname + "," + document.myform.CatNum.value + "," + document.myform.OutType2.value + "," + cattype + ")}";
    }
    if(dbname=='Num'){
        for (var i=0;i<document.myform.OutType.length;i++){
            if (document.myform.OutType[i].checked){
                var cattype=document.myform.OutType[i].value;
        }
        }
    if (cattype=='2'){
            dbname="{$Field(" + document.myform.Fieldnum.value + "," + dbname + "," + cattype + ")}";
    }else{
            if (cattype=='0'){
                dbname="{$Field(" + document.myform.Fieldnum.value + "," + dbname + "," + cattype + "," + document.myform.ZhengShu.value + ")}";
            }else{
                dbname="{$Field(" + document.myform.Fieldnum.value + "," + dbname + "," + cattype + "," + document.myform.XiaoShu.value + ")}";
            }
    }
    }
    if(dbname=='Time'){
    dbname="{$Field(" + document.myform.Fieldnum.value + "," + dbname + "," + document.myform.Timetype.value + "," + document.myform.Timemb.value + ")}";
    }
    if(dbname=='yn'){
    dbname="{$Field(" + document.myform.Fieldnum.value + "," + dbname + "," + document.myform.yny.value + "," + document.myform.ynn.value + ")}";
    }
    if(dbname=='GetUrl'){
        for (var i=0;i<document.myform.outype.length;i++){
            if (document.myform.outype[i].checked){
                var outype=document.myform.outype[i].value;
        }
        }
        dbname="{$Field(" + document.myform.Fieldnum.value + "," + dbname + "," + document.myform.dbtype.value + "," + outype + ")}";
    }
    if(dbname=='GetClass'|dbname=='GetSpecil'|dbname=='GetChannel'){
        for (var i=0;i<document.myform.outype.length;i++){
            if (document.myform.outype[i].checked){
                var outype=document.myform.outype[i].value;
        }
        }
        dbname="{$Field(" + document.myform.Fieldnum.value + "," + dbname + "," + outype + ")}";
    }
    window.returnValue=dbname;
    window.close();
}
</script>
</head>
<body>
<table id="main" width="100%">
<form method='post' action='' name='myform'>
    <tr class="tdbg"><td><strong>�ֶ����ƣ�</strong><input name='FieldName' type='text' id='FieldName' size='35' value="<% =fieldname %>" readonly></td></tr>
    <tr class="tdbg"><td><strong>������ͣ�</strong><select name="ftype" onChange="changemode()"><option value='Text'>�ı���</option>
<%
If (dbtype > 1 And dbtype < 7) Or dbtype = 131 Then
    response.write "<option value='Num' selected>������</option>"
    isknow = True
Else
    response.write "<option value='Num'>������</option>"
End If
If dbtype = 7 Then
    response.write "<option value='Time' selected>ʱ����</option>"
    isknow = True
Else
    response.write "<option value='Time'>ʱ����</option>"
End If
If dbtype = 11 Then
    response.write "<option value='yn' selected>�Ƿ���</option>"
    isknow = True
Else
    response.write "<option value='yn'>�Ƿ���</option>"
End If

If LCase(fieldname) = "articleid" Or LCase(fieldname) = "softid" Or LCase(fieldname) = "photoid" Or LCase(fieldname) = "productid" Then
        response.write "<option value='GetUrl' selected>����·��(ϵͳ����)</option>"
        isknow = True
    Else
        response.write "<option value='GetUrl'>����·��(ϵͳ����)</option>"
    End If

    If LCase(fieldname) = "classid" Then
        response.write "<option value='GetClass' selected>��Ŀ·��(ϵͳ����)</option>"
        isknow = True
    Else
        response.write "<option value='GetClass'>��Ŀ·��(ϵͳ����)</option>"
    End If

    response.write "<option value='GetSpecil'>ר��·��(ϵͳ����)</option>"

    If LCase(fieldname) = "channelid" Then
        response.write "<option value='GetChannel' selected>Ƶ��·��(ϵͳ����)</option>"
        isknow = True
    Else
        response.write "<option value='GetChannel'>Ƶ��·��(ϵͳ����)</option>"
    End If
%>
</select></td></tr>
<%
If isknow = False Then
    response.write "<tbody id='input1' style='display:'>"
Else
    response.write "<tbody id='input1' style='display:none'>"
End If
%>
    <tr class="tdbg"><td><strong>������ȣ�</strong><input name='CatNum' type='text' id='gotopic' size='20' value=0>&nbsp;&nbsp;&nbsp;<font color='#FF0000'>Ϊ0�򲻽ض�</font></td></tr>
    <tr class="tdbg"><td><strong>���˴���</strong><select name='OutType2'><option value='0' selected>����HTML���</option><option value='1'>������HTML���</option><option value='2'>����HTML���</option></select></td></tr>
    <tr class="tdbg"><td><strong>�ضϴ���</strong><Input type='radio' name='CatType' value='0' checked>��ʾ...&nbsp;&nbsp;<Input type='radio' name='CatType' value='1'>����ʾ...</td></tr>
</tbody>

<%
If ((dbtype > 1 And dbtype < 7) Or dbtype = 131) And Not (LCase(fieldname) = "articleid" Or LCase(fieldname) = "softid" Or LCase(fieldname) = "photoid" Or LCase(fieldname) = "productid" Or LCase(fieldname) = "classid" Or LCase(fieldname) = "channelid") Then
    response.write "<tbody id='input2' style='display:'>"
Else
    response.write "<tbody id='input2' style='display:none'>"
End If
%>
    <tr class="tdbg"><td><strong>�����ʽ��</strong><Input type='radio' name='OutType' value='0' checked onClick="input21.style.display='';input22.style.display='none'">���� <Input type='radio' name='OutType' value='1' onClick="input21.style.display='none';input22.style.display=''">С�� <Input type='radio' name='OutType' value='2' onClick="input21.style.display='none';input22.style.display='none'">�ٷ���</td></tr>
<%
        If ((dbtype > 1 And dbtype < 7) Or dbtype = 131) And Not (LCase(fieldname) = "articleid" Or LCase(fieldname) = "softid" Or LCase(fieldname) = "photoid" Or LCase(fieldname) = "productid" Or LCase(fieldname) = "classid" Or LCase(fieldname) = "channelid") Then
        response.write "<tbody id='input21' style='display:'>"
        Else
        response.write "<tbody id='input21' style='display:none'>"
        End If
%>
        <tr class="tdbg"><td><strong>�����ʽ��</strong><input name='ZhengShu' type='text' id='ZhengShu' size='10' value='0'>&nbsp;&nbsp;<font color='#FF0000'>������ֵ�������,Ϊ0�����ԭ��</font></td></tr></tbody>
    <tbody id='input22' style='display:none'><tr class="tdbg"><td><strong>С��λ����</strong><input name='XiaoShu' type='text' id='XiaoShu' size='4' value=2></td></tr></tbody>
</tbody>


<%
If dbtype = 7 Or dbtype = 135 Then
    response.write "<tbody id='input3' style='display:'>"
Else
    response.write "<tbody id='input3' style='display:none'>"
End If
%>
    <tr class="tdbg"><td><strong>�����ʽ��</strong><select name="Timetype" onChange="changetime()"><option value='0' selected>ģ�����</option><option value='1'>ģ�����(����)</option><option value='2'>ģ�����(��λ����)</option><option value='3'>��������</option></select></td></tr>
    <tr class="tdbg"><td><strong>������壺</strong><input name='Timemb' type='text' id='Timemb' size='35' value="{year}��{month}��{day}��"></td></tr>
</tbody>


<%
If dbtype = 11 Then
    response.write "<tbody id='input4' style='display:'>"
Else
    response.write "<tbody id='input4' style='display:none'>"
End If
%>
    <tr class="tdbg"><td><strong>Ϊ�������</strong><input name='yny' type='text' id='yny' size='20' value="��"></td></tr>
    <tr class="tdbg"><td><strong>Ϊ�������</strong><input name='ynn' type='text' id='ynn' size='20' value="��"></td></tr>
</tbody>


<%
If LCase(fieldname) = "articleid" Or LCase(fieldname) = "softid" Or LCase(fieldname) = "photoid" Or LCase(fieldname) = "productid" Or LCase(fieldname) = "classid" Or LCase(fieldname) = "channelid" Then
    response.write "<tbody id='input5' style='display:'>"
Else
    response.write "<tbody id='input5' style='display:none'>"
End If

If LCase(fieldname) = "articleid" Or LCase(fieldname) = "softid" Or LCase(fieldname) = "photoid" Or LCase(fieldname) = "productid" Then
%>
    <tr class="tdbg"><td><strong>�������</strong><select name="dbtype">
<%
If dbname = 1 Then
    response.write "<option value='Article' selected>������</option>"
Else
    response.write "<option value='Article'>������</option>"
End If
If dbname = 2 Then
    response.write "<option value='Soft' selected>������</option>"
Else
    response.write "<option value='Soft'>������</option>"
End If
If dbname = 3 Then
    response.write "<option value='Photo' selected>ͼƬ��</option>"
Else
    response.write "<option value='Photo'>ͼƬ��</option>"
End If
If dbname = 5 Then
    response.write "<option value='Product' selected>��Ʒ��</option>"
Else
    response.write "<option value='Product'>��Ʒ��</option>"
End If
%>
</select>&nbsp;</td></tr>
<%
End If
%>
<tr class="tdbg"><td><strong>�����ʽ��</strong>
<% if instr(lcase(fieldname),"channelid") = 0 then %>
    <Input type='radio' name='outype' value=3 checked>��� <Input type='radio' name='outype' value='1'>·�� <Input type='radio' name='outype' value='2'>����
<% else %>
    <Input type='radio' name='outype' value=1 checked>Ŀ¼ <Input type='radio' name='outype' value='2'>���� <Input type='radio' name='outype' value='3'>�ϴ�Ŀ¼
<% end if %>
</td></tr>
</tbody>

<tr class="tdbg" align="center"><td><input type='button' value="����" onclick="submitdate();">&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' value="ȡ��" onclick="window.close();"></td></tr>
<tr class="tdbg" height="100%"><td>&nbsp;<input name='Fieldnum' id='Fieldnum' value="<% =num %>" type='hidden'><br>&nbsp;<br>&nbsp;</td></tr>
</form>
</table>
</body>
</html>
