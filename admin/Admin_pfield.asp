<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
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
<title>字段属性设置</title>
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
        document.myform.Timemb.value="{year}年{month}月{day}日";
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
    <tr class="tdbg"><td><strong>字段名称：</strong><input name='FieldName' type='text' id='FieldName' size='35' value="<% =fieldname %>" readonly></td></tr>
    <tr class="tdbg"><td><strong>输出类型：</strong><select name="ftype" onChange="changemode()"><option value='Text'>文本型</option>
<%
If (dbtype > 1 And dbtype < 7) Or dbtype = 131 Then
    response.write "<option value='Num' selected>数字型</option>"
    isknow = True
Else
    response.write "<option value='Num'>数字型</option>"
End If
If dbtype = 7 Then
    response.write "<option value='Time' selected>时间型</option>"
    isknow = True
Else
    response.write "<option value='Time'>时间型</option>"
End If
If dbtype = 11 Then
    response.write "<option value='yn' selected>是否型</option>"
    isknow = True
Else
    response.write "<option value='yn'>是否型</option>"
End If

If LCase(fieldname) = "articleid" Or LCase(fieldname) = "softid" Or LCase(fieldname) = "photoid" Or LCase(fieldname) = "productid" Then
        response.write "<option value='GetUrl' selected>对象路径(系统内置)</option>"
        isknow = True
    Else
        response.write "<option value='GetUrl'>对象路径(系统内置)</option>"
    End If

    If LCase(fieldname) = "classid" Then
        response.write "<option value='GetClass' selected>栏目路径(系统内置)</option>"
        isknow = True
    Else
        response.write "<option value='GetClass'>栏目路径(系统内置)</option>"
    End If

    response.write "<option value='GetSpecil'>专题路径(系统内置)</option>"

    If LCase(fieldname) = "channelid" Then
        response.write "<option value='GetChannel' selected>频道路径(系统内置)</option>"
        isknow = True
    Else
        response.write "<option value='GetChannel'>频道路径(系统内置)</option>"
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
    <tr class="tdbg"><td><strong>输出长度：</strong><input name='CatNum' type='text' id='gotopic' size='20' value=0>&nbsp;&nbsp;&nbsp;<font color='#FF0000'>为0则不截断</font></td></tr>
    <tr class="tdbg"><td><strong>过滤处理：</strong><select name='OutType2'><option value='0' selected>解析HTML标记</option><option value='1'>不解析HTML标记</option><option value='2'>过滤HTML标记</option></select></td></tr>
    <tr class="tdbg"><td><strong>截断处理：</strong><Input type='radio' name='CatType' value='0' checked>显示...&nbsp;&nbsp;<Input type='radio' name='CatType' value='1'>不显示...</td></tr>
</tbody>

<%
If ((dbtype > 1 And dbtype < 7) Or dbtype = 131) And Not (LCase(fieldname) = "articleid" Or LCase(fieldname) = "softid" Or LCase(fieldname) = "photoid" Or LCase(fieldname) = "productid" Or LCase(fieldname) = "classid" Or LCase(fieldname) = "channelid") Then
    response.write "<tbody id='input2' style='display:'>"
Else
    response.write "<tbody id='input2' style='display:none'>"
End If
%>
    <tr class="tdbg"><td><strong>输出方式：</strong><Input type='radio' name='OutType' value='0' checked onClick="input21.style.display='';input22.style.display='none'">整数 <Input type='radio' name='OutType' value='1' onClick="input21.style.display='none';input22.style.display=''">小数 <Input type='radio' name='OutType' value='2' onClick="input21.style.display='none';input22.style.display='none'">百分数</td></tr>
<%
        If ((dbtype > 1 And dbtype < 7) Or dbtype = 131) And Not (LCase(fieldname) = "articleid" Or LCase(fieldname) = "softid" Or LCase(fieldname) = "photoid" Or LCase(fieldname) = "productid" Or LCase(fieldname) = "classid" Or LCase(fieldname) = "channelid") Then
        response.write "<tbody id='input21' style='display:'>"
        Else
        response.write "<tbody id='input21' style='display:none'>"
        End If
%>
        <tr class="tdbg"><td><strong>输出方式：</strong><input name='ZhengShu' type='text' id='ZhengShu' size='10' value='0'>&nbsp;&nbsp;<font color='#FF0000'>根据数值输出符号,为0则输出原数</font></td></tr></tbody>
    <tbody id='input22' style='display:none'><tr class="tdbg"><td><strong>小数位数：</strong><input name='XiaoShu' type='text' id='XiaoShu' size='4' value=2></td></tr></tbody>
</tbody>


<%
If dbtype = 7 Or dbtype = 135 Then
    response.write "<tbody id='input3' style='display:'>"
Else
    response.write "<tbody id='input3' style='display:none'>"
End If
%>
    <tr class="tdbg"><td><strong>输出格式：</strong><select name="Timetype" onChange="changetime()"><option value='0' selected>模板输出</option><option value='1'>模板输出(补零)</option><option value='2'>模板输出(限位补零)</option><option value='3'>函数处理</option></select></td></tr>
    <tr class="tdbg"><td><strong>输出摸板：</strong><input name='Timemb' type='text' id='Timemb' size='35' value="{year}年{month}月{day}日"></td></tr>
</tbody>


<%
If dbtype = 11 Then
    response.write "<tbody id='input4' style='display:'>"
Else
    response.write "<tbody id='input4' style='display:none'>"
End If
%>
    <tr class="tdbg"><td><strong>为真输出：</strong><input name='yny' type='text' id='yny' size='20' value="是"></td></tr>
    <tr class="tdbg"><td><strong>为假输出：</strong><input name='ynn' type='text' id='ynn' size='20' value="否"></td></tr>
</tbody>


<%
If LCase(fieldname) = "articleid" Or LCase(fieldname) = "softid" Or LCase(fieldname) = "photoid" Or LCase(fieldname) = "productid" Or LCase(fieldname) = "classid" Or LCase(fieldname) = "channelid" Then
    response.write "<tbody id='input5' style='display:'>"
Else
    response.write "<tbody id='input5' style='display:none'>"
End If

If LCase(fieldname) = "articleid" Or LCase(fieldname) = "softid" Or LCase(fieldname) = "photoid" Or LCase(fieldname) = "productid" Then
%>
    <tr class="tdbg"><td><strong>数据类别：</strong><select name="dbtype">
<%
If dbname = 1 Then
    response.write "<option value='Article' selected>文章型</option>"
Else
    response.write "<option value='Article'>文章型</option>"
End If
If dbname = 2 Then
    response.write "<option value='Soft' selected>下载型</option>"
Else
    response.write "<option value='Soft'>下载型</option>"
End If
If dbname = 3 Then
    response.write "<option value='Photo' selected>图片型</option>"
Else
    response.write "<option value='Photo'>图片型</option>"
End If
If dbname = 5 Then
    response.write "<option value='Product' selected>商品型</option>"
Else
    response.write "<option value='Product'>商品型</option>"
End If
%>
</select>&nbsp;</td></tr>
<%
End If
%>
<tr class="tdbg"><td><strong>输出方式：</strong>
<% if instr(lcase(fieldname),"channelid") = 0 then %>
    <Input type='radio' name='outype' value=3 checked>混合 <Input type='radio' name='outype' value='1'>路径 <Input type='radio' name='outype' value='2'>名称
<% else %>
    <Input type='radio' name='outype' value=1 checked>目录 <Input type='radio' name='outype' value='2'>名称 <Input type='radio' name='outype' value='3'>上传目录
<% end if %>
</td></tr>
</tbody>

<tr class="tdbg" align="center"><td><input type='button' value="插入" onclick="submitdate();">&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' value="取消" onclick="window.close();"></td></tr>
<tr class="tdbg" height="100%"><td>&nbsp;<input name='Fieldnum' id='Fieldnum' value="<% =num %>" type='hidden'><br>&nbsp;<br>&nbsp;</td></tr>
</form>
</table>
</body>
</html>
