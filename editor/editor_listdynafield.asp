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
Dim ID, sql, rs

ID = Trim(request("id"))

If ID="" Then
    response.write "ID错误"
Else
    ID=Clng(ID)
    If ID = 0 Then
        response.write "ID错误"
    Else
        Call Main
    End If
End If

Call CloseConn

Sub Main()
    Set rs=Conn.Execute("select * from PE_Label where LabelID=" & ID)
    If rs.bof and rs.EOF Then
        response.write "标签不存在"
    Else
%>
<html>
<head>
<title>显示动态标签参数输入框</title>
<script src="../JS/prototype.js"></script>
<script language="javascript">
function objectTag(itotal) {
        var PowerEasy="";
        //var errstat=0;
        for(i=0;i<itotal;i++){
            if($F("field_" + i) == ''){
                alert("请填写完整");
                Field.focus("field_" + i);
                errstat=1;
                return;
            }
            if(i<itotal-1){
                PowerEasy = PowerEasy + $F("field_" + i) + ","; 
            }else{
                PowerEasy = PowerEasy + $F("field_" + i); 
            }
        }
        //if(errstat==0){
	    var reval = '{$<% = rs("LabelName") %>('+PowerEasy+')}';  
	    window.returnValue = reval;
	    window.close();
        //}
}
</script>
<link href='Lable/Admin_Style.css' rel='stylesheet' type='text/css'>
</head>
<body>
<form name="form1">
<table width='100%' border='0' align='center' align='center' cellpadding='2' cellspacing='1' class='border'>
  <tr class='title'>
    <td colspan="2" align="center"><strong>请输入动态函数标签参数</strong></td>
  </tr>
<%
   Dim arrFieldList, i, total
   If Trim(rs("fieldlist") & "") = "" Then
       response.write "<tr class='tdbg'><td align='center'><font color=""red"">该标签未建立参数列表，请手动输入!</font></td></tr>"
   Else
       arrFieldList = Split(rs("fieldlist"), "|||")
       total = UBound(arrFieldList)
       For i = 0 To UBound(arrFieldList) - 1
          response.write "<tr class='tdbg'><td align='right'>" & arrFieldList(i) & "：</td><td><input type=""text"" id='field_" & i & "' name='field_" & i & "'></td></tr>"
       Next
    End If
    response.write "<tr class='tdbg'><td colspan=2 align='center'><input TYPE='button' value=' 确 定 ' onCLICK='objectTag(" & total & ")'></td>"   
%>
  </tr>
</table>
</form>
</body>
</html>
<%

    End If
    Set rs = Nothing
End Sub
%>
