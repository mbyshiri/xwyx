<!--#include file="editor_ChkPurview.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

%>
<HTML>
<HEAD>
<TITLE>��������ͼƬ</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="editor_dialog.css">
<base target="_self">
<script language="JavaScript">
function IsDigit(){
    return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
function ShowThumbSetting(x){
    if(eval("document.form1.CreateThumb"+x+".checked==true")){
        eval("Thumb_"+x+".style.display='';");
    }
    else{
        eval("Thumb_"+x+".style.display='none';");
    }
}
function Preview(num){
    var sfilename=document.all("FileName"+num).value;
    if(sfilename!=""){
        frmPreview.img.src=sfilename;
        frmPreview.img2.src=sfilename;
    }
    else{
        frmPreview.img.src="../images/nopic.gif";
        frmPreview.img2.src="../images/nopic.gif";
    }
    var iheight=document.all("height"+num).value;
    var iwidth=document.all("width"+num).value;
    if(iheight>0){
        if(iwidth>0){
            frmPreview.img.height=iheight;
            frmPreview.img.width=iwidth
        }
        else{
            frmPreview.img.height=iheight;
            frmPreview.img.width=iheight/frmPreview.img2.height*frmPreview.img2.width;
        }
    }
    else{
        if(iwidth>0){
            frmPreview.img.width=iwidth
            frmPreview.img.height=iwidth/frmPreview.img2.width*frmPreview.img2.height;
        }
        else{
            frmPreview.img.height=frmPreview.img2.height;
            frmPreview.img.width=frmPreview.img2.width;
        }
    }
    frmPreview.img.border=document.all("border"+num).value;
    frmPreview.img.style.borderColor =document.all("bordercolor"+num).value;
    frmPreview.img.style.filter=document.all("filter"+num).value;
    frmPreview.img.title=document.all("alttext"+num).value
}
function change_item(num){
    var p=Preview(num);
    for (td_i=0;td_i<10;td_i++){
        if (td_i==num){
            eval("td_"+td_i+".style.display='';");
            eval("tdcolor"+td_i+".style.backgroundColor='#ffffff';");
        }
        else{
            eval("td_"+td_i+".style.display=\"none\";");
            eval("tdcolor"+td_i+".style.backgroundColor='#D4D0C8'");
        }
   }
}
function mysub()
{
  esave.style.visibility="visible";
}
function SelectColor(what){
    var dEL = document.all(what);
    var sEL = document.all("s_"+what);
    var url = "editor_selcolor.asp?color="+encodeURIComponent(dEL.value);
    var arr = showModalDialog(url,window,"dialogWidth:280px;dialogHeight:250px;help:no;scroll:no;status:no");
    if (arr) {
        dEL.value=arr;
        sEL.style.backgroundColor=arr;
    }
}
//-->
</script>
</head>

<BODY bgColor=#D4D0C8 topmargin='15' leftmargin='15' >
<br>
<form name="form1" method="post" action="Upfile.asp" enctype="multipart/form-data">
  <table border='0' cellpadding='0' cellspacing='0' width='100%' align='center'>
   <tr>
     <td valign="top">
<%
    Dim i
    For i = 0 To 9
        Response.Write "<table width=100% border='0' align='center' cellpadding='0' cellspacing='2'>" & vbCrLf
        Response.Write "<tr id='tdcolor" & i & "'"
        If i = 0 Then Response.Write " bgcolor='#ffffff' "
        Response.Write " onCLICK='change_item(" & i & ")'><td width='50'>ͼƬ" & i + 1 & "��</td>" & vbCrLf
        Response.Write "<td><input name='FileName" & i & "' type='FILE'  size='30' onChange='change_item(" & i & ")' ></td>" & vbCrLf
        Response.Write "<td> �� ��>>> </td>" & vbCrLf
        Response.Write "</tr></table>" & vbCrLf
    Next
%>
<br>
<br>
���ӵ�ַ:
<input name="LinkUrl" type="text" id="LinkUrl" value="http://" size="40" maxlength="200">
<br>
<br>
˵������������ϴ���ЩͼƬʱ��һ���Լ������ӣ����Ҫ�޸����ӣ��������ϴ���ɺ��ڱ༭���޸��������ԡ�</td>
     <td valign='top'><fieldset><legend>ͼƬ��������</legend>
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height='220' align="center"><iframe id='frmPreview' width='350' height='220' frameborder='1' src='editor_imgPreview.asp'></iframe></td>
          </tr>
          <tr>
            <td>
<%
    For i = 0 To 9
        Response.Write "<table id='td_" & i & "' width=100%  height=100% border='0' align='center' cellpadding='0' cellspacing='2'"
        If i > 0 Then Response.Write " style='display:none'"
        Response.Write ">" & vbCrLf
        Response.Write "<tr><td colspan='2'>˵�����֣�<input name='alttext" & i & "' size=38 maxlength='100' onChange='Preview(" & i & ")'></td></tr>" & vbCrLf
        Response.Write "<tr><td>�߿��ϸ��<input name='border" & i & "' ONKEYPRESS='event.returnValue=IsDigit();'  value='0' size=5 maxlength='2' onChange='Preview(" & i & ")'>����</td>" & vbCrLf
        Response.Write "<td>�߿���ɫ��<input name ='bordercolor" & i & "' type=text size=7 value='' onChange='Preview(" & i & ")'>"
        Response.Write "&nbsp;<img border=0 src='images/rect.gif' width=18 style='cursor:hand' id='s_bordercolor" & i & "' onclick=""SelectColor('bordercolor" & i & "');Preview(" & i & ");""></td></tr>"
        Response.Write "<tr><td>����Ч����<select name='filter" & i & "' onChange='Preview(" & i & ")'>" & vbCrLf
        Response.Write "<option value=''  selected>��Ӧ��</option>" & vbCrLf
        Response.Write "<option value='Alpha(Opacity=50)'>��͸��Ч��</option>" & vbCrLf
        Response.Write "<option value='Alpha(Opacity=0, FinishOpacity=100, Style=1, StartX=0,StartY=0, FinishX=100, FinishY=140)'>����͸��Ч��</option>" & vbCrLf
        Response.Write "<option value='Alpha(Opacity=10, FinishOpacity=100, Style=2, StartX=30,StartY=30, FinishX=200, FinishY=200)'>����͸��Ч��</option>" & vbCrLf
        Response.Write "<option value='blur(add=1,direction=14,strength=15)'>ģ��Ч��</option>" & vbCrLf
        Response.Write "<option value='blur(add=true,direction=45,strength=30)'>�綯ģ��Ч��</option>" & vbCrLf
        Response.Write "<option value='Wave(Add=0, Freq=60, LightStrength=1, Phase=0,Strength=3)'>���Ҳ���Ч��</option>" & vbCrLf
        Response.Write "<option value='gray'>�ڰ���ƬЧ��</option>" & vbCrLf
        Response.Write "<option value='Chroma(Color=#FFFFFF)'>��ɫΪ͸��</option>" & vbCrLf
        Response.Write "<option value='DropShadow(Color=#999999, OffX=7, OffY=4, Positive=1)'>Ͷ����ӰЧ��</option>" & vbCrLf
        Response.Write "<option value='Shadow(Color=#999999, Direction=45)'>��ӰЧ��</option>" & vbCrLf
        Response.Write "<option value='Glow(Color=#ff9900, Strength=5)'>����Ч��</option>" & vbCrLf
        Response.Write "<option value='flipv'>��ֱ��ת��ʾ</option>" & vbCrLf
        Response.Write "<option value='fliph'>���ҷ�ת��ʾ</option>" & vbCrLf
        Response.Write "<option value='grays'>���Ͳ�ɫ��</option>" & vbCrLf
        Response.Write "<option value='xray'>X����ƬЧ��</option>" & vbCrLf
        Response.Write "<option value='invert'>��ƬЧ��</option>" & vbCrLf
        Response.Write "</select>" & vbCrLf
        Response.Write "</td>" & vbCrLf
        Response.Write "<td>ͼƬλ�ã�<select name='aligntype" & i & "'>" & vbCrLf
        Response.Write "<option value='' selected>Ĭ��λ��" & vbCrLf
        Response.Write "<option value='left'>����" & vbCrLf
        Response.Write "<option value='right' >����" & vbCrLf
        Response.Write "<option value='top'>����" & vbCrLf
        Response.Write "<option value='middle'>�в�" & vbCrLf
        Response.Write "<option value='bottom'>�ײ�" & vbCrLf
        Response.Write "<option value='absmiddle'>���Ծ���" & vbCrLf
        Response.Write "<option value='absbottom'>���Եײ�" & vbCrLf
        Response.Write "<option value='baseline'>����" & vbCrLf
        Response.Write "<option value='texttop'>�ı�����" & vbCrLf
        Response.Write "</select></td>" & vbCrLf
        Response.Write "</tr>" & vbCrLf
        Response.Write "<tr>" & vbCrLf
        Response.Write "<td>ͼƬ��ȣ�<input name='width" & i & "' value='' ONKEYPRESS='event.returnValue=IsDigit();' size=4 maxlength='4' onChange='Preview(" & i & ")'>����</td>" & vbCrLf
        Response.Write "<td>ͼƬ�߶ȣ�<input name='height" & i & "' value='' onKeyPress='event.returnValue=IsDigit();' size=4 maxlength='4' onChange='Preview(" & i & ")'>����</td>" & vbCrLf
        Response.Write "</tr><tr>" & vbCrLf
        Response.Write "<td>���¼�ࣺ<input name='vspace" & i & "' ONKEYPRESS='event.returnValue=IsDigit();' value='0' size=4 maxlength='2'>����</td>" & vbCrLf
        Response.Write "<td>���Ҽ�ࣺ<input name='hspace" & i & "' onKeyPress='event.returnValue=IsDigit();'  value='0' size=4 maxlength='2'>����</td>" & vbCrLf
        Response.Write "</tr>" & vbCrLf

        Response.Write "<tr>" & vbCrLf
        Response.Write "<td>�Ƿ����ͼƬ����JS���룺<INPUT TYPE='checkbox' NAME='zoom" & i & "' value='Yes' checked></td>" & vbCrLf
        Response.Write "<td></td>" & vbCrLf
        Response.Write "</tr>" & vbCrLf

        If PhotoObject > 0 Then
            Response.Write "<tr><td>�Ƿ��ˮӡ��<INPUT TYPE='checkbox' NAME='AddWatermark" & i & "' value='Yes' checked></td>" & vbCrLf
            Response.Write "<td>�Ƿ���������ͼ��<INPUT TYPE='checkbox' NAME='CreateThumb" & i & "' value='Yes' onCLICK='javascript:ShowThumbSetting(" & i & ");'"
            If i = 0 then 
                Response.Write "checked"
            End If
            Response.Write "></td></tr>"
            Response.Write "<tr style='display:"
            If i > 0 then 
                Response.Write "none"
            End If
            Response.Write"' id='Thumb_" & i & "'><td>����ͼ��ȣ�<input name='ThumbWidth" & i & "' ONKEYPRESS='event.returnValue=IsDigit();' value='" & Thumb_DefaultWidth & "' size=4 maxlength='3'>����</td>" & vbCrLf
            Response.Write "<td>����ͼ�߶ȣ�<input name='ThumbHeight" & i & "' onKeyPress='event.returnValue=IsDigit();'  value='" & Thumb_DefaultHeight & "' size=4 maxlength='3'>����</td></tr>" & vbCrLf
        End If
        Response.Write "</table>" & vbCrLf
    Next
%>
</td>
          </tr>
        </table></fieldset>
    </td></tr>
    <tr><td align='center' colspan='2' >
    <input name='FileType' type='hidden' value='BatchPic'>
    <input name='Anonymous' type='hidden' id='Anonymous' value='<%=PE_CLng(Trim(Request("Anonymous")))%>'>	
    <input name='ChannelID' type='hidden' id='ChannelID' value='<%=ChannelID%>'>
    <input name='cmdOK' type='submit' id='cmdOK' value='  ȷ��  '  onclick='javascript:mysub()'>&nbsp;&nbsp;
    <input name='cmdCancel' type=button id='cmdCancel' onclick='window.close();' value='  ȡ��  '>
     </td>
  </tr>
</table>
 <div id="esave" style="position:absolute; top:10px; left:200px; z-index:1; visibility:hidden">
    <TABLE WIDTH=400 BORDER=0 CELLSPACING=0 CELLPADDING=0>
      <TR><td width=20%></td>
    <TD width="60%">
    <TABLE WIDTH=100% height=100 BORDER=0 CELLSPACING=1 CELLPADDING=0>
    <TR>
      <td bgcolor="#0033FF" align=center><b><marquee align="middle" behavior="alternate" scrollamount="5"><font color=#FFFFFF>...�ļ��ϴ���...��ȴ�...</font></marquee></b></td>
    </tr>
    </table>
    </td><td width='20%'></td>
    </tr>
    </table>
  </div>
</form>
</body>
</html>