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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="editor_dialog.css">
</head>
<BODY bgColor=#D4D0C8 onload="InitDocument()">
<script language="JavaScript">
var sAction = "INSERT";
var sTitle = "����";

var oControl;
var oSeletion;
var sRangeType;

var sFromUrl = "http://";
var sAlt = "";
var sBorder = "0";
var sBorderColor = "";
var sFilter = "";
var sAlign = "";
var sWidth = "";
var sHeight = "";
var sVSpace = "";
var sHSpace = "";
var UpFileName="None";

var sCheckFlag = "file";

oSelection = dialogArguments.HtmlEdit.document.selection.createRange();
sRangeType = dialogArguments.HtmlEdit.document.selection.type;

if (sRangeType == "Control") {
    if (oSelection.item(0).tagName == "IMG"){
        sAction = "MODI";
        sTitle = "�޸�";
        sCheckFlag = "url";
        oControl = oSelection.item(0);
        sFromUrl = oControl.getAttribute("src", 2);
        sAlt = oControl.alt;
        sBorder = oControl.border;
        sBorderColor = oControl.style.borderColor;
        sFilter = oControl.style.filter;
        sAlign = oControl.align;
        sWidth = oControl.width;
        sHeight = oControl.height;
        sVSpace = oControl.vspace;
        sHSpace = oControl.hspace;

        //dlink.style.display='none';
    }
}


document.write("<title>ͼƬ���ԣ�" + sTitle + "��</title>");

// ����������ֵ��ָ��ֵƥ�䣬��ѡ��ƥ����
function SearchSelectValue(o_Select, s_Value){
    for (var i=0;i<o_Select.length;i++){
        if (o_Select.options[i].value == s_Value){
            o_Select.selectedIndex = i;
            return true;
        }
    }
    return false;
}
// ��ʼֵ
function InitDocument(){

    SearchSelectValue(styletype, sFilter);
    SearchSelectValue(aligntype, sAlign.toLowerCase());
        
    url.value = sFromUrl;
    alttext.value = sAlt;
    border.value = sBorder;
    t_bordercolor.value = sBorderColor;
    s_bordercolor.style.backgroundColor = sBorderColor;
    width.value = sWidth;
    height.value = sHeight;
    vspace.value = sVSpace;
    hspace.value = sHSpace;
    upfilename.value = UpFileName;
    
    if (sAction == "MODI") {
        frmPreview.img.src =sFromUrl;
        frmPreview.img2.src =sFromUrl;
        frmPreview.img.alt=sAlt;
        frmPreview.img.border=sBorder;
        frmPreview.img.style.borderColor =sBorderColor;
        frmPreview.img.style.backgroundColor = sBorderColor;
        frmPreview.img.width=sWidth;
        frmPreview.img.height=sHeight;
        frmPreview.img.vspace=sVSpace;
        frmPreview.img.hspace=sHSpace;
        frmPreview.img.style.filter=sFilter;
       }
}

function OK(){
    sFromUrl = url.value;
    sAlt = alttext.value;
    sBorder = border.value;
    sBorderColor = t_bordercolor.value;
    sFilter = styletype.options[styletype.selectedIndex].value;
    sAlign = aligntype.value;
    sWidth = frmPreview.img.width;
    sHeight = frmPreview.img.height;
    sVSpace = vspace.value;
    sHSpace = hspace.value;
    UpFileName = upfilename.value;
    if (sFromUrl==""|| sFromUrl=="http://"){
         alert("��������ͼƬ�ļ���ַ�������ϴ�ͼƬ�ļ���");
       url.focus();
       return false;
       }
    
    if (sAction == "MODI") {
        oControl.src = sFromUrl;
        oControl.alt = sAlt;
        oControl.border = sBorder;
        oControl.style.borderColor = sBorderColor;
        oControl.style.filter = sFilter;
        oControl.align = sAlign;
        oControl.width = sWidth;
        oControl.height = sHeight;
        oControl.style.width = sWidth;
        oControl.style.height = sHeight;
        oControl.vspace = sVSpace;
        oControl.hspace = sHSpace;
    }else{
        var sHTML = '';
        var slink = '';
        if (addlink.checked == true) {
            slink= ' <a href="'+sFromUrl+'" target=\'_blank\'>';
        }

        if (sFilter!=""){
            sHTML=sHTML+'filter:'+sFilter+';';
        }
        if (sBorderColor!=""){
            sHTML=sHTML+'border-color:'+sBorderColor+';';
        }
        if (sHTML!=""){
            sHTML=' style="'+sHTML+'"';
        }
        sHTML = sHTML+ slink +'<img id=HtmlEdit_TempElement_Img src="'+sFromUrl+'"'+sHTML;
        if (sBorder!=""){
            sHTML=sHTML+' border="'+sBorder+'"';
        }
        if (sAlt!=""){
            sHTML=sHTML+' alt="'+sAlt+'"';
        }
        if (sAlign!=""){
            sHTML=sHTML+' align="'+sAlign+'"';
        }
        if (sWidth!=""){
            sHTML=sHTML+' width="'+sWidth+'"';
        }
        if (sHeight!=""){
            sHTML=sHTML+' height="'+sHeight+'"';
        }
        if (sVSpace!=""){
            sHTML=sHTML+' vspace="'+sVSpace+'"';
        }
        if (sHSpace!=""){
            sHTML=sHTML+' hspace="'+sHSpace+'"';
        }
        <%
        If ShowType = 1 Or ShowType = 2 Or ShowType =3 Or ShowType = 6 Then
            Response.write "sHTML=sHTML+'>';" & vbCrLf
        Else
        %>
             if (zoom.checked == true) {
                sHTML=sHTML+'  onload="resizepic(this)" onmousewheel="return bbimg(this)" >';
                
             }else{
                sHTML=sHTML+' >';
             }
        <%
        End If
        %>
        if (addlink.checked == true) {
            sHTML=sHTML+' </a>';
        }
        dialogArguments.insertHTML(sHTML);
        var oTempElement = dialogArguments.HtmlEdit.document.getElementById("HtmlEdit_TempElement_Img");
        oTempElement.src = sFromUrl;
        oTempElement.removeAttribute("id");
        
    }

    if (UpFileName=="None"){
        window.returnValue = null;
    }else{
        window.returnValue = UpFileName;
    }
    window.close();
}
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
//=================================================
//��������Preview
//��  �ã�������ʾͼƬ
//=================================================
function Preview(){
    if(url.value!="http://"&&url.value!=""){
        frmPreview.img.src=url.value;
        frmPreview.img2.src=url.value;
    }
    else{
        frmPreview.img.src="../images/nopic.gif";
        frmPreview.img2.src="../images/nopic.gif";
    }
    var iheight=height.value;
    var iwidth=width.value;
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

    frmPreview.img.border=border.value;
    frmPreview.img.style.borderColor =t_bordercolor.value;
    frmPreview.img.style.filter=filter.value;
    frmPreview.img.title=alttext.value
 }
//=================================================
//��������SelectColor
//��  �ã���ʾ��ɫ��
//��  ����what  --- Ҫ�����ɫ�Ĳ���
//=================================================
function SelectColor(what){
    var dEL = document.all("t_"+what);
    var sEL = document.all("s_"+what);
    var url = "editor_selcolor.asp?color="+encodeURIComponent(dEL.value);
    var arr = showModalDialog(url,window,"dialogWidth:280px;dialogHeight:250px;help:no;scroll:no;status:no");
    if (arr) {
        dEL.value=arr;
        sEL.style.backgroundColor=arr;
    }
}
function SelectFile(){
  var arr=showModalDialog('<%=InstallDir & AdminDir%>/Admin_SelectFile.asp?DialogType=Pic&ChannelID=<%=ChannelID%>', '', 'dialogWidth:820px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    url.value=ss[0];
    upfilename.value="1$$$" + ss[0].replace("<%=FilesPath%>", "");
    Preview();
  }
}
</script>
  <table border=0 cellpadding=0 cellspacing=0 align=center width='95%'>
    <tr>
      <td><fieldset>
      <legend>Ԥ��ͼƬ</legend>
        <table border=0 cellpadding=0 cellspacing=5>
          <tr>
            <td align='center'><iframe id='frmPreview' width='350' height='220' frameborder='1' src='editor_imgPreview.asp'></iframe></td>
          </tr>
          <tr>
            <td>��ַ��
             <Input name="url" type=text id="url" style="width:243px" onChange="javascript:Preview()" size=30>
            <%if IsUpload=True And AdminName <> "" then %>
             <!--'���ӷ���ģ���õ��� ShowType=3-->
             <Input type="button" name="Submit" value="..." title="�����ϴ��ļ���ѡ��" onClick="SelectFile()">
            <%End if%>
            </td>
          </tr>
        </table>
      </fieldset></td>
      <td width=80 align="center" valign="middle" rowspan="6">
        <table border='0' cellpadding='0' cellspacing='0' width='100%' height='100%' align='center'>
          <tr>
            <td height='230'></td>
          </tr>
          <tr>
            <td width='20'></td>
            <td >
             <Input type="hidden" id="upfilename" value="">
             <Input type=submit value='  ȷ��  ' id=Ok onClick="OK()"><BR><BR>
             <Input type=button value='  ȡ��  ' onClick="window.close();">
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td height=5></td>
    </tr>
    <tr>
      <td><fieldset>
      <legend>��ʾЧ��</legend>
        <table border=0 cellpadding=0 cellspacing=5>
          <tr>
            <td>˵�����֣�</td>
            <td colspan=5>
             <Input name="alttext" type=text id=alttext style="width:243px" onChange="javascript:Preview()" size=38></td>
          </tr>
          <tr>
            <td nowrap>�߿��ϸ��</td>
            <td>
             <Input type=text id=border  name="border" size=10 value="" onKeyPress="event.returnValue=IsDigit();" onChange="javascript:Preview()"></td>
            <td width=40></td>
            <td nowrap>�߿���ɫ��</td>
            <td>
              <table border=0 cellpadding=0 cellspacing=0>
                <tr>
                  <td>
                   <Input type=text id=t_bordercolor name=t_bordercolor size=7 value="" onChange="javascript:Preview()"></td>
                  <td><img border=0 src="images/rect.gif" width=18 style="cursor:hand" id=s_bordercolor onClick="SelectColor('bordercolor');Preview();"> </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td>����Ч����</td>
            <td>
             <Select id=styletype style="width:72px" size=1 name="filter" onChange="javascript:Preview()">
               <option value='' selected>��</option>
               <option value='Alpha(Opacity=50)'>��͸��</option>
               <option value='Alpha(Opacity=0, FinishOpacity=100, Style=1, StartX=0, StartY=0, FinishX=100, FinishY=140)'>����͸��</option>
               <option value='Alpha(Opacity=10, FinishOpacity=100, Style=2, StartX=30, StartY=30, FinishX=200, FinishY=200)'>����͸��</option>
               <option value='blur(add=1,direction=14,strength=15)'>ģ��Ч��</option>
               <option value='blur(add=true,direction=45,strength=30)'>�綯ģ��</option>
               <option value='Wave(Add=0, Freq=60, LightStrength=1, Phase=0, Strength=3)'>���Ҳ���</option>
               <option value='gray'>�ڰ���Ƭ</option>
               <option value='Chroma(Color=#FFFFFF)'>��ɫ͸��</option>
               <option value='DropShadow(Color=#999999, OffX=7, OffY=4, Positive=1)'>Ͷ����Ӱ</option>
               <option value='Shadow(Color=#999999, Direction=45)'>��Ӱ</option>
               <option value='Glow(Color=#ff9900, Strength=5)'>����</option>
               <option value='flipv'>��ֱ��ת</option>
               <option value='fliph'>���ҷ�ת</option>
               <option value='grays'>���Ͳ�ɫ</option>
               <option value='xray'>X����Ƭ</option>
               <option value='invert'>��Ƭ</option>
             </Select>
            </td>
            <td width=40></td>
            <td>���뷽ʽ��</td>
            <td>
             <Select id=aligntype size=1 style="width:72px">
               <option value='' selected>Ĭ��</option>
               <option value='left'>����</option>
               <option value='right'>����</option>
               <option value='top'>����</option>
               <option value='middle'>�в�</option>
               <option value='bottom'>�ײ�</option>
               <option value='absmiddle'>���Ծ���</option>
               <option value='absbottom'>���Եײ�</option>
               <option value='baseline'>����</option>
               <option value='texttop'>�ı�����</option>
             </Select>
            </td>
          </tr>
          <tr>
            <td>ͼƬ��ȣ�</td>
            <td>
             <Input type=text id=width name=width size=10 onKeyPress="event.returnValue=IsDigit();"  onChange="javascript:Preview()" maxlength=4></td>
            <td width=40></td>
            <td>ͼƬ�߶ȣ�</td>
            <td>
             <Input type=text id=height name=height size=10 onKeyPress="event.returnValue=IsDigit();" maxlength=4 onChange="javascript:Preview()"></td>
          </tr>
          <tr>
            <td>���¼�ࣺ</td>
            <td>
             <Input type=text id=vspace size=10 value="" onKeyPress="event.returnValue=IsDigit();" maxlength=2 ></td>
            <td width=40></td>
            <td>���Ҽ�ࣺ</td>
            <td>
             <Input type=text id=hspace size=10 value="" onKeyPress="event.returnValue=IsDigit();" maxlength=2></td>
          </tr>
            <tr id=dlink style="display:''">
                <td colspan='2' ><INPUT TYPE='checkbox' NAME='zoom' id="zoom" value='Yes' checked>����ͼƬ����JS����</td>
                <td width=40></td>
                <td colspan='2'><INPUT TYPE='checkbox' NAME='addlink' id="addlink" value='Yes' checked>��ӵ�ԭʼͼƬ������</td>
            </tr>
        </table>
      </fieldset></td>
    </tr>
    <%if IsUpload=True then %>
    <tr>
      <td><fieldset align=left>
      <legend align=left>�ϴ�����ͼƬ</legend>
    <%
        Response.write "<iframe class=""TBGen"" style=""top:2px"" id=""UploadFiles"" src=""upload.asp?DialogType=pic"
        Response.write "&ChannelID=" & ChannelID
        If PE_CLng(Request(Trim("Anonymous"))) = 1 Then
            Response.write "&Anonymous=1"
        End If		
        If ModuleType=3 Then
            Response.write "&PhotoUpfileType=1"
        End If
        Response.write """ frameborder=0 scrolling=no width=""350"" height=""25""></iframe>"
        Response.write "</fieldset></td>"
        Response.write "</tr>"
    End if 
    %>
    <tr>
      <td height=5></td>
    </tr>
  </table>
</body>
</html>
