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
%>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <meta name="Keywords" content="动易网络科技有限公司，动易网站管理系统，动易，动力文章系统，文章系统，动力，整站系统，系统，网站制作，网站，设计，网页设计制作，软件，帮助，系统帮助，动易帮助，支持，安装帮助">
    <title>网站管理系统--标签导航</title>
</head>
<body leftmargin="0"  rightmargin="0"topmargin="0">

<!-- ******** 菜单效果开始 ******** -->
<table width="100%"  border="0" cellspacing="0" cellpadding="4" align="center">
  <tr>
    <td align="center" bgcolor="#0066FF"><b><font color="#ffffff">网站管理系统--标签导航</font></b></td>
  </tr>
</table>
<table width="90%"  border="0" cellspacing="0" cellpadding="2" align="center">
  <tr>
    <td height="50" valign="top" background="images/left_tdbg_01.gif">
      <style rel=stylesheet type=text/css>
        td {
        FONT-SIZE: 9pt; COLOR: #000000; FONT-FAMILY: 宋体,Dotum,DotumChe,Arial;line-height: 150%;
        }
        INPUT {
        BACKGROUND-COLOR: #ffffff;
        BORDER-BOTTOM: #666666 1px solid;
        BORDER-LEFT: #666666 1px solid;
        BORDER-RIGHT: #666666 1px solid;
        BORDER-TOP: #666666 1px solid;
        COLOR: #666666;
        HEIGHT: 18px;
        border-color: #666666 #666666 #666666 #666666; font-size: 9pt
        }
        .favMenu {
            BACKGROUND: #ffffff; COLOR: windowtext; CURSOR: hand;line-height: 150%;
        }
        .favMenu DIV {
            WIDTH: 100%;height: 5px;
        }
        .favMenu A {
            COLOR: windowtext; TEXT-DECORATION: none
        }
        .favMenu A:hover {
            COLOR: windowtext; TEXT-DECORATION: underline
        }
        .topFolder {
            
        }
        .topItem {

        }
        .subFolder {
            PADDING: 0px;BACKGROUND: #ffffff;
        }
        .subItem {
            PADDING: 0px;BACKGROUND: #ffffff;
        }
        .sub {
            BACKGROUND: #ffffff;DISPLAY: none; PADDING: 0px;
        }
        .sub .sub {
            BORDER: 0px;BACKGROUND: #ffffff;
        }
        .icon {
            HEIGHT: 18px; MARGIN-RIGHT: 5px; VERTICAL-ALIGN: absmiddle; WIDTH: 18px
        }
        .outer {
            BACKGROUND: #ffffff;PADDING: 0px;
        }
        .inner {
            BACKGROUND: #ffffff;PADDING: 0px;
        }
        .scrollButton {
            BACKGROUND: #ffffff; BORDER: #f6f6f6 1px solid; PADDING: 0px;
        }
        .flat {
            BACKGROUND: #ffffff; BORDER: #f6f6f6 1px solid; PADDING: 0px;
        }
    </style>
    <SCRIPT type=text/javascript>

    var selectedItem = null;
    var targetWin;

    document.onclick = handleClick;
    document.onmouseover = handleOver;
    document.onmouseout = handleOut;
    document.onmousedown = handleDown;
    document.onmouseup = handleUp;
    document.write(writeSubPadding(10));

    function handleClick() {
        el = getReal(window.event.srcElement, "tagName", "DIV");
        
        if ((el.className == "topFolder") || (el.className == "subFolder")) {
            el.sub = eval(el.id + "Sub");
            if (el.sub.style.display == null) el.sub.style.display = "none";
            if (el.sub.style.display != "block") {
                if (el.parentElement.openedSub != null) {
                    var opener = eval(el.parentElement.openedSub + ".opener");
                    ChangeFolderImg(opener,1)
                    hide(el.parentElement.openedSub);
                    if (opener.className == "topFolder")
                        outTopItem(opener);
                }
                el.sub.style.display = "block";
                el.sub.parentElement.openedSub = el.sub.id;
                el.sub.opener = el;
                ChangeFolderImg(el,2)
            }
            else {
                hide(el.sub.id);
                ChangeFolderImg(el,1)
            }
        }
        if ((el.className == "subItem") || (el.className == "subFolder")) {
            if (selectedItem != null)
                restoreSubItem(selectedItem);
            highlightSubItem(el);
        }
        if ((el.className == "topItem") || (el.className == "topFolder")) {
            if (selectedItem != null)
                restoreSubItem(selectedItem);
        }
        if ((el.className == "topItem") || (el.className == "subItem")) {
            if ((el.href != null) && (el.href != "")) {
                if ((el.target == null) || (el.target == "")) {
                    if (window.opener == null) {
                        if (document.all.tags("BASE").item(0) != null)
                            window.open(el.href, document.all.tags("BASE").item(0).target);
                        else
                            window.location = el.href;
                    }
                    else {
                        window.opener.location =  el.href;
                    }
                }
                else {
                    window.open(el.href, el.target);
                }
            }
        }
        var tmp  = getReal(el, "className", "favMenu");
        if (tmp.className == "favMenu") fixScroll(tmp);
    }
    function handleOver() {
        var fromEl = getReal(window.event.fromElement, "tagName", "DIV");
        var toEl = getReal(window.event.toElement, "tagName", "DIV");
        if (fromEl == toEl) return;
        el = toEl;
        if ((el.className == "topFolder") || (el.className == "topItem")) overTopItem(el);
        if ((el.className == "subFolder") || (el.className == "subItem")) overSubItem(el);
        if ((el.className == "topItem") || (el.className == "subItem")) {
            if (el.href != null) {
                if (el.oldtitle == null) el.oldtitle = el.title;
                if (el.oldtitle != "")
                    el.title = el.oldtitle + "\n" + el.href;
                else
                    el.title = el.oldtitle + el.href;
            }
        }
        if (el.className == "scrollButton") overscrollButton(el);
    }
    function handleOut() {
        var fromEl = getReal(window.event.fromElement, "tagName", "DIV");
        var toEl = getReal(window.event.toElement, "tagName", "DIV");
        if (fromEl == toEl) return;
        el = fromEl;
        if ((el.className == "topFolder") || (el.className == "topItem")) outTopItem(el);
        if ((el.className == "subFolder") || (el.className == "subItem")) outSubItem(el);
        if (el.className == "scrollButton") outscrollButton(el);
    }
    function handleDown() {
        el = getReal(window.event.srcElement, "tagName", "DIV");
        if (el.className == "scrollButton") {
            downscrollButton(el);
            var mark = Math.max(el.id.indexOf("Up"), el.id.indexOf("Down"));
            var type = el.id.substr(mark);
            var menuID = el.id.substring(0,mark);
            eval("scroll" + type + "(" + menuID + ")");
        }
    }
    function handleUp() {
        el = getReal(window.event.srcElement, "tagName", "DIV");
        if (el.className == "scrollButton") {
            upscrollButton(el);
            window.clearTimeout(scrolltimer);
        }
    }
    ////////////////////// EVERYTHING IS HANDLED ////////////////////////////
    function hide(elID) {
        var el = eval(elID);
        el.style.display = "none";
        el.parentElement.openedSub = null;
    }
    function writeSubPadding(depth) {
        var str, str2, val;
        var str = "<style type='text/css'>\n";
        for (var i=0; i < depth; i++) {
            str2 = "";
            val  = 0;
            for (var j=0; j < i; j++) {
                str2 += ".sub "
                val += 18;    //子栏目左边距数值
            }
            str += str2 + ".subFolder {padding-left: " + val + "px;}\n";
            str += str2 + ".subItem   {padding-left: " + val + "px;}\n";
        }
        str += "</style>\n";
        return str;
    }
    function overTopItem(el) {
        with (el.style) {
            background   = "#f8f8f8";
            paddingBottom = "0px";
        }
    }
    function outTopItem(el) {
        if ((el.sub != null) && (el.parentElement.openedSub == el.sub.id)) {
            with(el.style) {
                background = "#ffffff";
            }
        }
        else {
            with (el.style) {
                background = "#ffffff";
                padding = "0px";
            }
        }
    }
    function overSubItem(el) {
            el.style.background = "#F6F6F6";
        el.style.textDecoration = "underline";
    }
    function outSubItem(el) {
        el.style.background = "#ffffff";
        el.style.textDecoration = "none";
    }
    function highlightSubItem(el) {
        el.style.background = "#ffffff";
        el.style.color      = "#ff0000";
        selectedItem = el;
    }
    function restoreSubItem(el) {
        el.style.background = "#ffffff";
        el.style.color      = "menutext";
        selectedItem = null;
    }
    function overscrollButton(el) {
        overTopItem(el);
        el.style.padding = "0px";
    }
    function outscrollButton(el) {
        outTopItem(el);
        el.style.padding = "0px";
    }
    function downscrollButton(el) {
        with (el.style) {
            borderRight   = "0px solid buttonhighlight";
            borderLeft  = "0px solid buttonshadow";
            borderBottom    = "0px solid buttonhighlight";
            borderTop = "0px solid buttonshadow";
        }
    }
    function upscrollButton(el) {
        overTopItem(el);
        el.style.padding = "0px";
    }
    function getReal(el, type, value) {
        var temp = el;
        while ((temp != null) && (temp.tagName != "BODY")) {
            if (eval("temp." + type) == value) {
                el = temp;
                return el;
            }
            temp = temp.parentElement;
        }
        return el;
    }
    ////////////////////////////////////////////////////////////////////////////////////////
    // Fix the scrollbars
    var globalScrollContainer;
    var overflowTimeout = 1;

    function fixScroll(el) {
        globalScrollContainer = el;
        window.setTimeout('changeOverflow(globalScrollContainer)', overflowTimeout);
    }
    function changeOverflow(el) {
        if (el.offsetHeight > el.parentElement.clientHeight)
            window.setTimeout('globalScrollContainer.parentElement.style.overflow = "auto";', overflowTimeout);
        else
            window.setTimeout('globalScrollContainer.parentElement.style.overflow = "hidden";', overflowTimeout);
    }
    function ChangeFolderImg(el,type) {
        var FolderImg = eval(el.id + "Img");
        if (type == 1) {
            FolderImg.src = "images/foldericon1.gif"
        }
        else {
            FolderImg.src = "images/foldericon2.gif"
        }
    }
    function InsertAdjs(adjs){
        window.returnValue ="<IMG alt='#[!"+"script language=\"javascript\" src=\""+adjs+"\"!][!/"+"script!]#'  src=\"editor/images/jscript.gif\" border=0 $>"
        window.close();
    }
    </SCRIPT>

    <!--  广告位标签开始  -->
    <DIV class=topFolder id=AdJs><IMG id=AdJsImg class=icon src="images/foldericon1.gif">广告版位标签</DIV>
    <DIV class=sub id=AdJsSub>
    <%
        Dim rs
		Set rs = Conn.Execute("select ZoneID,ZoneName,UpdateTime from PE_AdZone")
        If rs.bof And rs.EOF Then
           Response.write ("<li>您还没有定义广告版位 </li>")
        Else
            Do While Not rs.EOF
                Response.write "<DIV class=subItem onclick=""InsertAdjs('" & GetZoneJSName(rs("ZoneID"), rs("UpdateTime")) & "')""><IMG class=icon src=""images/jscript.gif"">" & rs("ZoneName") & "</DIV>"
                rs.movenext
            Loop
        End If
        rs.Close
        Set rs = Nothing
    %>
    </DIV>
    <!--  广告位标签结束  -->
    <!--  区域采集JS标签  -->
     <DIV class=topFolder id=Field><IMG id=FieldImg class=icon src="images/foldericon1.gif">区域采集JS标签</DIV>
     <DIV class=sub id=FieldSub>
     <%
        Set rs = Conn.Execute("select  AreaName,AreaFile from PE_AreaCollection where AreaPassed=" & PE_True & " Order by AreaID desc")
        If rs.bof And rs.EOF Then
            Response.write ("<li>您还没有建立区域采集JS</li>")
        Else
            Do While Not rs.EOF
                Response.write "<DIV class=subItem onclick=""InsertAdjs('/AreaCollection/" & rs("AreaFile") & "')""><IMG class=icon src=""images/jscript.gif"">" & rs("AreaName") & "</DIV>"
                rs.movenext
            Loop
        End If
        rs.Close
        Set rs = Nothing
        conn.Close
        Set conn = Nothing
      %>
    </DIV>
 </DIV>
  </td>
 </tr>
   </td>
  </tr>
</table>
<%
Function GetZoneJSName(iZoneID, UpdateTime)
	GetZoneJSName = Installdir & ADDir & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/" & iZoneID & ".js"
End Function
%>