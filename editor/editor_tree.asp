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

Dim ChannelID,sql,rs,ModuleType
Dim InsertTemplate,InsertTemplateType
ChannelID=PE_CLng(trim(request("ChannelID")))
InsertTemplate=PE_CLng(Request("insertTemplate"))
InsertTemplateType=PE_CLng(Request("InsertTemplateType"))

Sub BacktrackEditor() 
    If InsertTemplate=1 Then
        Response.Write  "if(label!=null){" & vbCrLf
        Response.Write "    parent.insertTemplateLabel(label," & InsertTemplateType & ");" & vbCrLf
        Response.Write  "}" & vbCrLf
    Else
        Response.Write "window.returnValue = label" & vbCrLf
        Response.Write "window.close();" & vbCrLf
    End if
End Sub

sql = "select ModuleType from PE_Channel where ChannelID=" & PE_CLng(ChannelID)
Set rs = Conn.Execute(sql)
    If rs.bof And rs.EOF Then
    else
        ModuleType=rs("ModuleType")
    End If
    rs.Close
Set rs = Nothing
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
            FolderImg.src="images/foldericon1.gif"
        }
        else {
            FolderImg.src="images/foldericon2.gif"
        }
    }
    ////////////////////////////////////////////////////////////////////////////////////////
    // 标签调用
    //普通标签
    function InsertLabel(label){
    <%
      Call BacktrackEditor()
    %>
    }
    //其它标签
    function InsertAdjs(type,fiflepath){
        var str="";
        switch(type){
        case "SwitchFont":
            str="<a name=StranLink href=''>切换到繁體中文</a>"
            break;
        case "Adjs":
            break;
        default:
            alert("错误参数调用！");
            break;
       }
    <%  If InsertTemplate=1 Then %>
            str=str+"<"+"script language=\"javascript\" src=\""+fiflepath+"\"></"+"script>"
            
    <%  Else %>
            str=str+"<IMG alt='#[!"+"script language=\"javascript\" src=\""+fiflepath+"\"!][!/"+"script!]#'  src=\"editor/images/jscript.gif\" border=0 $>"
    <%
        End if
        If InsertTemplate=1 Then
            Response.Write "parent.insertTemplateLabel(str," & InsertTemplateType & ");" & vbCrLf
        Else
            Response.write "window.returnValue =str" & vbCrLf
        End if
    %>
       window.close();
    }
    //函数标签调用
    function FunctionLabel(url,width,height){
        var label = showModalDialog(url, "", "dialogWidth:"+width+"px; dialogHeight:"+height+"px; help: no; scroll:no; status: no"); 
        <%
          Call BacktrackEditor()
        %>
    }
    //函数试标签
    function FunctionLabel2(name){
        var str,label
        switch(name){
        case "ShowTopUser":
            str=prompt("请输入显示注册用户列表的数量.","5"); 
            label="{$"+name+"("+str+")}"
            break;
        case "【ArticleList_ChildClass】":
            str=prompt("循环显示文章栏目录列表：每行显示的列数","2"); 
                if (str!=null) {
            label=name+"【Cols="+str+"】{$rsClass_ClassUrl} 栏目记录集中栏目地址 {$rsClass_Readme} 说明 {$rsClass_ClassName}名称  后面请您加上您自定义的标签【/ArticleList_ChildClass】"
            }
            break;
        case "【SoftList_ChildClass】":
            str=prompt("循环显示下载栏目录列表：每行显示的列数","2"); 
                if (str!=null) {
            label=name+"【Cols="+str+"】{$rsClass_ClassUrl} 栏目记录集中栏目地址 {$rsClass_Readme} 说明 {$rsClass_ClassName}名称  后面请您加上您自定义的标签【/SoftList_ChildClass】"
            }
            break;
        case "【PhotoList_ChildClass】":
            str=prompt("循环显示图片栏目录列表：每行显示的列数","2"); 
                if (str!=null) {
            label=name+"【Cols="+str+"】{$rsClass_ClassUrl} 栏目记录集中栏目地址 {$rsClass_Readme} 说明 {$rsClass_ClassName}名称  后面请您加上您自定义的标签【/PhotoList_ChildClass】"
            }
            break;
        case "【PositionList_Content】":
            str=prompt("循环显示职位内容信息列表：每页显示的职位数","6");
                if (str!=null) {
            label = name + "【PerPageNum=" + str + "】说明：请在此加上人才招聘内容标签（除申请职位按钮标签{$SaveSupply}）【/PositionList_Content】"
            }
            break;
        case "DownloadUrl":
            str=prompt("一行显示的列数","3");
                if (str!=null) {
            label = "{$"+name + "(" + str + ")}"
            }
            break;
        case "ResumeError":
            label = "\n<"+"SCRIPT LANGUAGE='JavaScript'>"
            label += "\n<!--"
            label += "\n function ResumeError() {"
            label += "\n return true;"
            label += "\n }"
            label += "\n window.onerror = ResumeError;"
            label += "\n // -->"
            label += "\n</"+"SCRIPT>"
            break;
        default:
            alert("错误参数调用！");
            break;
        }
        <%
          Call BacktrackEditor()
        %>
    }
    //动态函数试标签
    function FunctionLabel3(name){
        str=prompt("请输入动态函数标签参数.","5"); 
        label="{$"+name+"("+str+")}"
        <%
          Call BacktrackEditor()
        %>
    }
    //超级函数标签 
    function SuperFunctionLabel (url,label,title,ModuleType,ChannelShowType,iwidth,iheight){
        var label = showModalDialog(url+"?ChannelID=<%=ChannelID%>&Action=Add&LabelName="+label+"&Title="+title+"&ModuleType="+ModuleType+"&ChannelShowType="+ChannelShowType+"&InsertTemplate=<%=InsertTemplate%>", "", "dialogWidth:"+iwidth+"px; dialogHeight:"+iheight+"px; help: no; scroll:yes; status: yes"); 
        <%
          Call BacktrackEditor()
        %>
    }      
    </SCRIPT>
    <!-- 首页 -->
    <DIV class=topItem>
      <IMG class=icon height=16 src="images/home.gif" style="HEIGHT: 16px">标签调用
    </DIV>
    <!-- 系统介绍 -->
    <DIV class=favMenu id=aMenu>
    <!-- 通用标签 -->
    <DIV class=topFolder id=web><IMG id=webImg class=icon src="images/foldericon1.gif">网站通用标签</DIV>
    <DIV class=sub id=webSub>
        <!-- 网站通用普通标签 -->
        <DIV class=subFolder id=subwebInsert><IMG id=subwebInsertImg class=icon src="images/foldericon1.gif"> 网站通用标签</DIV>
        <DIV class=sub id=subwebInsertSub>
            <DIV class=subItem onClick="InsertLabel('{$PageTitle}')"><IMG class=icon src="images/label.gif">显示浏览器的标题栏显示页面的标题信息</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowChannel}')"><IMG class=icon src="images/label.gif">显示顶部频道信息</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowPath}')"><IMG class=icon src="images/label.gif">显示导航信息</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$ShowVote}')"><IMG class=icon src="images/label.gif">显示网站调查</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SiteName}')"><IMG class=icon src="images/label.gif">显示网站名称</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$SiteUrl}')"><IMG class=icon src="images/label.gif">显示网站地址</DIV>
            <DIV class=subItem onClick="InsertLabel('{$InstallDir}')"><IMG class=icon src="images/label.gif">系统安装目录</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowAdminLogin}')"><IMG class=icon src="images/label.gif">显示管理登录及链接</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Copyright}')"><IMG class=icon src="images/label.gif">显示版权信息</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Meta_Keywords}')"><IMG class=icon src="images/label.gif">针对搜索引擎的关键字</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Meta_Description}')"><IMG class=icon src="images/label.gif">针对搜索引擎的说明</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowSiteCountAll}')"><IMG class=icon src="images/label.gif">显示所有注册会员</DIV>
			<DIV class=subItem onClick="InsertLabel('{$GetUserName}')"><IMG class=icon src="images/label.gif">显示当前用户用户名</DIV>
            <DIV class=subItem onClick="InsertLabel('{$WebmasterName}')"><IMG class=icon src="images/label.gif">显示站长姓名</DIV>
            <DIV class=subItem onClick="InsertLabel('{$WebmasterEmail}')"><IMG class=icon src="images/label.gif">显示站长Email链接</DIV>
            <DIV class=subItem onClick="InsertLabel('{$MenuJS}')"><IMG class=icon src="images/label.gif">下拉栏目JS代码</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Skin_CSS}')"><IMG class=icon src="images/label.gif">风格CSS</DIV>
        </DIV>
        <!-- 网站通用函数普通标签结速标签 -->
        <DIV class=subFolder id=subwebFunction><IMG id=subwebFunctionImg class=icon src="images/foldericon1.gif"> 网站通用函数标签</DIV>
        <DIV class=sub id=subwebFunctionSub>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_Logo.htm','240','140')"><IMG class=icon src="images/label2.gif">显示网站LOGO</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_Banner.htm','240','140')"><IMG class=icon src="images/label2.gif">显示网站Banner</DIV>   
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_SlideJs.htm','300','400')"><IMG class=icon src="images/label2.gif">显示全站通用幻灯片标签</DIV>	
			<DIV class=subItem onClick="FunctionLabel('Lable/PE_IsLogin.htm','450','140')"><IMG class=icon src="images/label2.gif">显示登录状态判断标签</DIV>		           
			 <DIV class=subItem onClick="FunctionLabel('Lable/PE_YN.htm','500','500')"><IMG class=icon src="images/label2.gif">显示条件判断标签</DIV>     
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_RLanguage.htm','300','200')"><IMG class=icon src="images/label2.gif">读取语言包标签</DIV>     
            <DIV class=subItem onClick="FunctionLabel2('ShowTopUser')"><IMG class=icon src="images/label2.gif">显示注册用户列表</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_Annouce.htm','240','140')"><IMG class=icon src="images/label2.gif">显示本站公告信息</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_Annouce2.htm','240','210')"><IMG class=icon src="images/label2.gif">显示详细公告信息</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_FSite.htm','330','260')"><IMG class=icon src="images/label2.gif">显示友情链接信息</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_FSite2.htm','330','510')"><IMG class=icon src="images/label2.gif">显示详细友情链接信息</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_ProducerList.htm','400','450')"><IMG class=icon src="images/label2.gif">显示厂商列表</DIV> 
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_Author_ShowList.htm','400','340')"><IMG class=icon src="images/label2.gif">显示作者列表</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_ShowSpecialList.htm','320','300')"><IMG class=icon src="images/label2.gif">显示指定频道专题</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_ShowBlogList.htm','400','400')"><IMG class=icon src="images/label2.gif">显示作品集排行</DIV>
        </DIV>
    </DIV>
    <!-- 频道通用标签 -->
    <DIV class=topFolder id=ChannelCommon><IMG id=ChannelCommonImg class=icon src="images/foldericon1.gif">频道通用标签</DIV>
    <DIV class=sub id=ChannelCommonSub>
        <DIV class=subItem onClick="InsertLabel('{$ChannelName}')"><IMG class=icon src="images/label.gif">显示频道名称</DIV>    
        <DIV class=subItem onClick="InsertLabel('{$ChannelID}')"><IMG class=icon src="images/label.gif">得到频道ID</DIV>    
        <DIV class=subItem onClick="InsertLabel('{$ChannelDir}')"><IMG class=icon src="images/label.gif">得到频道目录</DIV>    
        <DIV class=subItem onClick="InsertLabel('{$ChannelUrl}')"><IMG class=icon src="images/label.gif">频道目录路径</DIV>
        <DIV class=subItem onClick="InsertLabel('{$UploadDir}')"><IMG class=icon src="images/label.gif">频道上传目录路径</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ChannelPicUrl}')"><IMG class=icon src="images/label.gif">频道图片</DIV>
        <DIV class=subItem onClick="InsertLabel('{$Meta_Keywords_Channel}')"><IMG class=icon src="images/label.gif">针对搜索引擎的关键字</DIV>
        <DIV class=subItem onClick="InsertLabel('{$Meta_Description_Channel}')"><IMG class=icon src="images/label.gif">针对搜索引擎的说明</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ChannelShortName}')"><IMG class=icon src="images/label.gif">显示频道名</DIV>    
        <DIV class=subItem onClick="FunctionLabel('Lable/PE_ClassNavigation.htm','260','200')"><IMG class=icon src="images/label2.gif">显示栏目导航的HTML代码</DIV>
    </DIV>
    <!-- 频道专用页标签 -->
    <DIV class=topFolder id=Channel><IMG id=ChannelImg class=icon src="images/foldericon1.gif">频道专用标签</DIV>
    <DIV class=sub id=ChannelSub>
        <DIV class=subItem onClick="FunctionLabel('Lable/PE_AnnWin.htm','240','140')"><IMG class=icon src="images/label2.gif">弹出公告窗口</DIV>    
        <DIV class=subItem onClick="InsertLabel('{$ClassListUrl}')"><IMG class=icon src="images/label.gif">模板中“更多”处的链接</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ShowChildClass}')"><IMG class=icon src="images/label.gif">显示一级栏目下第二级栏目名</DIV>
        <DIV class=subItem onClick="FunctionLabel('Lable/PE_ShowChildClass.htm','330','360')"><IMG class=icon src="images/label2.gif">显示当前栏目的下一级子栏目</DIV>
        <DIV class=subItem onClick="FunctionLabel('Lable/PE_ShowBrotherClass.htm','330','360')"><IMG class=icon src="images/label2.gif">显示当前栏目的同级栏目</DIV>
	<DIV class=subItem onClick="InsertLabel('{$ParentDir}')"><IMG class=icon src="images/label.gif">得到当前栏目的父目录</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ClassDir}')"><IMG class=icon src="images/label.gif">得到当前栏目的目录</DIV>
        <DIV class=subItem onClick="InsertLabel('{$Readme}')"><IMG class=icon src="images/label.gif">得到当前栏目的说明</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ClassUrl}')"><IMG class=icon src="images/label.gif">得到当前栏目的链接地址</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ClassPicUrl}')"><IMG class=icon src="images/label.gif">栏目图片</DIV>
        <DIV class=subItem onClick="InsertLabel('{$Meta_Keywords_Class}')"><IMG class=icon src="images/label.gif">针对搜索引擎的关键字</DIV>
        <DIV class=subItem onClick="InsertLabel('{$Meta_Description_Class}')"><IMG class=icon src="images/label.gif">针对搜索引擎的说明</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ClassName}')"><IMG class=icon src="images/label.gif">显示当前栏目的名称</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ClassID}')"><IMG class=icon src="images/label.gif">得到当前栏目的ID</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ShowChannelCount}')"><IMG class=icon src="images/label.gif">显示频道统计信息</DIV>
        <DIV class=subItem onClick="InsertLabel('{$SpecialName}')"><IMG class=icon src="images/label.gif">显示当前专题的专题名称</DIV>
        <DIV class=subItem onClick="InsertLabel('{$SpecialPicUrl}')"><IMG class=icon src="images/label.gif">显示专题图片</DIV>
        <DIV class=subItem onClick="InsertLabel('{$GetAllSpecial}')"><IMG class=icon src="images/label.gif">显示全部专题</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ShowPage}')"><IMG class=icon src="images/label.gif">显示分页标签</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ShowPage_en}')"><IMG class=icon src="images/label.gif">显示英文分页标签</DIV>
        <DIV class=subItem onClick="InsertLabel('{$InstallDir}{$ChannelDir}')"><IMG class=icon src="images/label.gif">频道安装目录</DIV>
    </DIV>
    <!-- 频道搜索页标签 -->
    <DIV class=topFolder id=ChannelSearch><IMG id=ChannelSearchImg class=icon src="images/foldericon1.gif">频道搜索页标签</DIV>
    <DIV class=sub id=ChannelSearchSub>
        <DIV class=subItem onClick="InsertLabel('{$ResultTitle}')"><IMG class=icon src="images/label.gif">显示搜索的是什么内容信息</DIV>
        <DIV class=subItem onClick="InsertLabel('{$SearchResult}')"><IMG class=icon src="images/label.gif">搜索结果</DIV>
        <DIV class=subItem onClick="InsertLabel('{$Keyword}')"><IMG class=icon src="images/label.gif">搜索关键字</DIV>
    </DIV>
    <!-- 内容页通用标签 -->
    <DIV class=topFolder id=ContentCommon><IMG id=ContentCommonImg class=icon src="images/foldericon1.gif">内容页通用标签</DIV>
    <DIV class=sub id=ContentCommonSub>
        <DIV class=subItem onClick="InsertLabel('{$ClassID}')"><IMG class=icon src="images/label.gif">得到当前栏目的ID</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ClassName}')"><IMG class=icon src="images/label.gif">显示当前栏目的名称</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ClassDir}')"><IMG class=icon src="images/label.gif">得到当前栏目的目录</DIV>
        <DIV class=subItem onClick="InsertLabel('{$Readme}')"><IMG class=icon src="images/label.gif">得到当前栏目的说明</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ClassUrl}')"><IMG class=icon src="images/label.gif">得到当前栏目的链接地址</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ParentDir}')"><IMG class=icon src="images/label.gif">得到当前栏目的父目录</DIV>
        <DIV class=subItem onClick="InsertLabel('{$Keyword}')"><IMG class=icon src="images/label.gif">搜索关键字</DIV>
    </DIV>

<% if ModuleType=1 or ModuleType=0 then %>
    <!-- 文章频道标签 -->
     <DIV class=topFolder id=Article><IMG id=ArticleImg class=icon src="images/foldericon1.gif">文章标签</DIV>
     <DIV class=sub id=ArticleSub>
        <!-- 文章通用频道标签 -->
        <DIV class=subFolder id=subArticleChannelFunction><IMG id=subArticleChannelFunctionImg class=icon src="images/foldericon1.gif"> 文章频道标签</DIV>
        <DIV class=sub id=subArticleChannelFunctionSub>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_label.asp','GetArticleList','文章列表函数标签',1,'GetList',800,700)"><IMG class=icon src="images/label3.gif">显示文章标题等信息</DIV>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_label.asp','GetPicArticle','显示图片文章标签',1,'GetPic',700,500)"><IMG class=icon src="images/label3.gif">显示图片文章</DIV>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_label.asp','GetSlidePicArticle','显示幻灯片文章标签',1,'GetSlide',700,500)"><IMG class=icon src="images/label3.gif">显示幻灯片文章</DIV>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_CustomListLabel.asp','CustomListLable','文章自定义列表标签',1,'GetArticleCustom',720,700)"><IMG class=icon src="images/label3.gif">文章自定义列表标签</DIV>
        </Div>
        <DIV class=subFolder id=subArticleClass><IMG id=subArticleClassImg class=icon src="images/foldericon1.gif"> 文章栏目标签</DIV>
        <DIV class=sub id=subArticleClassSub>
            <DIV class=subItem onClick="FunctionLabel2('【ArticleList_ChildClass】')"><IMG class=icon src="images/label2.gif">循环显示文章栏目录列表</DIV> 
            <DIV class=subItem onClick="InsertLabel('【ArticleList_CurrentClass】{$rsClass_ClassUrl} 栏目记录集中栏目地址 {$rsClass_Readme}说明 {$rsClass_ClassName}名称  后面请您加上您自定义的标签【/ArticleList_CurrentClass】')"><IMG class=icon src="images/label.gif">当前栏目列表(同时存在文章及子栏目)循环标签</DIV>
        </DIV>
         <!-- 文章通用频道标签结束 -->
         <!-- 文章频道内容标签 -->
         <DIV class=subFolder id=subArticleChannelContent><IMG id=subArticleChannelContentImg class=icon src="images/foldericon1.gif"> 文章内容标签</DIV>
         <DIV class=sub id=subArticleChannelContentSub>
            <DIV class=subItem onClick="InsertLabel('{$ArticleID}')"><IMG class=icon src="images/label.gif">当前文章的I D</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ArticleProtect}')"><IMG class=icon src="images/label.gif">根据频道设置得到防复制功能的代码</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ArticleProperty}')"><IMG class=icon src="images/label.gif">显示当前文章的属性（热门、推荐、等级）</DIV>        
            <DIV class=subItem onClick="InsertLabel('{$ArticleTitle}')"><IMG class=icon src="images/label.gif">显示文章正标题</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ArticleSign}')"><IMG class=icon src="images/label.gif">自动签收文章</DIV>			
            <DIV class=subItem onClick="InsertLabel('{$ArticleUrl}')"><IMG class=icon src="images/label.gif">显示文章网址</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_InputerInfo.htm','380','200')"><IMG class=icon src="images/label.gif">读取文章录入者信息</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ArticleTitle2}')"><IMG class=icon src="images/label.gif">显示文章显示页导航处当前文章标题信息</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ArticleInfo}')"><IMG class=icon src="images/label.gif">整体显示文章作者、文章来源、点击数、更新时间信息</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ArticleSubheading}')"><IMG class=icon src="images/label.gif">显示文章副标题</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Subheading}')"><IMG class=icon src="images/label.gif">自定义列表副标题</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$ReadPoint}')"><IMG class=icon src="images/label.gif">阅读点数</DIV>            
            <DIV class=subItem onClick="InsertLabel('{$Author}')"><IMG class=icon src="images/label.gif">作者</DIV>
            <DIV class=subItem onClick="InsertLabel('{$CopyFrom}')"><IMG class=icon src="images/label.gif">文章来源</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Editor}')"><IMG class=icon src="images/label.gif">责任编辑</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Hits}')"><IMG class=icon src="images/label.gif">点击数</DIV>
            <DIV class=subItem onClick="InsertLabel('{$UpdateTime}')"><IMG class=icon src="images/label.gif">更新时间信息</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$ArticleIntro}')"><IMG class=icon src="images/label.gif">显示文章简介</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ArticleContent}')"><IMG class=icon src="images/label.gif">显示文章的具体的内容</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PrevArticle}')"><IMG class=icon src="images/label.gif">显示上一篇文章</DIV>
            <DIV class=subItem onClick="InsertLabel('{$NextArticle}')"><IMG class=icon src="images/label.gif">显示下一篇文章</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ArticleEditor}')"><IMG class=icon src="images/label.gif">显示文章录入和责任编辑姓名</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ArticleAction}')"><IMG class=icon src="images/label.gif">显示【发表评论】【告诉好友】【打印此文】【关闭窗口】</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_CorrelativeArticle.htm','280','385')"><IMG class=icon src="images/label2.gif">显示相关文章</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ManualPagination}')"><IMG class=icon src="images/label.gif">采用手动分页方式显示文章具体的内容</DIV>
            <DIV class=subItem onClick="InsertLabel('{$AutoPagination}')"><IMG class=icon src="images/label.gif">采用自动分页方式显示文章具体的内容</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Vote}')"><IMG class=icon src="images/label.gif">显示调查</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_GetSubTitleHtml.htm','340','200')"><IMG class=icon src="images/label2.gif">正文分页导航</DIV>
        </DIV>
        <!-- 文章频道内容标签结束 -->
    </DIV>
    <!-- 文章频道标签结束 -->
    <%
    End if
    if  ModuleType=2 or ModuleType=0 then %>
    <!-- 下载频道标签 -->
    <DIV class=topFolder id=Soft><IMG id=SoftImg class=icon src="images/foldericon1.gif">下载标签</DIV>
    <DIV class=sub id=SoftSub>
         <!-- 下载通用频道标签 -->
         <DIV class=subFolder id=subSoftChannelFunction><IMG id=subSoftChannelFunctionImg class=icon src="images/foldericon1.gif"> 下载频道标签</DIV>
         <DIV class=sub id=subSoftChannelFunctionSub>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_label.asp','GetSoftList','下载列表函数标签',2,'GetList',800,700)"><IMG class=icon src="images/label3.gif">显示软件标题等信息</DIV>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_label.asp','GetPicSoft','显示图片下载标签',2,'GetPic',700,500)"><IMG class=icon src="images/label3.gif">显示图片下载</DIV>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_label.asp','GetSlidePicSoft','显示幻灯片下载标签',2,'GetSlide',700,500)"><IMG class=icon src="images/label3.gif">显示幻灯片下载</DIV>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_CustomListLabel.asp','CustomListLable','下载自定义列表标签',2,'GetSoftCustom',720,700)"><IMG class=icon src="images/label3.gif">下载自定义列表标签</DIV>
        </DIV>
        <DIV class=subFolder id=subSoftClassFunction><IMG id=subSoftClassFunctionImg class=icon src="images/foldericon1.gif"> 下载栏目标签</DIV>
        <DIV class=sub id=subSoftClassFunctionSub>
            <DIV class=subItem onClick="FunctionLabel2('【SoftList_ChildClass】')"><IMG class=icon src="images/label2.gif"> 循环显示下载栏目录列表</DIV>
            <DIV class=subItem onClick="InsertLabel('【SoftList_CurrentClass】{$rsClass_ClassUrl} 栏目记录集中栏目地址 {$rsClass_Readme}说明 {$rsClass_ClassName}名称  后面请您加上您自定义的标签【/SoftList_CurrentClass】')"><IMG class=icon src="images/label.gif">当前栏目列表(同时存在下载及子栏目)循环标签</DIV>
        </DIV>
        <!-- 下载通用频道标签结束 -->
        <!-- 下载频道内容标签 -->
        <DIV class=subFolder id=subSoftChannelContent><IMG id=subSoftChannelContentImg class=icon src="images/foldericon1.gif"> 下载内容标签</DIV>
        <DIV class=sub id=subSoftChannelContentSub>
            <DIV class=subItem onClick="InsertLabel('{$SoftID}')"><IMG class=icon src="images/label.gif">软件ID</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SoftName}')"><IMG class=icon src="images/label.gif">软件名称</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SoftVersion}')"><IMG class=icon src="images/label.gif">显示软件版本</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SoftSize} K')"><IMG class=icon src="images/label.gif">软件文件大小</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SoftSize_M}')"><IMG class=icon src="images/label.gif">显示软件大小（以M 为单位）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$DecompressPassword}')"><IMG class=icon src="images/label.gif">解压密码</DIV>
            <DIV class=subItem onClick="InsertLabel('{$OperatingSystem}')"><IMG class=icon src="images/label.gif">运行平台</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Hits}')"><IMG class=icon src="images/label.gif">下载次数总计</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Author}')"><IMG class=icon src="images/label.gif">开 发 商</DIV>
            <DIV class=subItem onClick="InsertLabel('{$DayHits}')"><IMG class=icon src="images/label.gif">下载次数本日</DIV>
            <DIV class=subItem onClick="InsertLabel('{$WeekHits}')"><IMG class=icon src="images/label.gif">下载次数本周</DIV>
            <DIV class=subItem onClick="InsertLabel('{$MonthHits}')"><IMG class=icon src="images/label.gif">下载次数本月</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Stars}')"><IMG class=icon src="images/label.gif">软件等级</DIV>
            <DIV class=subItem onClick="InsertLabel('{$CopyFrom}')"><IMG class=icon src="images/label.gif">软件来源</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SoftLink}')"><IMG class=icon src="images/label.gif">显示软件的演示地址和注册地址</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SoftType}')"><IMG class=icon src="images/label.gif">软件类别</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SoftLanguage}')"><IMG class=icon src="images/label.gif">软件语言</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SoftProperty}')"><IMG class=icon src="images/label.gif">软件属性</DIV>
            <DIV class=subItem onClick="InsertLabel('{$UpdateTime}')"><IMG class=icon src="images/label.gif">更新时间</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Editor}')"><IMG class=icon src="images/label.gif">软件添加审核</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Inputer}')"><IMG class=icon src="images/label.gif">软件添加录入</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SoftPicUrl}')"><IMG class=icon src="images/label.gif">显示下载图片</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_SoftPic.htm','240','140')"><IMG class=icon src="images/label2.gif">显示下载图片详细</DIV>
            <DIV class=subItem onClick="InsertLabel('{$DemoUrl}')"><IMG class=icon src="images/label.gif">显示演示地址</DIV>
            <DIV class=subItem onClick="InsertLabel('{$RegUrl}')"><IMG class=icon src="images/label.gif">显示注册地址</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SoftPoint}')"><IMG class=icon src="images/label.gif">显示收费软件的下载点数</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$CopyrightType}')"><IMG class=icon src="images/label.gif">授权方式</DIV>    
            <DIV class=subItem onClick="FunctionLabel2('DownloadUrl')"><IMG class=icon src="images/label2.gif">软件下载地址</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SoftIntro}')"><IMG class=icon src="images/label.gif">软件简介</DIV>
            <DIV class=subItem onClick="InsertLabel('{$CorrelativeSoft}')"><IMG class=icon src="images/label.gif">相关软件</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SoftAuthor}')"><IMG class=icon src="images/label.gif">显示软件作者</DIV>
            <DIV class=subItem onClick="InsertLabel('{$AuthorEmail}')"><IMG class=icon src="images/label.gif">显示作者Email</DIV>
            <DIV class=subItem onClick="InsertLabel('{$BrowseTimes}')"><IMG class=icon src="images/label.gif">显示软件的浏览量</DIV>
        </DIV>
        <!-- 下载频道内容标签结束 -->
    </DIV>
    <%
    End if
    If  ModuleType=3 or ModuleType=0 then %>
    <!-- 图片频道标签 -->
     <DIV class=topFolder id=Photo><IMG id=PhotoImg class=icon src="images/foldericon1.gif">图片标签</DIV>
     <DIV class=sub id=PhotoSub>
        <!-- 图片通用频道标签 -->
        <DIV class=subFolder id=subPhotoChannelFunction><IMG id=subPhotoChannelFunctionImg class=icon src="images/foldericon1.gif"> 图片频道标签</DIV>
        <DIV class=sub id=subPhotoChannelFunctionSub>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_label.asp','GetPhotoList','图片列表函数标签',3,'GetList',800,700)"><IMG class=icon src="images/label3.gif">显示图片标题等信息</DIV>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_label.asp','GetPicPhoto','显示图片图文标签',3,'GetPic',700,550)"><IMG class=icon src="images/label3.gif">显示图片</DIV>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_label.asp','GetSlidePicPhoto','显示幻灯片图片标签',3,'GetSlide',700,550)"><IMG class=icon src="images/label3.gif">显示幻灯片图片</DIV>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_CustomListLabel.asp','CustomListLable','图片自定义列表标签',3,'GetPhotoCustom',720,700)"><IMG class=icon src="images/label3.gif">图片自定义列表标签</DIV>
        </DIV>
        <DIV class=subFolder id=subPhotoClassFunction><IMG id=subPhotoClassFunctionImg class=icon src="images/foldericon1.gif"> 图片栏目标签</DIV>
        <DIV class=sub id=subPhotoClassFunctionSub>
            <DIV class=subItem onClick="FunctionLabel2('【PhotoList_ChildClass】')"><IMG class=icon src="images/label2.gif">循环显示图片栏目录列表</DIV>
            <DIV class=subItem onClick="InsertLabel('【PhotoList_CurrentClass】{$rsClass_ClassUrl} 栏目记录集中栏目地址 {$rsClass_Readme}说明 {$rsClass_ClassName}名称  后面请您加上您自定义的标签【/PhotoList_CurrentClass】')"><IMG class=icon src="images/label.gif">当前栏目列表(同时存在图片及子栏目)循环标签</DIV>
        </DIV>
        <!-- 图片频道通用标签结束 -->
        <!-- 图片频道内容标签 -->
        <DIV class=subFolder id=subPhotoChannelContent><IMG id=subPhotoChannelContentImg class=icon src="images/foldericon1.gif"> 图片内容标签</DIV>
        <DIV class=sub id=subPhotoChannelContentSub>
            <DIV class=subItem onClick="InsertLabel('{$PhotoID}')"><IMG class=icon src="images/label.gif">图片I D</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PhotoName}')"><IMG class=icon src="images/label.gif">显示图片名称</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Hits}')"><IMG class=icon src="images/label.gif">查看次数总计</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Author}')"><IMG class=icon src="images/label.gif">图片作者</DIV>
            <DIV class=subItem onClick="InsertLabel('{$CopyFrom}')"><IMG class=icon src="images/label.gif">图片来源</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PhotoProperty}')"><IMG class=icon src="images/label.gif">显示图片属性</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Stars}')"><IMG class=icon src="images/label.gif">推荐等级</DIV>
            <DIV class=subItem onClick="InsertLabel('{$UpdateTime}')"><IMG class=icon src="images/label.gif">更新时间</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Editor}')"><IMG class=icon src="images/label.gif">显示图片的责任编辑</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Inputer}')"><IMG class=icon src="images/label.gif">显示图片录入者</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PhotoPoint}')"><IMG class=icon src="images/label.gif">收费图片所需的点数</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PhotoIntro}')"><IMG class=icon src="images/label.gif">显示图片简介</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PrevPhotoUrl}')"><IMG class=icon src="images/label.gif">上一张图片的链接地址</DIV>
            <DIV class=subItem onClick="InsertLabel('{$NextPhotoUrl}')"><IMG class=icon src="images/label.gif">下一张图片的链接地址</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ViewPhoto}')"><IMG class=icon src="images/label.gif">显示图片或Flash</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_PhotoUrlList.htm','300','270')"><IMG class=icon src="images/label2.gif">显示图片地址列表</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PhotoUrl}')"><IMG class=icon src="images/label.gif">图片地址列表中的第一个地址</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_CorrelativePhoto.htm','240','140')"><IMG class=icon src="images/label2.gif">相关图片列表</DIV>        
            <DIV class=subItem onClick="InsertLabel('{$PhotoSize} K')"><IMG class=icon src="images/label.gif">图片大小</DIV>
            <DIV class=subItem onClick="InsertLabel('{$DayHits}')"><IMG class=icon src="images/label.gif">查看次数本日</DIV>
            <DIV class=subItem onClick="InsertLabel('{$WeekHits}')"><IMG class=icon src="images/label.gif">查看次数本周</DIV>
            <DIV class=subItem onClick="InsertLabel('{$MonthHits}')"><IMG class=icon src="images/label.gif">查看次数本月</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$PhotoThumb}')"><IMG class=icon src="images/label.gif">显示图片缩略图</DIV>
            <DIV class=subItem onClick="InsertLabel('{$GetUrlArray}')"><IMG class=icon src="images/label.gif">获取图片地址的初始化JS</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_PhotoThumb.htm','240','140')"><IMG class=icon src="images/label2.gif">显示指定大小的图片缩略图</DIV>
        </DIV>
        <!-- 图片频道内容标签结束 -->
     </DIV>
    <%
    End if
    if  ModuleType=4 or ModuleType=0 then %>
    <!--  留言频道函数  -->
     <DIV class=topFolder id=Guest><IMG id=GuestImg class=icon src="images/foldericon1.gif">留言函数</DIV>
     <DIV class=sub id=GuestSub>
        <!-- 留言板通用标签 -->
        <DIV class=subFolder id=subGuestCommonFunction><IMG id=subGuestCommonFunctionImg class=icon src="images/foldericon1.gif">留言板通用标签</DIV>
        <DIV class=sub id=subGuestCommonFunctionSub>
            <DIV class=subItem onClick="InsertLabel('{$GetGKindList}')"><IMG class=icon src="images/label.gif">显示留言类别横向导航</DIV>
            <DIV class=subItem onClick="InsertLabel('{$GuestBook_top}')"><IMG class=icon src="images/label.gif">显示顶部功能菜单</DIV>
            <DIV class=subItem onClick="InsertLabel('{$GuestBook_Mode}')"><IMG class=icon src="images/label.gif">显示留言模式（游客/ 会员模式）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$GuestBook_See}')"><IMG class=icon src="images/label.gif">显示留言查看模式（留言板/ 讨论区模式）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$GuestBook_Appear}')"><IMG class=icon src="images/label.gif">显示留言发表模式（审核发表/ 直接发表）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowGueststyle}')"><IMG class=icon src="images/label.gif">切换到另一种查看方式</DIV>
            <DIV class=subItem onClick="InsertLabel('{$GuestBook_Search}')"><IMG class=icon src="images/label.gif">显示留言搜索表单</DIV>
        </DIV>
        <!-- 留言板通用标签 -->
        <DIV class=subFolder id=subGuestIndexFunction><IMG id=subGuestIndexFunctionImg class=icon src="images/foldericon1.gif">留言板通用标签</DIV>
        <DIV class=sub id=subGuestIndexFunctionSub>
            <DIV class=subItem onClick="InsertLabel('{$GuestMain}')"><IMG class=icon src="images/label.gif">显示留言列表</DIV>    
        </DIV>
        <!-- 编辑留言页标签 -->
        <DIV class=subFolder id=subGuestEditFunction><IMG id=subGuestEditFunctionImg class=icon src="images/foldericon1.gif">编辑留言页标签</DIV>
        <DIV class=sub id=subGuestEditFunctionSub>
            <DIV class=subItem onClick="InsertLabel('{$WriteGuest}')"><IMG class=icon src="images/label.gif">签写留言</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowJS_Guest}')"><IMG class=icon src="images/label.gif">留言Js验证</DIV>
            <DIV class=subItem onClick="InsertLabel('{$WriteTitle}')"><IMG class=icon src="images/label.gif">显示留言标题</DIV>
            <DIV class=subItem onClick="InsertLabel('{$GetGKind_Option}')"><IMG class=icon src="images/label.gif">显示留言类别</DIV>
            <DIV class=subItem onClick="InsertLabel('{$GuestFace}')"><IMG class=icon src="images/label.gif">显示留言心情</DIV>
            <DIV class=subItem onClick="InsertLabel('{$GuestContent}')"><IMG class=icon src="images/label.gif">显示留言内容</DIV>
            <DIV class=subItem onClick="InsertLabel('{$saveedit}')"><IMG class=icon src="images/label.gif">标记是否为编辑留言</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ReplyId}')"><IMG class=icon src="images/label.gif">回复主题id</DIV>
            <DIV class=subItem onClick="InsertLabel('{$saveeditid}')"><IMG class=icon src="images/label.gif">要编辑留言的ID</DIV>
        </DIV>
        <!-- 留言回复页标签 -->
        <DIV class=subFolder id=subGuestReplyFunction><IMG id=subGuestReplyFunctionImg class=icon src="images/foldericon1.gif">留言回复页标签</DIV>
        <DIV class=sub id=subGuestReplyFunctionSub>
            <DIV class=subItem onClick="InsertLabel('{$ReplyGuest}')"><IMG class=icon src="images/label.gif">回复留言主函数</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowJS_Guest}')"><IMG class=icon src="images/label.gif">留言Js验证</DIV>
            <DIV class=subItem onClick="InsertLabel('{$WriteTitle}')"><IMG class=icon src="images/label.gif">显示留言标题</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ReplyId}')"><IMG class=icon src="images/label.gif">回复主题id</DIV>
        </DIV>
        <!-- 留言搜索页标签 -->
        <DIV class=subFolder id=subGuestSearchFunction><IMG id=subGuestSearchFunctionImg class=icon src="images/foldericon1.gif">留言搜索页标签</DIV>
        <DIV class=sub id=subGuestSearchFunctionSub>
            <DIV class=subItem onClick="InsertLabel('{$ResultTitle}')"><IMG class=icon src="images/label.gif">搜索结果标题</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SearchResult}')"><IMG class=icon src="images/label.gif">搜索结果</DIV>
        </DIV>
     </DIV>
    <%
    End if
    if ModuleType=5 or ModuleType=0 then%>
    <!--  商城频道标签  -->
     <DIV class=topFolder id=Shop><IMG id=ShopImg class=icon src="images/foldericon1.gif">商城标签</DIV>
     <DIV class=sub id=ShopSub>
        <!-- 商城通用频道标签 -->
        <DIV class=subFolder id=subShopChannelFunction><IMG id=subShopChannelFunctionImg class=icon src="images/foldericon1.gif"> 商城通用标签</DIV>
        <DIV class=sub id=subShopChannelFunctionSub>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_label.asp','GetProductList','商城列表函数标签',5,'GetList',800,750)"><IMG class=icon src="images/label3.gif">显示商品标题等信息</DIV>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_label.asp','GetPicProduct','显示图片商城标签',5,'GetPic',700,600)"><IMG class=icon src="images/label3.gif">显示图片商品</DIV>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_label.asp','GetSlidePicProduct','显示幻灯片商城标签',5,'GetSlide',700,460)"><IMG class=icon src="images/label3.gif">显示幻灯片商品</DIV>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_CustomListLabel.asp','CustomListLable','商城自定义列表标签',5,'GetProductCustom',720,700)"><IMG class=icon src="images/label3.gif">商品自定义列表标签</DIV>
        </DIV>
        <!--  商城频内容页标签 -->
        <DIV class=subFolder id=subshopcontent><IMG id=subshopcontentImg class=icon src="images/foldericon1.gif"> 商城内容标签</DIV>
        <DIV class=sub id=subshopcontentSub>
            <DIV class=subItem onClick="InsertLabel('{$ProductID}')"><IMG class=icon src="images/label.gif">商品ID</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ProductName}')"><IMG class=icon src="images/label.gif">商品名称</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ProductNum}')"><IMG class=icon src="images/label.gif">商品数量</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ProductModel}')"><IMG class=icon src="images/label.gif">商品型号</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ProductStandard}')"><IMG class=icon src="images/label.gif">商品规格</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ProducerName}')"><IMG class=icon src="images/label.gif">生 产 商</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PresentExp}')"><IMG class=icon src="images/label.gif">购物积分</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PresentPoint}')"><IMG class=icon src="images/label.gif">赠送点券</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PresentMoney}')"><IMG class=icon src="images/label.gif">返还的现金券</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PointName}')"><IMG class=icon src="images/label.gif">点券的名称</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PointUnit}')"><IMG class=icon src="images/label.gif">点券的单位</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Stocks}')"><IMG class=icon src="images/label.gif">显示库存量</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ServiceTerm}')"><IMG class=icon src="images/label.gif">提供服务时间</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$TrademarkName}')"><IMG class=icon src="images/label.gif">品牌商标</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ProductTypeName}')"><IMG class=icon src="images/label.gif">商品类别</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Price_Your}')"><IMG class=icon src="images/label.gif">当前访问者的价格</DIV>
            <DIV class=subItem onClick="InsertLabel('{$UpdateTime}')"><IMG class=icon src="images/label.gif">上架时间</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_PhotoThumb.htm','240','140')"><IMG class=icon src="images/label2.gif">商品缩列图</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Hits}')"><IMG class=icon src="images/label.gif">商品点击数</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$ProductProperty}')"><IMG class=icon src="images/label.gif">显示当前商品的属性（热门、推荐、等级）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ProductIntro}')"><IMG class=icon src="images/label.gif">商品简介</DIV>
            <DIV class=subItem onClick="InsertLabel('{$OnTop}')"><IMG class=icon src="images/label.gif">显示固顶</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Hot}')"><IMG class=icon src="images/label.gif">显示热卖</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Elite}')"><IMG class=icon src="images/label.gif">显示推荐</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Stars}')"><IMG class=icon src="images/label.gif">推荐等级</DIV>
            <DIV class=subItem onClick="InsertLabel('{$LimitNum}')"><IMG class=icon src="images/label.gif">限够数量</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Discount}')"><IMG class=icon src="images/label.gif">降价折扣</DIV>
            <DIV class=subItem onClick="InsertLabel('{$BeginDate}～{$EndDate}')"><IMG class=icon src="images/label.gif">显示优惠日期</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Price_Original}')"><IMG class=icon src="images/label.gif">原始零售价</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Price_Market}')"><IMG class=icon src="images/label.gif">显示市场价</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Price}')"><IMG class=icon src="images/label.gif">显示商城价</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Price_Member}')"><IMG class=icon src="images/label.gif">显示会员价</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Unit}')"><IMG class=icon src="images/label.gif">显示商品单位</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SalePromotion}')"><IMG class=icon src="images/label.gif">显示促销方案</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ProductExplain}')"><IMG class=icon src="images/label.gif">显示商品说明</DIV>
            <DIV class=subItem onClick="InsertLabel('{$CorrelativeProduct}')"><IMG class=icon src="images/label.gif">显示相关商品</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Vote}')"><IMG class=icon src="images/label.gif">显示调查</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_CorrelativeProduct.htm','240','260')"><IMG class=icon src="images/label2.gif">显示详细相关商品</DIV>
        </DIV>
        <!-- 商城频内容页标签结束  -->
        <!-- 购物车标签 -->
        <DIV class=subFolder id=subshopping><IMG id=subshoppingImg class=icon src="images/foldericon1.gif"> 我的购物车</DIV>
        <DIV class=sub id=subshoppingSub>
            <DIV class=subItem onClick="InsertLabel('{$ShowTips_Login}')"><IMG class=icon src="images/label.gif">用户登录提示</DIV>            
            <DIV class=subItem onClick="InsertLabel('{$UserName}')"><IMG class=icon src="images/label.gif">用户名</DIV>
            <DIV class=subItem onClick="InsertLabel('{$GroupName}')"><IMG class=icon src="images/label.gif">用户组</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Discount_Member}%')"><IMG class=icon src="images/label.gif">会员折扣率</DIV>
            <DIV class=subItem onClick="InsertLabel('{$IsOffer}')"><IMG class=icon src="images/label.gif">享受折上折优惠</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Balance}')"><IMG class=icon src="images/label.gif">资金余额</DIV>
            <DIV class=subItem onClick="InsertLabel('{$UserPoint}')"><IMG class=icon src="images/label.gif">可用点数</DIV>
            <DIV class=subItem onClick="InsertLabel('{$UserExp}')"><IMG class=icon src="images/label.gif">可用积分</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowTips_CartIsEmpty}')"><IMG class=icon src="images/label.gif">购物车提示</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowCart}')"><IMG class=icon src="images/label.gif">显示购物车中的商品</DIV>
        </DIV>
        <!-- 购物车标签结束 -->
        <!-- 收银台标签 -->
        <DIV class=subFolder id=subshopcash><IMG id=subshopcashImg class=icon src="images/foldericon1.gif"> 收　银　台</DIV>
        <DIV class=sub id=subshopcashSub>
            <DIV class=subItem onClick="InsertLabel('{$ShowTips_Login}')"><IMG class=icon src="images/label.gif">用户登录提示</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$ShowTips_CartIsEmpty}')"><IMG class=icon src="images/label.gif">购物车提示</DIV>        
            <DIV class=subItem onClick="InsertLabel('{$UserName}')"><IMG class=icon src="images/label.gif">用户名</DIV>
            <DIV class=subItem onClick="InsertLabel('{$GroupName}')"><IMG class=icon src="images/label.gif">用户组</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Discount_Member}%')"><IMG class=icon src="images/label.gif">会员折扣率</DIV>
            <DIV class=subItem onClick="InsertLabel('{$IsOffer}')"><IMG class=icon src="images/label.gif">享受折上折优惠</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Balance}')"><IMG class=icon src="images/label.gif">资金余额</DIV>
            <DIV class=subItem onClick="InsertLabel('{$UserPoint}')"><IMG class=icon src="images/label.gif">可用点数</DIV>
            <DIV class=subItem onClick="InsertLabel('{$UserExp}')"><IMG class=icon src="images/label.gif">可用积分</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ContacterName}')"><IMG class=icon src="images/label.gif">收货人姓名</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Address}')"><IMG class=icon src="images/label.gif">收货人地址</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ZipCode}')"><IMG class=icon src="images/label.gif">收货人邮编</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Phone}')"><IMG class=icon src="images/label.gif">收货人电话</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Email}')"><IMG class=icon src="images/label.gif">收货人邮箱</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PaymentTypeList}')"><IMG class=icon src="images/label.gif">付款方式</DIV>
            <DIV class=subItem onClick="InsertLabel('{$DeliverTypeList}')"><IMG class=icon src="images/label.gif">送货方式</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$InvoiceInfo}')"><IMG class=icon src="images/label.gif">发票信息</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Remark}')"><IMG class=icon src="images/label.gif">备注/留言</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$ShowCart}')"><IMG class=icon src="images/label.gif">显示购物车中的商品</DIV>
        </DIV>
        <!-- 收银台标签结束 -->
        <!-- 订单预览标签 -->
        <DIV class=subFolder id=subshopPreview><IMG id=subshopPreviewImg class=icon src="images/foldericon1.gif"> 订单预览标签</DIV>
        <DIV class=sub id=subshopPreviewSub>
            <DIV class=subItem onClick="InsertLabel('{$ShowTips_CartIsEmpty}')"><IMG class=icon src="images/label.gif">购物车提示</DIV>        
            <DIV class=subItem onClick="InsertLabel('{$UserName}')"><IMG class=icon src="images/label.gif">用户名</DIV>
            <DIV class=subItem onClick="InsertLabel('{$GroupName}')"><IMG class=icon src="images/label.gif">用户组</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Discount_Member}%')"><IMG class=icon src="images/label.gif">会员折扣率</DIV>
            <DIV class=subItem onClick="InsertLabel('{$IsOffer}')"><IMG class=icon src="images/label.gif">享受折上折优惠</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Balance}')"><IMG class=icon src="images/label.gif">资金余额</DIV>
            <DIV class=subItem onClick="InsertLabel('{$UserPoint}')"><IMG class=icon src="images/label.gif">可用点数</DIV>
            <DIV class=subItem onClick="InsertLabel('{$UserExp}')"><IMG class=icon src="images/label.gif">可用积分</DIV>
            <DIV class=subItem onClick="InsertLabel('{$TrueName}')"><IMG class=icon src="images/label.gif">收货人名称</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Address}')"><IMG class=icon src="images/label.gif">收货人地址</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ZipCode}')"><IMG class=icon src="images/label.gif">收货人邮编</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Phone}')"><IMG class=icon src="images/label.gif">收货人电话</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowPaymentTypeList}')"><IMG class=icon src="images/label.gif">付款方式</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Company}')"><IMG class=icon src="images/label.gif">发票信息</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowDeliverTypeList}')"><IMG class=icon src="images/label.gif">送货方式</DIV>        
            <DIV class=subItem onClick="InsertLabel('{$ShowCart}')"><IMG class=icon src="images/label.gif">显示购物车中的商品</DIV>
        </DIV>
        <!-- 订单预览标签结束 -->
        <!-- 显示订单成功标签 -->
        <DIV class=subFolder id=subshopSucceed><IMG id=subshopSucceedImg class=icon src="images/foldericon1.gif"> 显示订单成功标签</DIV>
        <DIV class=sub id=subshopSucceedSub>
            <DIV class=subItem onClick="InsertLabel('{$OrderFormNum}')"><IMG class=icon src="images/label.gif">订单号码</DIV>        
            <DIV class=subItem onClick="InsertLabel('{$TotalMoney}')"><IMG class=icon src="images/label.gif">交易金额</DIV>
        </DIV>
        <!-- 显示订单成功标签结束 -->
        <!-- 显示订单成功标签 -->
        <DIV class=subFolder id=subshopPayment><IMG id=subshopPaymentImg class=icon src="images/foldericon1.gif"> 在线支付标签</DIV>
        <DIV class=sub id=subshopPaymentSub>
            <DIV class=subItem onClick="InsertLabel('{$OrderFormNum}')"><IMG class=icon src="images/label.gif">订单号码</DIV>        
            <DIV class=subItem onClick="InsertLabel('{$MoneyTotal}')"><IMG class=icon src="images/label.gif">订单金额</DIV>
            <DIV class=subItem onClick="InsertLabel('{$MoneyReceipt}')"><IMG class=icon src="images/label.gif">已 付 款</DIV>
            <DIV class=subItem onClick="InsertLabel('{$MoneyNeedPay}')"><IMG class=icon src="images/label.gif">需要支付</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PaymentNum}')"><IMG class=icon src="images/label.gif">支付序列号</DIV>        
            <DIV class=subItem onClick="InsertLabel('￥{$vMoney}')"><IMG class=icon src="images/label.gif">支付金额</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PayOnlineRate}')"><IMG class=icon src="images/label.gif">手续费</DIV>
            <DIV class=subItem onClick="InsertLabel('￥{$v_amount}')"><IMG class=icon src="images/label.gif">实际划款金额</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PayOnlineProviderName}')"><IMG class=icon src="images/label.gif">在线支付平台提供商</DIV>        
            <DIV class=subItem onClick="InsertLabel('{$HiddenField}')"><IMG class=icon src="images/label.gif">支付隐藏字段</DIV>
        </DIV>
        <!-- 显示订单成功标签结束 -->
     </DIV>
    <% 
    End if
    If (ModuleType=7 or ModuleType=0) And FoundInArr(AllModules, "House", ",") Then %>
    <!--  房产频道函数  -->
     <DIV class=topFolder id=House><IMG id=HouseImg class=icon src="images/foldericon1.gif">房产标签</DIV>
    <DIV class=sub id=HouseSub>
         <!-- 房产通用频道标签 -->
         <DIV class=subFolder id=subHouseChannelFunction><IMG id=subHouseChannelFunctionImg class=icon src="images/foldericon1.gif"> 房产频道标签</DIV>
         <DIV class=sub id=subHouseChannelFunctionSub>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_HouseList.htm',560,360)"><IMG class=icon src="images/label2.gif">显示房产地址等信息</DIV>
        </DIV>
        <!-- 房产通用频道标签结束 -->
        <!-- 房产频道内容标签 -->
        <DIV class=subFolder id=subHouseChannelContent><IMG id=subHouseChannelContentImg class=icon src="images/foldericon1.gif"> 房产内容标签</DIV>
        <DIV class=sub id=subHouseChannelContentSub>
            <DIV class=subItem onClick="InsertLabel('{$HeZhuType}')"><IMG class=icon src="images/label.gif">合租类型</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseDiZhi}')"><IMG class=icon src="images/label.gif">房屋地址</DIV>
            <DIV class=subItem onClick="InsertLabel('{$My}')"><IMG class=icon src="images/label.gif">我的简介</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Chum}')"><IMG class=icon src="images/label.gif">室友要求</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseHuXing}')"><IMG class=icon src="images/label.gif">房屋户型</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseHuXing1}')"><IMG class=icon src="images/label.gif">分租部分</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseHuXing2}')"><IMG class=icon src="images/label.gif">共用部分</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseXingZhi}')"><IMG class=icon src="images/label.gif">房屋性质（新房、二手房等）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseChanQuan}')"><IMG class=icon src="images/label.gif">房屋产权</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseJianCheng}')"><IMG class=icon src="images/label.gif">建成日期</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseJianCheng1}')"><IMG class=icon src="images/label.gif">期望建成日期范围（始）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseJianCheng2}')"><IMG class=icon src="images/label.gif">期望建成日期范围（终）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseMianJi}')"><IMG class=icon src="images/label.gif">房屋面积</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseMianJi1}')"><IMG class=icon src="images/label.gif">期望面积范围（始）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseMianJi2}')"><IMG class=icon src="images/label.gif">期望面积范围（终）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseLouCeng}')"><IMG class=icon src="images/label.gif">楼层</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseLeiXing}')"><IMG class=icon src="images/label.gif">物业类型</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseChaoXiang}')"><IMG class=icon src="images/label.gif">房屋朝向</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseShuiDian}')"><IMG class=icon src="images/label.gif">水电设施</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseSheShi}')"><IMG class=icon src="images/label.gif">基础设施（电梯、车库、供热、水电表等）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseZhuangXiu}')"><IMG class=icon src="images/label.gif">装修程度</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseDianQi}')"><IMG class=icon src="images/label.gif">电器设备</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseWeiSheng}')"><IMG class=icon src="images/label.gif">卫生设施</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseJiaJu}')"><IMG class=icon src="images/label.gif">附带家具</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseXinXi}')"><IMG class=icon src="images/label.gif">信息设施</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseGongJia}')"><IMG class=icon src="images/label.gif">附近公交</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$HouseHuanJing}')"><IMG class=icon src="images/label.gif">配套市政</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$JiaoFangStartDate}')"><IMG class=icon src="images/label.gif">交房日期</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseQiTa}')"><IMG class=icon src="images/label.gif">其它说明（如：房屋图片）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$TotalPrice}')"><IMG class=icon src="images/label.gif">房屋价格（用于出售）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseZuJin}')"><IMG class=icon src="images/label.gif">房屋租金（用于出租）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HousePrice1}')"><IMG class=icon src="images/label.gif">期望价格范围（最低，用于求购）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HousePrice2}')"><IMG class=icon src="images/label.gif">期望价格范围（最高，用于求购）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseZuJin1}')"><IMG class=icon src="images/label.gif">期望租金范围（最低，用于求租）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseZuJin2}')"><IMG class=icon src="images/label.gif">期望租金范围（最高，用于求租）</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseZhiFu}')"><IMG class=icon src="images/label.gif">支付方式</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HousePriceType}')"><IMG class=icon src="images/label.gif">价格单位</DIV>
            <DIV class=subItem onClick="InsertLabel('{$HouseZuJinType}')"><IMG class=icon src="images/label.gif">租金单位</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ZuLinStartDate}')"><IMG class=icon src="images/label.gif">租赁时间范围（始）</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$ZuLinEndDate}')"><IMG class=icon src="images/label.gif">租赁时间范围（末）</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$JiaoFangStartDate}')"><IMG class=icon src="images/label.gif">交房日期</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ContactPhone}')"><IMG class=icon src="images/label.gif">联系电话</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ContactName}')"><IMG class=icon src="images/label.gif">联系人</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ContactEmail}')"><IMG class=icon src="images/label.gif">电子邮箱</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ContactQQ}')"><IMG class=icon src="images/label.gif">联系ＱＱ</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Editor}')"><IMG class=icon src="images/label.gif">房产信息作者</DIV>
            <DIV class=subItem onClick="InsertLabel('{$UpdateTime}')"><IMG class=icon src="images/label.gif">发布日期</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Hits}')"><IMG class=icon src="images/label.gif">点击数</DIV>
        </DIV>
        <!-- 房产频道内容标签结束 -->
    </DIV>
    <%
    End if
    If (ModuleType=8 or ModuleType=0) And FoundInArr(AllModules, "Job", ",") Then %>
    <!--  人才招聘频道函数  -->
     <DIV class=topFolder id=Job><IMG id=JobImg class=icon src="images/foldericon1.gif">人才招聘标签</DIV>
    <DIV class=sub id=JobSub>
        <!-- 人才招聘通用频道标签 -->
        <DIV class=subFolder id=subJobChannelFunction><IMG id=subJobChannelFunctionImg class=icon src="images/foldericon1.gif"> 人才招聘频道首页（列表）标签</DIV>
        <DIV class=sub id=subJobChannelFunctionSub>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_label.asp','GetPositionList','职位列表函数标签',8,'GetPositionList',650,500)"><IMG class=icon src="images/label3.gif">显示（所有、最新、紧急）职位名称等信息</DIV>
        </DIV>
        <DIV class=subFolder id=subJobChannelFunction2><IMG id=subJobChannelFunction2Img class=icon src="images/foldericon1.gif"> 人才招聘频道首页（内容）标签</DIV>
        <DIV class=sub id=subJobChannelFunction2Sub>
            <DIV class=subItem onClick="FunctionLabel2('【PositionList_Content】')"><IMG class=icon src="images/label2.gif">循环显示职位内容信息</DIV>
         </DIV>
        <DIV class=subFolder id=subJobChannelFunction3><IMG id=subJobChannelFunction3Img class=icon src="images/foldericon1.gif"> 人才招聘频道搜索结果页标签</DIV>
        <DIV class=sub id=subJobChannelFunction3Sub>
            <DIV class=subItem onClick="SuperFunctionLabel('editor_label.asp','GetSearchResult','职位搜索结果列表函数标签',8,'GetSearchResult',590,450)"><IMG class=icon src="images/label3.gif">显示（搜索结果）职位名称等信息</DIV>
        </DIV>
        <!-- 人才招聘通用频道标签结束 -->
        <!-- 人才招聘频道内容标签 -->
        <DIV class=subFolder id=subJobChannelContent><IMG id=subJobChannelContentImg class=icon src="images/foldericon1.gif"> 人才招聘内容标签</DIV>
        <DIV class=sub id=subJobChannelContentSub>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_CorrelativePosition.htm',560,360)"><IMG class=icon src="images/label2.gif">显示相关职位名称等信息</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PositionName}')"><IMG class=icon src="images/label.gif">职位名称</DIV>
            <DIV class=subItem onClick="InsertLabel('{$WorkPlaceName}')"><IMG class=icon src="images/label.gif">工作地点</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PositionNum}')"><IMG class=icon src="images/label.gif">招聘人数</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ReleaseDate}')"><IMG class=icon src="images/label.gif">发布日期</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ValidDate}')"><IMG class=icon src="images/label.gif">有效期</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SubCompanyName}')"><IMG class=icon src="images/label.gif">用人单位</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Contacter}')"><IMG class=icon src="images/label.gif">联系人</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Telephone}')"><IMG class=icon src="images/label.gif">联系电话</DIV>
            <DIV class=subItem onClick="InsertLabel('{$Address}')"><IMG class=icon src="images/label.gif">联系地址</DIV>
            <DIV class=subItem onClick="InsertLabel('{$E_mail}')"><IMG class=icon src="images/label.gif">联系E_mail</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PositionDescription')"><IMG class=icon src="images/label.gif">职位描述</DIV>
            <DIV class=subItem onClick="InsertLabel('{$DutyRequest}')"><IMG class=icon src="images/label.gif">任职要求</DIV>
            <DIV class=subItem onClick="InsertLabel('{$PositionStatus}')"><IMG class=icon src="images/label.gif">职位状态</DIV>
            <DIV class=subItem onClick="InsertLabel('{$SaveSupply}')"><IMG class=icon src="images/label.gif">申请职位按钮</DIV>
        </DIV>
        <!-- 人才招聘频道内容标签结束 -->
    </DIV>
    <%End if
    If (ModuleType=6 or ModuleType=0) And FoundInArr(AllModules, "Supply", ",") then%>
        <DIV class=subFolder id=subsupplyInfo><IMG id=subsupplyInfoImg class=icon src="images/foldericon1.gif">供求信息页标签</DIV>
        <DIV class=sub id=subsupplyInfoSub>
        <DIV class=subItem onClick="FunctionLabel('Lable/PE_SupplyInfo.htm','600','700')"><IMG class=icon src="images/label2.gif">供求信息列表标签</DIV>
        <DIV class=subItem onClick="FunctionLabel('Lable/PE_SupplyLasterInfo.htm','600','350')"><IMG class=icon src="images/label2.gif">供求最新信息列表标签</DIV>
        <DIV class=subFolder id=subsupplyInfoContent><IMG id=subsupplyInfoContentImg class=icon src="images/foldericon1.gif"> 供求信息内容标签</DIV>
        <DIV class=sub id=subsupplyInfoContentSub>
            <DIV class=subItem onClick="InsertLabel('{$SupplyInfoTitle}')"><IMG class=icon src="images/label.gif">信息标题</DIV> 
            <DIV class=subItem onClick="InsertLabel('{$SupplyInfoType}')"><IMG class=icon src="images/label.gif">信息类型</DIV> 
            <DIV class=subItem onClick="InsertLabel('{$TradeType}')"><IMG class=icon src="images/label.gif">交易方式</DIV> 
            <DIV class=subItem onClick="InsertLabel('{$UserName}')"><IMG class=icon src="images/label.gif">发 布 人</DIV> 
            <DIV class=subItem onClick="InsertLabel('{$UpdateTime}')"><IMG class=icon src="images/label.gif">发布日期</DIV> 
            <DIV class=subItem onClick="InsertLabel('{$EndTime}')"><IMG class=icon src="images/label.gif">有效期至</DIV> 
            <DIV class=subItem onClick="InsertLabel('{$SupplyIntro}')"><IMG class=icon src="images/label.gif">详细内容</DIV> 
            <DIV class=subItem onClick="InsertLabel('{$Province}')"><IMG class=icon src="images/label.gif">发布人所属的省</DIV> 
            <DIV class=subItem onClick="InsertLabel('{$City}')"><IMG class=icon src="images/label.gif">发布人所属的市</DIV> 
            <DIV class=subItem onClick="InsertLabel('{$Address}')"><IMG class=icon src="images/label.gif">联系地址</DIV> 
            <DIV class=subItem onClick="InsertLabel('{$ZipCode}')"><IMG class=icon src="images/label.gif">邮　　编</DIV> 
            <DIV class=subItem onClick="InsertLabel('{$Email}')"><IMG class=icon src="images/label.gif">电子邮件</DIV> 
            <DIV class=subItem onClick="InsertLabel('{$CompanyName}')"><IMG class=icon src="images/label.gif">公 司 名</DIV> 
            <DIV class=subItem onClick="InsertLabel('{$Department}')"><IMG class=icon src="images/label.gif">所属部门</DIV> 
            <DIV class=subItem onClick="InsertLabel('{$CompanyAddress}')"><IMG class=icon src="images/label.gif">公司地址</DIV> 
            <DIV class=subItem onClick="InsertLabel('{$SupplyAction}')"><IMG class=icon src="images/label.gif">显示【发表评论】【告诉好友】【打印此文】【关闭窗口】</DIV>
        </DIV>
        <DIV class=subItem onClick="FunctionLabel('Lable/PE_SupplySearchInfo.htm','500','250')"><IMG class=icon src="images/label2.gif">供求信息搜索条件标签</DIV>    
        <DIV class=subItem onClick="InsertLabel('{$SearchResul}')"><IMG class=icon src="images/label.gif">显示搜索结果页标签</DIV> 
        </DIV>
    </DIV>
    <%End if%>
     <!--  作者,来源,厂商,品牌,标签  -->
     <DIV class=topFolder id=Aomb><IMG id=AombImg class=icon src="images/foldericon1.gif">作者,来源,厂商,品牌</DIV>
     <DIV class=sub id=AombSub>
         <!-- 作者 标签 -->
         <DIV class=subFolder id=Author><IMG id=AuthorImg class=icon src="images/foldericon1.gif">作者标签</DIV>
         <DIV class=sub id=AuthorSub>
            <DIV class=subItem onClick="InsertLabel('{$AuthorName}')"><IMG class=icon src="images/label.gif">作者姓名</DIV>    
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_Author_Photo.htm','240','150')"><IMG class=icon src="images/label2.gif">作者照片</DIV>    
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_Author_List.htm','240','230')"><IMG class=icon src="images/label2.gif">显示作者列表</DIV>        
            <DIV class=subItem onClick="InsertLabel('{$AuthorSex}')"><IMG class=icon src="images/label.gif">作者性别</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$AuthorAddTime}')"><IMG class=icon src="images/label.gif">文集批准时间</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$AuthorBirthDay}')"><IMG class=icon src="images/label.gif">作者生日</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$AuthorCompany}')"><IMG class=icon src="images/label.gif">作者公司</DIV>
            <DIV class=subItem onClick="InsertLabel('{$AuthorDepartment}')"><IMG class=icon src="images/label.gif">作者部门</DIV>
            <DIV class=subItem onClick="InsertLabel('{$AuthorAddress}')"><IMG class=icon src="images/label.gif">作者地址</DIV>
            <DIV class=subItem onClick="InsertLabel('{$AuthorTel}')"><IMG class=icon src="images/label.gif">作者电话</DIV>
            <DIV class=subItem onClick="InsertLabel('{$AuthorFax}')"><IMG class=icon src="images/label.gif">作者传真</DIV>
            <DIV class=subItem onClick="InsertLabel('{$AuthorZipCode}')"><IMG class=icon src="images/label.gif">作者邮编</DIV>
            <DIV class=subItem onClick="InsertLabel('{$AuthorHomePage}')"><IMG class=icon src="images/label.gif">作者主页</DIV>
            <DIV class=subItem onClick="InsertLabel('{$AuthorEmail}')"><IMG class=icon src="images/label.gif">作者邮件</DIV>
            <DIV class=subItem onClick="InsertLabel('{$AuthorQQ}')"><IMG class=icon src="images/label.gif">作者QQ</DIV>
            <DIV class=subItem onClick="InsertLabel('{$AuthorType}')"><IMG class=icon src="images/label.gif">作者分类</DIV>
            <DIV class=subItem onClick="InsertLabel('{$AuthorIntro}')"><IMG class=icon src="images/label.gif">作者说明</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_Author_ArtList.htm','350','330')"><IMG class=icon src="images/label2.gif">作者文章列表</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_Author_ShowList.htm','400','345')"><IMG class=icon src="images/label2.gif">显示作者列表</DIV>
         </DIV>
         <!-- 来源 标签 -->
         <DIV class=subFolder id=origin><IMG id=originImg class=icon src="images/foldericon1.gif">来源标签</DIV>
         <DIV class=sub id=originSub>
            <DIV class=subItem onClick="InsertLabel('{$ShowPhoto}')"><IMG class=icon src="images/label.gif">来源图片</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowName}')"><IMG class=icon src="images/label.gif">来源姓名</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowContacterName}')"><IMG class=icon src="images/label.gif">联系人</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowAddress}')"><IMG class=icon src="images/label.gif">地址</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowTel}')"><IMG class=icon src="images/label.gif">电话</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowFax}')"><IMG class=icon src="images/label.gif">传真</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowZipCode}')"><IMG class=icon src="images/label.gif">邮编</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowMail}')"><IMG class=icon src="images/label.gif">信箱</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowHomePage}')"><IMG class=icon src="images/label.gif">主页</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowEmail}')"><IMG class=icon src="images/label.gif">邮件</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowQQ}')"><IMG class=icon src="images/label.gif">QQ</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowType}')"><IMG class=icon src="images/label.gif">分类</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowMemo}')"><IMG class=icon src="images/label.gif">简介</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowArticleList}')"><IMG class=icon src="images/label.gif">显示文章列表</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowCopyFromList}')"><IMG class=icon src="images/label.gif">来源列表</DIV>  
         </DIV>
         <!-- 厂商标签 -->
         <DIV class=subFolder id=manufacturer><IMG id=manufacturerImg class=icon src="images/foldericon1.gif">厂商标签</DIV>
         <DIV class=sub id=manufacturerSub>
            <DIV class=subItem onClick="InsertLabel('{$ShowPhoto}')"><IMG class=icon src="images/label.gif">厂商图片</DIV>        
            <DIV class=subItem onClick="InsertLabel('{$ShowName}')"><IMG class=icon src="images/label.gif">姓名</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowProducerShortName}')"><IMG class=icon src="images/label.gif">缩写</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowBirthDay}')"><IMG class=icon src="images/label.gif">建立日期</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowAddress}')"><IMG class=icon src="images/label.gif">地址</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$ShowTel}')"><IMG class=icon src="images/label.gif">电话</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$ShowFax}')"><IMG class=icon src="images/label.gif">传真</DIV>    
            <DIV class=subItem onClick="InsertLabel('{$ShowZipCode}')"><IMG class=icon src="images/label.gif">邮编</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowHomePage}')"><IMG class=icon src="images/label.gif">主页</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowEmail}')"><IMG class=icon src="images/label.gif">邮件</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowType}')"><IMG class=icon src="images/label.gif">分类</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowtrademarkList}')"><IMG class=icon src="images/label.gif">持有品牌</DIV>        
            <DIV class=subItem onClick="InsertLabel('{$ShowMemo}')"><IMG class=icon src="images/label.gif">简介</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_Product_List.htm','400','230')"><IMG class=icon src="images/label2.gif">显示商品列表</DIV>
         </DIV>
         <!-- 品牌标签 -->
         <DIV class=subFolder id=brand><IMG id=brandImg class=icon src="images/foldericon1.gif">品牌标签</DIV>
         <DIV class=sub id=brandSub>
            <DIV class=subItem onClick="InsertLabel('{$ShowPhoto}')"><IMG class=icon src="images/label.gif">品牌图片</DIV>        
            <DIV class=subItem onClick="InsertLabel('{$ShowName}')"><IMG class=icon src="images/label.gif">姓名</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowType}')"><IMG class=icon src="images/label.gif">分类</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowProducerName}')"><IMG class=icon src="images/label.gif">所属厂商</DIV>        
            <DIV class=subItem onClick="InsertLabel('{$ShowMemo}')"><IMG class=icon src="images/label.gif">简介</DIV>
            <DIV class=subItem onClick="FunctionLabel('Lable/PE_Product_List.htm','400','230')"><IMG class=icon src="images/label2.gif">显示商品列表</DIV>
            <DIV class=subItem onClick="InsertLabel('{$ShowtrademarkList}')"><IMG class=icon src="images/label.gif">显示品牌列表</DIV>  
         </DIV>
     </DIV>
     <!-- Rss标签 -->
     <DIV class=topFolder id=RssItem><IMG id=RssItemImg class=icon src="images/foldericon1.gif">RSS</DIV>
     <DIV class=sub id=RssItemSub>
        <DIV class=subItem onClick="InsertLabel('{$Rss}')"><IMG class=icon src="images/label.gif">RSS标签显示</DIV>    
        <DIV class=subItem onClick="InsertLabel('{$RssElite}')"><IMG class=icon src="images/label.gif">RSS推荐标签显示</DIV>    
        <DIV class=subItem onClick="InsertLabel('{$RssHot}')"><IMG class=icon src="images/label.gif">RSS热点文章标签显示</DIV>
     </DIV>
     <DIV class=topFolder id=AnnounceItem><IMG id=AnnounceItemImg class=icon src="images/foldericon1.gif">公告标签</DIV>
     <DIV class=sub id=AnnounceItemSub>
        <DIV class=subItem onClick="InsertLabel('{$AnnounceList}')"><IMG class=icon src="images/label.gif">公告列表</DIV>     
     </DIV>
     <DIV class=topFolder id=FriendItem><IMG id=FriendItemImg class=icon src="images/foldericon1.gif">友情链接标签</DIV>
     <DIV class=sub id=FriendItemSub>
        <DIV class=subItem onClick="InsertLabel('{$FriendSiteList}')"><IMG class=icon src="images/label.gif">友情链接列表</DIV>
     </DIV>
     <DIV class=topFolder id=VoteItem><IMG id=VoteItemImg class=icon src="images/foldericon1.gif">调查标签</DIV>
     <DIV class=sub id=VoteItemSub>
         <DIV class=subItem onClick="InsertLabel('[VoteItem] 请在这里输入要循环调查的标签[/VoteItem] ')"><IMG class=icon src="images/label.gif">循环显示调查项目</DIV>
        <DIV class=subItem onClick="InsertLabel('{$VoteTitle}')"><IMG class=icon src="images/label.gif">显示调查标题</DIV>
        <DIV class=subItem onClick="InsertLabel('{$TotalVote}')"><IMG class=icon src="images/label.gif">调查投票总数</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ItemNum}')"><IMG class=icon src="images/label.gif">调查选项数字</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ItemSelect}')"><IMG class=icon src="images/label.gif">调查选项名称</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ItemPer}')"><IMG class=icon src="images/label.gif">调查选项所占百分比</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ItemAnswer}')"><IMG class=icon src="images/label.gif">调查选项所得票数</DIV>
        <DIV class=subItem onClick="InsertLabel('{$VoteForm}')"><IMG class=icon src="images/label.gif">调查选项内容</DIV>        
        <DIV class=subItem onClick="InsertLabel('{$OtherVote}')"><IMG class=icon src="images/label.gif">查看其它调查项目</DIV>
     </DIV>
     <!-- Wap标签 -->
     <DIV class=topFolder id=WapItem><IMG id=WapItemImg class=icon src="images/foldericon1.gif">Wap标签</DIV>
     <DIV class=sub id=WapItemSub>    
        <DIV class=subItem onClick="InsertLabel('{$Wap}')"><IMG class=icon src="images/label.gif">WAP标签显示</DIV>    
     </DIV>
     <!-- 会员标签 -->
     <DIV class=topFolder id=associatorItem><IMG id=associatorItemImg class=icon src="images/foldericon1.gif">会员管理标签</DIV>
     <DIV class=sub id=associatorItemSub>
        <DIV class=subItem onClick="InsertLabel('{$UserFace}')"><IMG class=icon src="images/label.gif">会员头像</DIV>    
        <DIV class=subItem onClick="InsertLabel('{$TrueName}')"><IMG class=icon src="images/label.gif">姓名</DIV>
        <DIV class=subItem onClick="InsertLabel('{$Sex}')"><IMG class=icon src="images/label.gif">性别</DIV>
        <DIV class=subItem onClick="InsertLabel('{$BirthDay}')"><IMG class=icon src="images/label.gif">诞辰</DIV>
        <DIV class=subItem onClick="InsertLabel('{$Company}')"><IMG class=icon src="images/label.gif">公司</DIV>
        <DIV class=subItem onClick="InsertLabel('{$Department}')"><IMG class=icon src="images/label.gif">部门</DIV>
        <DIV class=subItem onClick="InsertLabel('{$Address}')"><IMG class=icon src="images/label.gif">地址</DIV>
        <DIV class=subItem onClick="InsertLabel('{$HomePhone}')"><IMG class=icon src="images/label.gif">电话</DIV>
        <DIV class=subItem onClick="InsertLabel('{$Fax}')"><IMG class=icon src="images/label.gif">传真</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ZipCode}')"><IMG class=icon src="images/label.gif">邮编</DIV>
        <DIV class=subItem onClick="InsertLabel('{$HomePage}')"><IMG class=icon src="images/label.gif">主页</DIV>
        <DIV class=subItem onClick="InsertLabel('{$Email}')"><IMG class=icon src="images/label.gif">邮件</DIV>
        <DIV class=subItem onClick="InsertLabel('{$QQ}')"><IMG class=icon src="images/label.gif">QQ</DIV>
        <DIV class=subItem onClick="InsertLabel('{$ShowUserList}')"><IMG class=icon src="images/label.gif">会员列表</DIV>
     </DIV>
     <!--  自定义标签  -->
     <%
        Set rs=Server.CreateObject("ADODB.Recordset")
        sql="select LabelID,LabelName,LabelClass,LabelType,fieldlist from PE_Label Where LabelType=0 Order by LabelClass,LabelID desc"
        rs.open sql,conn,1,1
        If not(rs.bof and rs.EOF) Then
            response.Write("<DIV class=topFolder id=Label><IMG id=LabelImg class=icon src=""images/foldericon1.gif"">自定义静态标签</DIV>")
            response.Write("<DIV class=sub id=LabelSub>")
            Do while not rs.eof
                Response.write "<DIV class=subItem onclick=""InsertLabel('{$" & rs("LabelName")& "}')""><IMG class=icon src=""images/label.gif"">"
                If Trim(rs("LabelClass") & "") <> "" Then Response.write "<font color=#999999>[" & rs("LabelClass") & "]</font>"
                Response.write rs("LabelName") &"</DIV>"
                rs.movenext
            loop
            response.Write("</DIV>")
        End If
        rs.close
        sql="select LabelID,LabelName,LabelClass,LabelType,fieldlist from PE_Label Where LabelType=1 Order by LabelClass,LabelID desc"
        rs.open sql,conn,1,1
        If not(rs.bof and rs.EOF) Then
            response.Write("<DIV class=topFolder id=Label1><IMG id=Label1Img class=icon src=""images/foldericon1.gif"">自定义动态标签</DIV>")
            response.Write("<DIV class=sub id=Label1Sub>")
            Do while not rs.eof
                Response.write "<DIV class=subItem onclick=""InsertLabel('{$" & rs("LabelName")& "}')""><IMG class=icon src=""images/label.gif"">"
                If Trim(rs("LabelClass") & "") <> "" Then Response.write "<font color=#999999>[" & rs("LabelClass") & "]</font>"
                Response.write rs("LabelName") &"</DIV>"
                rs.movenext
            loop
            response.Write("</DIV>")
        End If
        rs.close
        sql="select LabelID,LabelName,LabelClass,LabelType,fieldlist from PE_Label Where LabelType=2 Order by LabelClass,LabelID desc"
        rs.open sql,conn,1,1
        If not(rs.bof and rs.EOF) Then
            response.Write("<DIV class=topFolder id=Label2><IMG id=Label2Img class=icon src=""images/foldericon1.gif"">自定义采集标签</DIV>")
            response.Write("<DIV class=sub id=Label2Sub>")
            Do while not rs.eof
                Response.write "<DIV class=subItem onclick=""InsertLabel('{$" & rs("LabelName")& "}')""><IMG class=icon src=""images/label.gif"">"
                If Trim(rs("LabelClass") & "") <> "" Then Response.write "<font color=#999999>[" & rs("LabelClass") & "]</font>"
                Response.write rs("LabelName") &"</DIV>"
                rs.movenext
            loop
            response.Write("</DIV>")
        End If
        rs.close
        sql="select LabelID,LabelName,LabelClass,LabelType,fieldlist from PE_Label Where LabelType=3 Order by LabelClass,LabelID desc"
        rs.open sql,conn,1,1
        If not(rs.bof and rs.EOF) Then
            response.Write("<DIV class=topFolder id=Label3><IMG id=Label3Img class=icon src=""images/foldericon1.gif"">自定义函数标签</DIV>")
            response.Write("<DIV class=sub id=Label3Sub>")
            Do while not rs.eof
                Response.write "<DIV class=subItem onclick=""FunctionLabel('editor_listdynafield.asp?id=" & rs("LabelID")& "','400','480')""><IMG class=icon src=""images/label.gif"">"
                If Trim(rs("LabelClass") & "") <> "" Then Response.write "<font color=#999999>[" & rs("LabelClass") & "]</font>"
                Response.write rs("LabelName") &"</DIV>"
                rs.movenext
            loop
            response.Write("</DIV>")
        End If
        rs.close
        set rs=nothing
      %>
     <!--  自定义字段标签  -->
     <DIV class=topFolder id=Field><IMG id=FieldImg class=icon src="images/foldericon1.gif">自定义字段标签</DIV>
     <DIV class=sub id=FieldSub>
     <%
        sql="select  LabelName,FieldName from PE_Field where ChannelID=" & PE_Clng(ChannelID) & " Order by FieldID desc"
        Set rs=Server.CreateObject("ADODB.Recordset")
            rs.open sql,conn,1,1
            if rs.bof and  rs.eof then
                response.Write("<li>您还没有自定义字段标签,或自定义字段标签只显示在所属频道</li>")
            else
                Do while not rs.eof
                    Response.write"<DIV class=subItem onclick=""InsertLabel('" & rs("LabelName")& "')""><IMG class=icon src=""images/label.gif"">"& rs("FieldName") &"</DIV>"
                    rs.movenext
                loop
            end if 
            rs.close
        set rs=nothing
      %>
    </DIV>
    <!--  广告位标签开始  -->
    <DIV class=topFolder id=AdJs><IMG id=AdJsImg class=icon src="images/foldericon1.gif">广告版位标签</DIV>
    <DIV class=sub id=AdJsSub>
    <%

        Function GetZoneJSName(iZoneID, UpdateTime)
            Set XmlDoc = CreateObject("Microsoft.XMLDOM")
            XmlDoc.async = False
            XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))
            GetZoneJSName = InstallDir & ADDir & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/" & iZoneID & ".js"
            Set XmlDoc = Nothing
        End Function
        sql="select ZoneID,ZoneName,UpdateTime from PE_AdZone"
        Set rs=Server.CreateObject("ADODB.Recordset")
        rs.open sql,conn,1,1
        if rs.bof and  rs.eof then
           response.Write("<li>您还没有定义广告版位 </li>")
        else
            Do while not rs.eof             
                Response.write"<DIV class=subItem onclick=""InsertAdjs('Adjs','" & GetZoneJSName(rs("ZoneID"), rs("UpdateTime")) & "')""><IMG class=icon src=""images/jscript.gif"">"& rs("ZoneName") &"</DIV>"
                rs.movenext
            loop
        end if 
        rs.close
        set rs=nothing
        conn.close
        set conn=nothing
    %>
    </DIV>
    <!--  广告位标签结束  -->
     <!-- 其它JS标签  -->
     <DIV class=topFolder id=OtherJS><IMG id=OtherJSImg class=icon src="images/foldericon1.gif">其它JS标签</DIV>
     <DIV class=sub id=OtherJSSub>
     <DIV class=subItem onClick="InsertAdjs('SwitchFont','{$InstallDir}js/gb_big5.js')"><IMG class=icon src="images/jscript.gif">切换到繁體中文</DIV>
     <DIV class=subItem onClick="FunctionLabel2('ResumeError')"><IMG class=icon src="images/jscript.gif">屏蔽页面JS错误</DIV>
</DIV>
  </td>
 </tr>
   </td>
  </tr>
</table>
<!-- ******** 菜单效果结束 ******** -->
    <!-- 显示说明 -->
    <table width='100%' height='60' border='0' align='center' cellpadding='0' cellspacing='0' bgcolor="#EEF4FF" style='border: 1px solid #0066FF;'>
      <tr align="center">
        <td height="22" colspan="2" bgcolor='#0066FF'><font color="#FFFFFF">==&gt;&nbsp;显示说明&nbsp;&lt;==</font></td>
      </tr>
      <tr>
        <td width="9%" rowspan="3">&nbsp;</td>
        <td width="91%"><IMG class=icon src="images/label.gif"> >>>  普通标签 </td>
      </tr>
      <tr>
        <td><IMG class=icon src="images/label2.gif"> >>> 函数标签 </td>
      </tr>
      <tr>
        <td><IMG class=icon src="images/label3.gif"> >>> 超级函数标签 </td>
      </tr>
    </table>
    <!-- 显示结束 -->
</body>
</html>
