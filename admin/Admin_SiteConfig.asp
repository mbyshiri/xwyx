<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Edition.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 1      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限


Response.Write "<html><head><title>网站配置</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "<script language='JavaScript'>" & vbCrLf
Response.Write "function SelectColor(sEL,form){" & vbCrLf
Response.Write "    var dEL = document.all(sEL);" & vbCrLf
Response.Write "    var url = '../Editor/editor_selcolor.asp?color='+encodeURIComponent(sEL);" & vbCrLf
Response.Write "    var arr = showModalDialog(url,window,'dialogWidth:280px;dialogHeight:250px;help:no;scroll:no;status:no');" & vbCrLf
Response.Write "    if (arr) {" & vbCrLf
Response.Write "        form.value=arr;" & vbCrLf
Response.Write "        sEL.style.backgroundColor=arr;" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "</script>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' Class='border'>" & vbCrLf
Call ShowPageTitle("网 站 信 息 配 置", 10001)
Response.Write "</table>" & vbCrLf

If Action = "SaveConfig" Then
    Call SaveConfig
    Call WriteEntry(1, AdminName, "修改网站信息配置")
Else
    Call ModifyConfig
End If
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Response.Write "</body></html>"
Call CloseConn




Sub ModifyConfig()
    Dim sqlConfig, rsConfig
    
    sqlConfig = "select * from PE_Config"
    Set rsConfig = Server.CreateObject("ADODB.Recordset")
    rsConfig.Open sqlConfig, Conn, 1, 3
    If rsConfig.BOF And rsConfig.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>网站配置数据丢失，请使用初始数据库进行恢复。</li>"
        rsConfig.Close
        Set rsConfig = Nothing
        Exit Sub
    End If
    
    Dim RegFields_MustFill, Modules
    RegFields_MustFill = rsConfig("RegFields_MustFill")
    Modules = rsConfig("Modules")

    Response.Write "<script language='javascript'>" & vbCrLf
    Response.Write "var tID=0;" & vbCrLf
    Response.Write "function ShowTabs(ID){" & vbCrLf
    Response.Write "  if(ID!=tID){" & vbCrLf
    Response.Write "    TabTitle[tID].className='title5';" & vbCrLf
    Response.Write "    TabTitle[ID].className='title6';" & vbCrLf
    Response.Write "    Tabs[tID].style.display='none';" & vbCrLf
    Response.Write "    Tabs[ID].style.display='';" & vbCrLf
    Response.Write "    tID=ID;" & vbCrLf
    Response.Write "  }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function IsDigit()" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  return ((event.keyCode >= 48) && (event.keyCode <= 57));" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<form name='myform' id='myform' method='POST' action='Admin_SiteConfig.asp' >" & vbCrLf
    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'><tr align='center' height='24'>"
    Response.Write "<td id='TabTitle' class='title6' onclick='ShowTabs(0)'>网站信息</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(1)'>网站选项</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(2)'>会员选项</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(3)'>邮件选项</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(4)'>缩略图选项</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(5)'>搜索选项</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(6)'>商城选项</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(7)'>留言本选项</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(8)'>Rss/WAP设置</td>" & vbCrLf
    Response.Write "<td id='TabTitle' class='title5' onclick='ShowTabs(9)'>手机短信设置</td>" & vbCrLf
    Response.Write "<td>&nbsp;</td></tr></table>"

    
    Response.Write "<table width='100%' border='0' cellpadding='5' cellspacing='1' Class='border'><tr><td class='tdbg'>" & vbCrLf
    Response.Write "<table width='95%' border='0' align='center' cellpadding='2' cellspacing='1' bgcolor='#FFFFFF'>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>网站名称：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='SiteName' type='text' id='SiteName' value='" & rsConfig("SiteName") & "' size='40' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>网站标题：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='SiteTitle' type='text' id='SiteTitle' value='" & rsConfig("SiteTitle") & "' size='40' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>网站地址：</strong><br>请填写完整URL地址</td>" & vbCrLf
    Response.Write "      <td><input name='SiteUrl' type='text' id='SiteUrl' value='" & rsConfig("SiteUrl") & "' size='40' maxlength='255'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><font color=red><strong>安装目录：</strong><br>系统安装目录（相对于根目录的位置）<br>系统会自动获得正确的路径，但需要手工保存设置。</font></td>" & vbCrLf
    Response.Write "      <td><input name='InstallDir' type='text' id='InstallDir' value='" & InstallDir & "' size='40' maxlength='30' readonly></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>LOGO地址：</strong><br>请填写完整URL地址</td>" & vbCrLf
    Response.Write "      <td><input name='LogoUrl' type='text' id='LogoUrl' value='" & rsConfig("LogoUrl") & "' size='40' maxlength='255'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>Banner地址：</strong><br>请填写完整URL地址</td>" & vbCrLf
    Response.Write "      <td><input name='BannerUrl' type='text' id='BannerUrl' value='" & rsConfig("BannerUrl") & "' size='40' maxlength='255'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>站长姓名：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='WebmasterName' type='text' id='WebmasterName' value='" & rsConfig("WebmasterName") & "' size='40' maxlength='20'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>站长信箱：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='WebmasterEmail' type='text' id='WebmasterEmail' value='" & rsConfig("WebmasterEmail") & "' size='40' maxlength='100'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>版权信息：</strong><br>支持HTML标记，不能使用双引号</td>" & vbCrLf
    Response.Write "      <td><textarea name='Copyright' cols='60' rows='4' id='Copyright'>" & rsConfig("Copyright") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>网站META关键词：</strong><br>针对搜索引擎设置的关键词<br>多个关键词请用,号分隔</td>" & vbCrLf
    Response.Write "      <td><textarea name='Meta_Keywords' cols='60' rows='4' id='Meta_Keywords'>" & rsConfig("Meta_Keywords") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>网站META网页描述：</strong><br>针对搜索引擎设置的网页描述<br>多个描述请用,号分隔</td>" & vbCrLf
    Response.Write "      <td><textarea name='Meta_Description' cols='60' rows='4' id='Meta_Description'>" & rsConfig("Meta_Description") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否显示网站频道：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowSiteChannel' value='1' " & IsRadioChecked(rsConfig("ShowSiteChannel"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='ShowSiteChannel' value='0' " & IsRadioChecked(rsConfig("ShowSiteChannel"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否显示管理登录链接：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowAdminLogin' value='1' " & IsRadioChecked(rsConfig("ShowAdminLogin"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='ShowAdminLogin' value='0' " & IsRadioChecked(rsConfig("ShowAdminLogin"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否保存远程图片到本地：</strong><br>如果从其它网站上复制的内容中包含图片，则将图片复制到本站服务器上</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableSaveRemote' value='1' " & IsRadioChecked(rsConfig("EnableSaveRemote"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableSaveRemote' value='0' " & IsRadioChecked(rsConfig("EnableSaveRemote"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否开放友情链接申请：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableLinkReg' value='1' " & IsRadioChecked(rsConfig("EnableLinkReg"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableLinkReg' value='0' " & IsRadioChecked(rsConfig("EnableLinkReg"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否统计友情链接点击数：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableCountFriendSiteHits' value='1' " & IsRadioChecked(rsConfig("EnableCountFriendSiteHits"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableCountFriendSiteHits' value='0' " & IsRadioChecked(rsConfig("EnableCountFriendSiteHits"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否使用软键盘输入密码：</strong><br>若选择是，则会员登录后台时使用软键盘输入密码，适合网吧等场所上网使用。</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableSoftKey' value='1' " & IsRadioChecked(rsConfig("EnableSoftKey"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableSoftKey' value='0' " & IsRadioChecked(rsConfig("EnableSoftKey"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否使用频道、栏目、专题自设内容：</strong><br>若选择是，频道、栏目、专题管理会增加自设内容选项。</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='IsCustom_Content' value='1' " & IsRadioChecked(rsConfig("IsCustom_Content"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='IsCustom_Content' value='0' " & IsRadioChecked(rsConfig("IsCustom_Content"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
	
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否启用网站匿名投稿功能：</strong><br>若选择是，网站会启用匿名投稿用户组，匿名投稿模板，前台匿名投稿功能。</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowAnonymous' value='1' " & IsRadioChecked(PE_CBool(rsConfig("ShowAnonymous")), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='ShowAnonymous' value='0' " & IsRadioChecked(PE_CBool(rsConfig("ShowAnonymous")), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf	
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>FSO(FileSystemObject)组件的名称：</strong><br>某些网站为了安全，将FSO组件的名称进行更改以达到禁用FSO的目的。如果你的网站是这样做的，请在此输入更改过的名称。</td>" & vbCrLf
    Response.Write "      <td><input name='objName_FSO' type='text' id='objName_FSO' value='" & rsConfig("objName_FSO") & "' size='40' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>后台管理目录：</strong><br>为了安全，您可以修改后台管理目录（默认为Admin），改过以后，需要再设置此处</td>" & vbCrLf
    Response.Write "      <td><input name='AdminDir' type='text' id='AdminDir' value='" & rsConfig("AdminDir") & "' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>网站广告目录：</strong><br>为了不让广告拦截软件拦截网站的广告，您可以修改广告JS的存放目录（默认为AD），改过以后，需要再设置此处</td>" & vbCrLf
    Response.Write "      <td><input name='ADDir' type='text' id='ADDir' value='" & rsConfig("ADDir") & "' size='20' maxlength='20'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>弹出公告窗口的间隔时间：</strong><br>以小时为单位，为0时每次刷新页面时都弹出公告。</td>" & vbCrLf
    Response.Write "      <td><input name='AnnounceCookieTime' type='text' id='AnnounceCookieTime' value='" & rsConfig("AnnounceCookieTime") & "' size='10' maxlength='10'> 小时</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>网站热点的点击数最小值：</strong><br>只有点击数达到此数值，才会作为网站的热点内容显示。</td>" & vbCrLf
    Response.Write "      <td><input name='HitsOfHot' type='text' id='HitsOfHot' value='" & rsConfig("HitsOfHot") & "' size='10' maxlength='10'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>模块管理选项：</strong><br>控制网站启用的模块。</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <table width='100%'>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='Advertisement'" & IsModulesSelected(Modules, "Advertisement") & ">网站广告管理</td>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='FriendSite'" & IsModulesSelected(Modules, "FriendSite") & ">友情链接管理</td>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='Announce'" & IsModulesSelected(Modules, "Announce") & ">网站公告管理</td>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='Vote'" & IsModulesSelected(Modules, "Vote") & ">网站调查管理</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf

    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='KeyLink'" & IsModulesSelected(Modules, "KeyLink") & ">站内链接管理</td>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='Rtext'" & IsModulesSelected(Modules, "Rtext") & ">字符过滤管理</td>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='Collection'" & IsModulesSelected(Modules, "Collection") & ">采集管理</td>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='SMS'" & IsModulesSelected(Modules, "SMS") & ">手机短信</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    If SystemEdition = "GPS" Or SystemEdition = "EPS" Or SystemEdition = "ECS" Or SystemEdition = "IPS" Or SystemEdition = "All" Then
    Response.Write "          <td><input name='Modules' type='checkbox' value='Survey'" & IsModulesSelected(Modules, "Survey") & ">问卷调查管理</td>" & vbCrLf
    End If
    If SystemEdition = "IPS" Or SystemEdition = "All" Then
    Response.Write "            <td><input name='Modules' type='checkbox' value='Supply'" & IsModulesSelected(Modules, "Supply") & ">供求信息管理</td>" & vbCrLf
    Response.Write "            <td><input name='Modules' type='checkbox' value='House'" & IsModulesSelected(Modules, "House") & ">房产中心管理</td>" & vbCrLf
    End If
    If SystemEdition = "GPS" Or SystemEdition = "EPS" Or SystemEdition = "ECS" Or SystemEdition = "All" Then
    Response.Write "            <td><input name='Modules' type='checkbox' value='Job'" & IsModulesSelected(Modules, "Job") & ">人才招聘管理</td>" & vbCrLf
    End If
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    If SystemEdition = "eShop" Or SystemEdition = "ECS" Or SystemEdition = "All" Then
    Response.Write "            <td><input name='Modules' type='checkbox' value='CRM'" & IsModulesSelected(Modules, "CRM") & ">客户关系管理</td>" & vbCrLf
    End If
    If SystemEdition = "GPS" Or SystemEdition = "EPS" Or SystemEdition = "ECS" Or SystemEdition = "All" Then
    Response.Write "          <td><input name='Modules' type='checkbox' value='Classroom'" & IsModulesSelected(Modules, "Classroom") & ">室场登记管理</td>" & vbCrLf
    End If
    If SystemEdition = "EPS" Or SystemEdition = "All" Then
    Response.Write "          <td><input name='Modules' type='checkbox' value='Sdms'" & IsModulesSelected(Modules, "Sdms") & ">学生学籍管理</td>" & vbCrLf
    End If
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong><font color=red>网站首页的扩展名：</font></strong><br>若选择前四项，即启用了网站首页的生成HTML功能。</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='FileExt_SiteIndex' type='radio' value='0' " & IsRadioChecked(rsConfig("FileExt_SiteIndex"), 0) & ">.html &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteIndex' type='radio' value='1' " & IsRadioChecked(rsConfig("FileExt_SiteIndex"), 1) & ">.htm &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteIndex' type='radio' value='2' " & IsRadioChecked(rsConfig("FileExt_SiteIndex"), 2) & ">.shtml &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteIndex' type='radio' value='3' " & IsRadioChecked(rsConfig("FileExt_SiteIndex"), 3) & ">.shtm &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteIndex' type='radio' value='4' " & IsRadioChecked(rsConfig("FileExt_SiteIndex"), 4) & ">.asp " & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong><font color=red>全站专题的扩展名：</font></strong><br>若选择前四项，即启用了全站专题的生成HTML功能。</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='FileExt_SiteSpecial' type='radio' value='0' " & IsRadioChecked(rsConfig("FileExt_SiteSpecial"), 0) & ">.html &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteSpecial' type='radio' value='1' " & IsRadioChecked(rsConfig("FileExt_SiteSpecial"), 1) & ">.htm &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteSpecial' type='radio' value='2' " & IsRadioChecked(rsConfig("FileExt_SiteSpecial"), 2) & ">.shtml &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteSpecial' type='radio' value='3' " & IsRadioChecked(rsConfig("FileExt_SiteSpecial"), 3) & ">.shtm &nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input name='FileExt_SiteSpecial' type='radio' value='4' " & IsRadioChecked(rsConfig("FileExt_SiteSpecial"), 4) & ">.asp " & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>链接地址方式：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='SiteUrlType' type='radio' value='0' " & IsRadioChecked(rsConfig("SiteUrlType"), 0) & "> 相对路径（形如：&lt;a href='/News/200509/1358.html'&gt;标题&lt;/a&gt;）<br>&nbsp;&nbsp;&nbsp;&nbsp;当一个网站有多个域名时，一般采用此方式<br>&nbsp;&nbsp;&nbsp;&nbsp;当一个网站有多个镜像网站时，必须采用此方式<br>" & vbCrLf
    Response.Write "        <input name='SiteUrlType' type='radio' value='1' " & IsRadioChecked(rsConfig("SiteUrlType"), 1) & "> 绝对路径（形如：&lt;a href='http://www.powereasy.net/News/200509/1358.html'&gt;标题&lt;/a&gt;）<br>&nbsp;&nbsp;&nbsp;&nbsp;当要把频道做为子站点来访问时，必须使用此方式<br>&nbsp;&nbsp;&nbsp;&nbsp;要使用此方式，必须把网站URL设置正确。" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>来访限定方式：</strong><br><font color='red'>此功能只对ASP访问方式有效。如果你以前生成了HTML文件，则启用此功能后，这些HTML文件仍可以访问（除非手工删除）。可以使用此功能配合频道、栏目、及文章的权限设置和生成HTML方式来达到整站限定IP访问，或者只对有权限设置的内容进行IP限定。</font></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='LockIPType' type='radio' value='0' " & IsRadioChecked(rsConfig("LockIPType"), 0) & ">  不启用来访限定功能，任何IP都可以访问本站。<br>" & vbCrLf
    Response.Write "        <input name='LockIPType' type='radio' value='1' " & IsRadioChecked(rsConfig("LockIPType"), 1) & ">  仅仅启用白名单，只允许白名单中的IP访问本站。<br>" & vbCrLf
    Response.Write "        <input name='LockIPType' type='radio' value='2' " & IsRadioChecked(rsConfig("LockIPType"), 2) & ">  仅仅启用黑名单，只禁止黑名单中的IP访问本站。<br>" & vbCrLf
    Response.Write "        <input name='LockIPType' type='radio' value='3' " & IsRadioChecked(rsConfig("LockIPType"), 3) & ">  同时启用白名单与黑名单，先判断IP是否在白名单中，如果不在，则禁止访问；如果在则再判断是否在黑名单中，如果IP在黑名单中则禁止访问，否则允许访问。<br>" & vbCrLf
    Response.Write "        <input name='LockIPType' type='radio' value='4' " & IsRadioChecked(rsConfig("LockIPType"), 4) & ">  同时启用白名单与黑名单，先判断IP是否在黑名单中，如果不在，则允许访问；如果在则再判断是否在白名单中，如果IP在白名单中则允许访问，否则禁止访问。 " & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Response.Write "   <tr class='tdbg'>     " & vbCrLf
    Response.Write "     <td width='40%' class='tdbg5'>               <strong>IP段白名单</strong>：<br>" & vbCrLf
    Response.Write "      (注：添加多个限定IP段，请用<font color='red'>回车</font>分隔。 <br>" & vbCrLf
    Response.Write "      限制IP段的书写方式，中间请用英文四个小横杠连接，如 " & vbCrLf
    Response.Write "      <font color='red'>219.100.93.32----219.100.93.255</font> 就限定了IP 219.100.93.32 到IP 219.100.93.255 这个IP段的访问。当页面为asp方式时才有效。) </td>      " & vbCrLf
    Response.Write "     <td class='tdbg'>" & vbCrLf

    Response.Write " <textarea name='LockIPWhite' cols='60' rows='8' id='LockIP'>" & vbCrLf
    Dim rsLockIP, arrLockIP, i, arrLockIPCut
    If InStr(rsConfig("LockIP"), "|||") > 0 Then
        rsLockIP = Split(rsConfig("LockIP"), "|||")
        If InStr(rsLockIP(0), "$$$") > 0 Then
            arrLockIP = Split(Trim(rsLockIP(0)), "$$$")
            For i = 0 To UBound(arrLockIP)
                arrLockIPCut = Split(Trim(arrLockIP(i)), "----")
                Response.Write DecodeIP(arrLockIPCut(0)) & "----" & DecodeIP(arrLockIPCut(1))
                If i < UBound(arrLockIP) Then Response.Write Chr(10)
            Next
        ElseIf rsLockIP(0) <> "" Then
            arrLockIPCut = Split(Trim(rsLockIP(0)), "----")
            Response.Write DecodeIP(arrLockIPCut(0)) & "----" & DecodeIP(arrLockIPCut(1))
        End If
    End If
    Response.Write "</textarea>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "   <tr class='tdbg'>     " & vbCrLf
    Response.Write "     <td width='40%' class='tdbg5'>               <strong>IP段黑名单</strong>：<br>" & vbCrLf
    Response.Write "      (注：同上。) <br>" & vbCrLf
    Response.Write "      </td>      " & vbCrLf
    Response.Write "     <td class='tdbg'>" & vbCrLf

    Response.Write " <textarea name='LockIPBlack' cols='60' rows='8' id='LockIP'>" & vbCrLf

    If InStr(rsConfig("LockIP"), "|||") > 0 Then
        rsLockIP = Split(rsConfig("LockIP"), "|||")
        If InStr(rsLockIP(1), "$$$") > 0 Then
            arrLockIP = Split(Trim(rsLockIP(1)), "$$$")
            For i = 0 To UBound(arrLockIP)
                arrLockIPCut = Split(Trim(arrLockIP(i)), "----")
                Response.Write DecodeIP(arrLockIPCut(0)) & "----" & DecodeIP(arrLockIPCut(1))
                If i < UBound(arrLockIP) Then Response.Write Chr(10)
            Next
        ElseIf rsLockIP(1) <> "" Then
            arrLockIPCut = Split(Trim(rsLockIP(1)), "----")
            Response.Write DecodeIP(arrLockIPCut(0)) & "----" & DecodeIP(arrLockIPCut(1))
        End If
    End If
    Response.Write "</textarea>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否允许新会员注册：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableUserReg' value='1' " & IsRadioChecked(rsConfig("EnableUserReg"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableUserReg' value='0' " & IsRadioChecked(rsConfig("EnableUserReg"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>新会员注册是否需要邮件验证：</strong><br>若选择“是”，则会员注册后系统会发一封带有验证码的邮件给此会员，会员必须在通过邮件验证后才能真正成为正式注册会员</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EmailCheckReg' value='1' " & IsRadioChecked(rsConfig("EmailCheckReg"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EmailCheckReg' value='0' " & IsRadioChecked(rsConfig("EmailCheckReg"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>新会员注册是否需要管理员认证：</strong><br>若选择是，则会员必须在通过管理员认证后才能真正成功正式注册会员。</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='AdminCheckReg' value='1' " & IsRadioChecked(rsConfig("AdminCheckReg"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='AdminCheckReg' value='0' " & IsRadioChecked(rsConfig("AdminCheckReg"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>每个Email是否允许注册多次：</strong><br>若选择是，则利用同一个Email可以注册多个会员。</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableMultiRegPerEmail' value='1' " & IsRadioChecked(rsConfig("EnableMultiRegPerEmail"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableMultiRegPerEmail' value='0' " & IsRadioChecked(rsConfig("EnableMultiRegPerEmail"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>会员注册时是否启用验证码功能：</strong><br>启用验证码功能可以在一定程度上防止暴力营销软件或注册机自动注册</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableCheckCodeOfReg' value='1' " & IsRadioChecked(rsConfig("EnableCheckCodeOfReg"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableCheckCodeOfReg' value='0' " & IsRadioChecked(rsConfig("EnableCheckCodeOfReg"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>会员注册时是否启用回答问题验证功能：</strong><br>启用此功能，可以最大程度上防止暴力营销软件或注册机自动注册，也可以用于某些特殊场合，防止无关人员注册会员。</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableQAofReg' value='1' " & IsRadioChecked(rsConfig("EnableQAofReg"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableQAofReg' value='0' " & IsRadioChecked(rsConfig("EnableQAofReg"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否启用会员中心模板功能：</strong><br>如果启用会员中心模板，可以在会员模板管理中修改首页模板，如果自己修改过会员中心模板，请先添加好相应模板功能之后在启用此功能。</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowUserModel' value='1' " & IsRadioChecked(PE_CBool(rsConfig("ShowUserModel")), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='ShowUserModel' value='0' " & IsRadioChecked(PE_CBool(rsConfig("ShowUserModel")), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf	
    Dim arrQA

    arrQA = Split(rsConfig("QAofReg") & "", "$$$")
    If UBound(arrQA) <> 5 Then arrQA = Split("问题一$$$答案一$$$问题二$$$答案二$$$问题三$$$答案三", "$$$")
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>问题一：</strong><br>如果启用验证功能，则问题一和答案必须填写。</td>" & vbCrLf
    Response.Write "      <td>问题：<input type='text' name='RegQuestion1' value='" & Trim(arrQA(0)) & "' size='50'><br>答案：<input type='text' name='RegAnswer1' value='" & Trim(arrQA(1)) & "' size='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>问题二：</strong><br>问题二可以选填</td>" & vbCrLf
    Response.Write "      <td>问题：<input type='text' name='RegQuestion2' value='" & Trim(arrQA(2)) & "' size='50'><br>答案：<input type='text' name='RegAnswer2' value='" & Trim(arrQA(3)) & "' size='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>问题三：</strong><br>问题三可以选填</td>" & vbCrLf
    Response.Write "      <td>问题：<input type='text' name='RegQuestion3' value='" & Trim(arrQA(4)) & "' size='50'><br>答案：<input type='text' name='RegAnswer3' value='" & Trim(arrQA(5)) & "' size='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>新会员注册时用户名最少字符数：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='UserNameLimit' type='text' id='UserNameLimit' value='" & rsConfig("UserNameLimit") & "' size='6' maxlength='5'> 个字符</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>新会员注册时用户名最多字符数：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='UserNameMax' type='text' id='UserNameMax' value='" & rsConfig("UserNameMax") & "' size='6' maxlength='5'> 个字符</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>禁止注册的用户名：</strong><br>在右边指定的用户名将被禁止注册，每个用户名请用“|”符号分隔</td>" & vbCrLf
    Response.Write "      <td><input type='text' name='UserName_RegDisabled' value='" & rsConfig("UserName_RegDisabled") & "' size='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>会员注册时的必填项目：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <table width='100%'>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='UserName' checked disabled>用户名</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Password' checked disabled>密码</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Question' checked disabled>密码问题</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Answer' checked disabled>问题答案</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Email' checked disabled>Email</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Homepage'" & IsMustFill(RegFields_MustFill, "Homepage") & ">主页</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='QQ'" & IsMustFill(RegFields_MustFill, "QQ") & ">QQ号码</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='ICQ'" & IsMustFill(RegFields_MustFill, "ICQ") & ">ICQ号码</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='MSN'" & IsMustFill(RegFields_MustFill, "MSN") & ">MSN帐号</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='UC'" & IsMustFill(RegFields_MustFill, "UC") & ">UC号码</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='OfficePhone'" & IsMustFill(RegFields_MustFill, "OfficePhone") & ">办公电话</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='HomePhone'" & IsMustFill(RegFields_MustFill, "HomePhone") & ">家庭电话</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Mobile'" & IsMustFill(RegFields_MustFill, "Mobile") & ">手机号码</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Fax'" & IsMustFill(RegFields_MustFill, "Fax") & ">传真号码</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='PHS'" & IsMustFill(RegFields_MustFill, "PHS") & ">小灵通</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td colspan='2'><input name='RegFields_MustFill' type='checkbox' value='Region'" & IsMustFill(RegFields_MustFill, "Region") & ">国家/地区＋省市/州郡＋城市</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Address'" & IsMustFill(RegFields_MustFill, "Address") & ">联系地址</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='ZipCode'" & IsMustFill(RegFields_MustFill, "ZipCode") & ">邮政编码</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Yahoo'" & IsMustFill(RegFields_MustFill, "Yahoo") & ">雅虎通帐号</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='TrueName'" & IsMustFill(RegFields_MustFill, "TrueName") & ">真实姓名</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Birthday'" & IsMustFill(RegFields_MustFill, "Birthday") & ">出生日期</td>" & vbCrLf
    'Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Vocation'" & IsMustFill(RegFields_MustFill, "Vocation") & ">职业</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='IDCard'" & IsMustFill(RegFields_MustFill, "IDCard") & ">身份证号码</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Aim'" & IsMustFill(RegFields_MustFill, "Aim") & ">Aim帐号</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Company'" & IsMustFill(RegFields_MustFill, "Company") & ">公司/单位</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Department'" & IsMustFill(RegFields_MustFill, "Department") & ">部门</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='PosTitle'" & IsMustFill(RegFields_MustFill, "PosTitle") & ">职务</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Marriage'" & IsMustFill(RegFields_MustFill, "Marriage") & ">婚姻状况</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Income'" & IsMustFill(RegFields_MustFill, "Income") & ">收入情况</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='UserFace'" & IsMustFill(RegFields_MustFill, "UserFace") & ">用户头像</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='FaceWidth'" & IsMustFill(RegFields_MustFill, "FaceWidth") & ">头像宽度</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='FaceHeight'" & IsMustFill(RegFields_MustFill, "FaceHeight") & ">头像高度</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Sign'" & IsMustFill(RegFields_MustFill, "Sign") & ">签名档</td>" & vbCrLf
    Response.Write "            <td><input name='RegFields_MustFill' type='checkbox' value='Privacy'" & IsMustFill(RegFields_MustFill, "Privacy") & ">隐私设定</td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>新会员注册时赠送的积分：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='PresentExp' type='text' id='PresentExp' value='" & rsConfig("PresentExp") & "' size='6' maxlength='5'> 分积分</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>新会员注册时赠送的金钱：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='PresentMoney' type='text' id='PresentMoney' value='" & rsConfig("PresentMoney") & "' size='6' maxlength='5'> 元人民币（为0时不赠送）</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>新会员注册时赠送的点数：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='PresentPoint' type='text' id='PresentPoint' value='" & rsConfig("PresentPoint") & "' size='6' maxlength='5'> 点点券（为0时不赠送）</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>新会员注册时赠送的有效期：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='PresentValidNum' type='text' id='PresentValidNum' value='" & rsConfig("PresentValidNum") & "' size='6' maxlength='5'>"
    
    Response.Write "      <select name='PresentValidUnit' id='PresentValidUnit'><option value='1' "
    If rsConfig("PresentValidUnit") = 1 Then Response.Write " selected"
    Response.Write ">天</option><option value='2' "
    If rsConfig("PresentValidUnit") = 2 Then Response.Write " selected"
    Response.Write ">月</option><option value='3' "
    If rsConfig("PresentValidUnit") = 3 Then Response.Write " selected"
    Response.Write ">年</option></select>"
    
    Response.Write "（为0时不赠送，为－1表示无限期）</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>会员登录时是否启用验证码功能：</strong><br>启用验证码功能可以在一定程度上防止会员密码被暴力破解</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableCheckCodeOfLogin' value='1' " & IsRadioChecked(rsConfig("EnableCheckCodeOfLogin"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableCheckCodeOfLogin' value='0' " & IsRadioChecked(rsConfig("EnableCheckCodeOfLogin"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>会员每登录一次奖励的积分：</strong><br>一天只计算一次</td>" & vbCrLf
    Response.Write "      <td><input name='PresentExpPerLogin' type='text' id='PresentExpPerLogin' value='" & rsConfig("PresentExpPerLogin") & "' size='6' maxlength='5'> 分积分</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>会员的资金与点券的兑换比率：</strong></td>" & vbCrLf
    Response.Write "      <td>每 <input name='MoneyExchangePoint' type='text' id='MoneyExchangePoint' value='" & FormatNumber(rsConfig("MoneyExchangePoint"), 2, vbTrue, vbFalse, vbTrue) & "' size='6' maxlength='5'> 元钱可兑换 <strong><font color='#FF0000'>1</font></strong> 点点券</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>会员的资金与有效期的兑换比率：</strong></td>" & vbCrLf
    Response.Write "      <td>每 <input name='MoneyExchangeValidDay' type='text' id='MoneyExchangeValidDay' value='" & FormatNumber(rsConfig("MoneyExchangeValidDay"), 2, vbTrue, vbFalse, vbTrue) & "' size='6' maxlength='5'> 元钱可兑换 <strong><font color='#FF0000'>1</font></strong> 天有效期</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>会员的积分与点券的兑换比率：</strong></td>" & vbCrLf
    Response.Write "      <td>每 <input name='UserExpExchangePoint' type='text' id='UserExpExchangePoint' value='" & FormatNumber(rsConfig("UserExpExchangePoint"), 2, vbTrue, vbFalse, vbTrue) & "' size='6' maxlength='5'> 分积分可兑换 <strong><font color='#FF0000'>1</font></strong> 点点券</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>会员的积分与有效期的兑换比率：</strong></td>" & vbCrLf
    Response.Write "      <td>每 <input name='UserExpExchangeValidDay' type='text' id='UserExpExchangeValidDay' value='" & FormatNumber(rsConfig("UserExpExchangeValidDay"), 2, vbTrue, vbFalse, vbTrue) & "' size='6' maxlength='5'> 分积分可兑换 <strong><font color='#FF0000'>1</font></strong> 天有效期</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>点券的名称：</strong><br>例如：动易币、点券、金币</td>" & vbCrLf
    Response.Write "      <td><input name='PointName' type='text' id='PointName' value='" & rsConfig("PointName") & "' size='6' maxlength='5'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>点券的单位：</strong>例如：点、个</td>" & vbCrLf
    Response.Write "      <td><input name='PointUnit' type='text' id='PointUnit' value='" & rsConfig("PointUnit") & "' size='6' maxlength='5'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>新会员注册时发送的验证邮件内容：</strong><br>邮件内容支持HTML<br><font color='red'>标签说明：</font><br>{$CheckNum}：验证码<br>{$CheckUrl}：会员注册验证地址</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfRegCheck' cols='60' rows='5'>" & rsConfig("EmailOfRegCheck") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>邮件发送组件：</strong><br>" & vbCrLf
    Response.Write "        请一定要选择服务器上已安装的组件(打√的)<br>" & vbCrLf
    Response.Write "        如果您的服务器不支持(打×的)下列组件，请选择“无”</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <select name='MailObject' id='MailObject'>" & vbCrLf
    Response.Write "          <option value='0'" & IsOptionSelected(rsConfig("MailObject"), 0) & ">无</option>" & vbCrLf
    Response.Write "          <option value='1'" & IsOptionSelected(rsConfig("MailObject"), 1) & ">Jmail " & ShowInstalled("JMail.SMTPMail") & "</option>" & vbCrLf
    Response.Write "          <option value='2'" & IsOptionSelected(rsConfig("MailObject"), 2) & ">CDONTS " & ShowInstalled("CDONTS.NewMail") & "</option>" & vbCrLf
    Response.Write "          <option value='3'" & IsOptionSelected(rsConfig("MailObject"), 3) & ">ASPEMAIL " & ShowInstalled("Persits.MailSender") & "</option>" & vbCrLf
    Response.Write "          <option value='4'" & IsOptionSelected(rsConfig("MailObject"), 4) & ">WebEasyMail " & ShowInstalled("easymail.MailSend") & "</option>" & vbCrLf
    Response.Write "        </select>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>SMTP服务器地址：</strong><br>" & vbCrLf
    Response.Write "        用来发送邮件的SMTP服务器<br>" & vbCrLf
    Response.Write "        如果你不清楚此参数含义，请联系你的空间商 </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='MailServer' type='text' id='MailServer' value='" & rsConfig("MailServer") & "' size='40'>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>SMTP登录用户名：</strong><br>" & vbCrLf
    Response.Write "        当你的服务器需要SMTP身份验证时还需设置此参数</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='MailServerUserName' type='text' id='MailServerUserName' value='" & rsConfig("MailServerUserName") & "' size='40'>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>SMTP登录密码：</strong><br>" & vbCrLf
    Response.Write "        当你的服务器需要SMTP身份验证时还需设置此参数 </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='MailServerPassWord' type='password' id='MailServerPassWord' value='" & rsConfig("MailServerPassWord") & "' size='40'>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>SMTP域名：</strong><br>" & vbCrLf
    Response.Write "        如果用“name@domain.com”这样的用户名登录时，请指明domain.com</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='MailDomain' type='text' id='MailDomain' value='" & rsConfig("MailDomain") & "' size='40'>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>生成缩略图组件：</strong><br>请一定要选择服务器上已安装的组件</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <select name='PhotoObject' id='PhotoObject'>" & vbCrLf
    Response.Write "          <option value='0'" & IsOptionSelected(rsConfig("PhotoObject"), 0) & ">无</option>" & vbCrLf
    Response.Write "          <option value='1'" & IsOptionSelected(rsConfig("PhotoObject"), 1) & ">AspJpeg组件 " & ShowInstalled("Persits.Jpeg") & "</option>" & vbCrLf
    'Response.Write "          <option value='2'" & IsOptionSelected(rsConfig("PhotoObject"), 2) & ">SA-ImgWriter组件 " & ShowInstalled("SoftArtisans.ImageGen") & "</option>" & vbCrLf
    'Response.Write "          <option value='3'" & IsOptionSelected(rsConfig("PhotoObject"), 3) & ">SJCatSoft V2.6组件 " & ShowInstalled("sjCatSoft.Thumbnail") & "</option>" & vbCrLf
    Response.Write "        </select>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>缩略图默认宽度：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Thumb_DefaultWidth' type='text' value='" & rsConfig("Thumb_DefaultWidth") & "' size='10' maxlength='10'  ONKEYPRESS=""javascript:event.returnValue=IsDigit()""> 像素&nbsp;&nbsp;&nbsp;&nbsp;设为0时，将以高度为准按比例缩小。</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>缩略图默认高度：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Thumb_DefaultHeight' type='text' value='" & rsConfig("Thumb_DefaultHeight") & "' size='10' maxlength='10'  ONKEYPRESS=""javascript:event.returnValue=IsDigit()""> 像素&nbsp;&nbsp;&nbsp;&nbsp;设为0时，将以宽度为准按比例缩小。</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>缩略图算法：</strong></td>" & vbCrLf
    Response.Write "      <td><input type='radio' name='Thumb_Arithmetic' value='0' " & IsRadioChecked(rsConfig("Thumb_Arithmetic"), 0) & "> 常规算法：宽度和高度都大于0时，直接缩小成指定大小，其中一个为0时，按比例缩小<br>" & vbCrLf
    Response.Write "        <input type='radio' name='Thumb_Arithmetic' value='1' " & IsRadioChecked(rsConfig("Thumb_Arithmetic"), 1) & "> 裁剪法：宽度和高度都大于0时，先按最佳比例缩小再裁剪成指定大小，其中一个为0时，按比例缩小。<br>" & vbCrLf
    Response.Write "        <input type='radio' name='Thumb_Arithmetic' value='2' " & IsRadioChecked(rsConfig("Thumb_Arithmetic"), 2) & "> 补充法：在指定大小的背景图上附加上按最佳比例缩小的图片。</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='ThumbBackgroundColor' " & ISdisplay(rsConfig("Thumb_Arithmetic"), 2) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>缩略图底色：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Thumb_BackgroundColor' type='text' value='" & rsConfig("Thumb_BackgroundColor") & "' size='10' maxlength='10'><img border=0 src='../Editor/images/rect.gif' width=18 style='cursor:hand;backgroundColor:" & rsConfig("Thumb_BackgroundColor") & "' id=s_bordercolor onClick='SelectColor(this,Thumb_BackgroundColor)'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>图像质量：</strong><br>缩略图及加水印后的图像质量</td>" & vbCrLf
    Response.Write "      <td><input name='PhotoQuality' type='text' value='" & rsConfig("PhotoQuality") & "' size='10' maxlength='10'  ONKEYPRESS=""javascript:event.returnValue=IsDigit()""> &nbsp;&nbsp;&nbsp;&nbsp;请输入50－100间的数字。数字越大，图像质量越好。建议设为90。</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf

    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>水印类型：</strong></td>" & vbCrLf
    Response.Write "      <td><input type='radio' name='Watermark_Type' value='0'  " & IsRadioChecked(rsConfig("Watermark_Type"), 0) & " onClick=""PE_Watermark_Text.style.display='';PE_Watermark_Text_FontName.style.display='';PE_Watermark_Text_FontSize.style.display='';PE_Watermark_Text_FontColor.style.display='';PE_Watermark_Text_Bold.style.display='';PE_Watermark_Images_FileName.style.display='none';PE_Watermark_Images_Transparence.style.display='none';PE_Watermark_Images_BackgroundColor.style.display='none'"" > 文字水印&nbsp;&nbsp;"
    Response.Write "          <input type='radio' name='Watermark_Type' value='1'  " & IsRadioChecked(rsConfig("Watermark_Type"), 1) & " onClick=""PE_Watermark_Text.style.display='none';PE_Watermark_Text_FontName.style.display='none';PE_Watermark_Text_FontSize.style.display='none';PE_Watermark_Text_FontColor.style.display='none';PE_Watermark_Text_Bold.style.display='none';PE_Watermark_Images_FileName.style.display='';PE_Watermark_Images_Transparence.style.display='';PE_Watermark_Images_BackgroundColor.style.display=''"" > 图片水印</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Text' " & ISdisplay(rsConfig("Watermark_Type"), 0) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>水印文字：</strong><br>水印文字字数不宜超过15个字符，不支持任何WEB编码标记</td>" & vbCrLf
    Response.Write "      <td><input name='Watermark_Text' type='text' value='" & rsConfig("Watermark_Text") & "' size='40' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Text_FontName' " & ISdisplay(rsConfig("Watermark_Type"), 0) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>文字字体：</strong></td>" & vbCrLf
    Response.Write "      <td>"
    Response.Write "        <SELECT name=""Watermark_Text_FontName"" >" & vbCrLf
    Response.Write "            <option value=""宋体"" " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "宋体") & ">宋体</option>" & vbCrLf
    Response.Write "            <option value=""楷体_GB2312"" " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "楷体_GB2312") & ">楷体</option>" & vbCrLf
    Response.Write "            <option value=""仿宋_GB2312"" " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "仿宋_GB2312") & ">新宋体</option>" & vbCrLf
    Response.Write "            <option value=""黑体"" " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "黑体") & ">黑体</option>" & vbCrLf
    Response.Write "            <option value=""隶书"" " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "隶书") & ">隶书</option>" & vbCrLf
    Response.Write "            <option value=""幼圆"" " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "幼圆") & ">幼圆</option>" & vbCrLf
    Response.Write "            <option value=""Andale Mono"" " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Andale Mono") & ">Andale Mono</OPTION> " & vbCrLf
    Response.Write "            <option value=""Arial""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Arial") & ">Arial</OPTION> " & vbCrLf
    Response.Write "            <option value=""Arial Black""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Arial Black") & ">Arial Black</OPTION> " & vbCrLf
    Response.Write "            <option value=""Book Antiqua""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Book Antiqua") & ">Book Antiqua</OPTION>" & vbCrLf
    Response.Write "            <option value=""Century Gothic""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Century Gothic") & ">Century Gothic</OPTION> " & vbCrLf
    Response.Write "            <option value=""Comic Sans MS""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Comic Sans MS") & ">Comic Sans MS</OPTION>" & vbCrLf
    Response.Write "            <option value=""Courier New""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Courier New") & ">Courier New</OPTION>" & vbCrLf
    Response.Write "            <option value=""Georgia""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Georgia") & ">Georgia</OPTION>" & vbCrLf
    Response.Write "            <option value=""Impact""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Impact") & ">Impact</OPTION>" & vbCrLf
    Response.Write "            <option value=""Tahoma""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Tahoma") & ">Tahoma</OPTION>" & vbCrLf
    Response.Write "            <option value=""Times New Roman""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Times New Roman") & ">Times New Roman</OPTION>" & vbCrLf
    Response.Write "            <option value=""Trebuchet MS""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Trebuchet MS") & ">Trebuchet MS</OPTION>" & vbCrLf
    Response.Write "            <option value=""Script MT Bold""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Script MT Bold") & ">Script MT Bold</OPTION>" & vbCrLf
    Response.Write "            <option value=""Stencil""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Stencil") & ">Stencil</OPTION>" & vbCrLf
    Response.Write "            <option value=""Verdana""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Verdana") & ">Verdana</OPTION>" & vbCrLf
    Response.Write "            <option value=""Lucida Console""  " & IsOptionSelected(rsConfig("Watermark_Text_FontName"), "Lucida Console") & ">Lucida Console</OPTION>" & vbCrLf
    Response.Write "        </SELECT>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Text_FontSize' " & ISdisplay(rsConfig("Watermark_Type"), 0) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>文字大小：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Watermark_Text_FontSize' type='text' value='" & rsConfig("Watermark_Text_FontSize") & "' size='10' maxlength='10'  ONKEYPRESS=""javascript:event.returnValue=IsDigit()""> 像素</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Text_FontColor' " & ISdisplay(rsConfig("Watermark_Type"), 0) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>文字颜色：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Watermark_Text_FontColor' type='text' value='" & rsConfig("Watermark_Text_FontColor") & "' size='10' maxlength='10'><img border=0 src='../Editor/images/rect.gif' width=18 style='cursor:hand;backgroundColor:" & rsConfig("Watermark_Text_FontColor") & "' id=s_bordercolor onClick='SelectColor(this,Watermark_Text_FontColor)'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Text_Bold' " & ISdisplay(rsConfig("Watermark_Type"), 0) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否粗体：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "          <SELECT name='Watermark_Text_Bold' >" & vbCrLf
    Response.Write "            <OPTION value='0'  " & IsOptionSelected(rsConfig("Watermark_Text_Bold"), False) & ">否</OPTION>" & vbCrLf
    Response.Write "            <OPTION value='1'  " & IsOptionSelected(rsConfig("Watermark_Text_Bold"), True) & ">是</OPTION>" & vbCrLf
    Response.Write "          </SELECT>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Images_FileName' " & ISdisplay(rsConfig("Watermark_Type"), 1) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>水印图片文件名：</strong><br>这里请填写图片文件的相对路径，以“\”开头</td>" & vbCrLf
    Response.Write "      <td><input name='Watermark_Images_FileName' type='text' value='" & rsConfig("Watermark_Images_FileName") & "' size='40' maxlength='40'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Images_Transparence' " & ISdisplay(rsConfig("Watermark_Type"), 1) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>图片透明度：</strong><br> 100% 为不透明</td>" & vbCrLf
    Response.Write "      <td><input name='Watermark_Images_Transparence' type='text' value='" & rsConfig("Watermark_Images_Transparence") & "' size='3' maxlength='3'  ONKEYPRESS=""javascript:event.returnValue=IsDigit()"">%</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg' id='PE_Watermark_Images_BackgroundColor' " & ISdisplay(rsConfig("Watermark_Type"), 1) & ">" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>图片底色：</strong><br>若想去除水印图片的底色，请在此填入底色的RGB值。</td>" & vbCrLf
    Response.Write "      <td><input name='Watermark_Images_BackgroundColor' type='text' value='" & rsConfig("Watermark_Images_BackgroundColor") & "' size='10' maxlength='10'><img border=0 src='../Editor/images/rect.gif' width=18 style='cursor:hand;backgroundColor:" & rsConfig("Watermark_Images_BackgroundColor") & "' id=s_bordercolor onClick='SelectColor(this,Watermark_Images_BackgroundColor)'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>坐标起点位置：</strong></td>" & vbCrLf
    Response.Write "      <td>"
    Response.Write "        <SELECT NAME='Watermark_Position' >" & vbCrLf
    Response.Write "            <option value='0' " & IsOptionSelected(rsConfig("Watermark_Position"), 0) & ">左上</option>" & vbCrLf
    Response.Write "            <option value='1' " & IsOptionSelected(rsConfig("Watermark_Position"), 1) & ">左下</option>" & vbCrLf
    Response.Write "            <option value='2' " & IsOptionSelected(rsConfig("Watermark_Position"), 2) & ">居中</option>" & vbCrLf
    Response.Write "            <option value='3' " & IsOptionSelected(rsConfig("Watermark_Position"), 3) & ">右上</option>" & vbCrLf
    Response.Write "            <option value='4' " & IsOptionSelected(rsConfig("Watermark_Position"), 4) & ">右下</option>" & vbCrLf
    Response.Write "        </SELECT>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>坐标位置：&nbsp;</strong>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td>X：<input name='Watermark_Position_X' type='text' value='" & rsConfig("Watermark_Position_X") & "' size='10' maxlength='10'  ONKEYPRESS=""javascript:event.returnValue=IsDigit()""> 像素<br>Y：<input name='Watermark_Position_Y' type='text' value='" & rsConfig("Watermark_Position_Y") & "' size='10' maxlength='10'  ONKEYPRESS=""javascript:event.returnValue=IsDigit()""> 像素</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>每次搜索时间间隔</strong>：<br>" & vbCrLf
    Response.Write "        设置合理的每次搜索时间间隔，可以避免恶意搜索而消耗大量系统资源</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='SearchInterval' type='text' id='SearchInterval' value='" & rsConfig("SearchInterval") & "' size='10' maxlength='10'> 秒" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>搜索返回最多的结果数</strong>：<br>" & vbCrLf
    Response.Write "        返回搜索的结果数和消耗的资源成正比，请合理设置，建议不要设置过大</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='SearchResultNum' type='text' id='SearchResultNum' value='" & rsConfig("SearchResultNum") & "' size='10' maxlength='10'> 条记录" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>通用搜索页的每页信息数</strong>：</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='MaxPerPage_SearchResult' type='text' id='MaxPerPage_SearchResult' value='" & rsConfig("MaxPerPage_SearchResult") & "' size='10' maxlength='10'> 条记录/页" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否启用全文搜索</strong><br>" & vbCrLf
    Response.Write "        ACCESS数据库不建议开启<BR>SQL数据库做了全文索引可以开启" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='SearchContent' value='1' " & IsRadioChecked(rsConfig("SearchContent"), True) & "> 启用&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input type='radio' name='SearchContent' value='0' " & IsRadioChecked(rsConfig("SearchContent"), False) & "> 禁用" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否允许游客购买商品：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableGuestBuy' value='1' " & IsRadioChecked(rsConfig("EnableGuestBuy"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableGuestBuy' value='0' " & IsRadioChecked(rsConfig("EnableGuestBuy"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>商品价格是否含税：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='IncludeTax' value='1' " & IsRadioChecked(rsConfig("IncludeTax"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='IncludeTax' value='0' " & IsRadioChecked(rsConfig("IncludeTax"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>税率设置：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='TaxRate' type='text' id='TaxRate' value='" & rsConfig("TaxRate") & "'  size='6' maxlength='6' style='text-align:center'>%</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    
'    Response.Write "    <tr class='tdbg'>" & vbCrLf
'    Response.Write "      <td width='40%' class='tdbg5'><strong>在线支付平台：</strong></td>" & vbCrLf
'    Response.Write "      <td><select name='PayOnlineProvider' id='PayOnlineProvider'>" & vbCrLf
'    Response.Write "          <option value='0'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 0) & ">无</option>" & vbCrLf
'    Response.Write "          <option value='1'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 1) & ">网银在线1.1版</option>" & vbCrLf
'    Response.Write "          <option value='2'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 2) & ">中国在线支付网</option>" & vbCrLf
'    Response.Write "          <option value='3'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 3) & ">上海环迅IPS</option>" & vbCrLf
'    Response.Write "          <option value='4'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 4) & ">广东银联</option>" & vbCrLf
'    Response.Write "          <option value='5'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 5) & ">西部支付</option>" & vbCrLf
'    Response.Write "          <option value='6'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 6) & ">易付通</option>" & vbCrLf
'    Response.Write "          <option value='7'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 7) & ">云网支付</option>" & vbCrLf
'    Response.Write "          <option value='8'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 8) & ">支付宝支付</option>" & vbCrLf
'    Response.Write "          <option value='9'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 9) & ">快钱支付</option>" & vbCrLf
'    Response.Write "          <option value='10'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 10) & ">网银在线2.0版</option>" & vbCrLf
'    Response.Write "          <option value='11'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 11) & ">快钱神州行</option>" & vbCrLf
'    Response.Write "          <option value='13'" & IsOptionSelected(rsConfig("PayOnlineProvider"), 13) & ">财付通</option>" & vbCrLf
'    Response.Write "        </select>" & vbCrLf
'    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='http://www.powereasy.net/payreg.html' target='_blank'>点此注册网银在线商户</a>"
'    Response.Write "      </td>" & vbCrLf
'    Response.Write "    </tr>" & vbCrLf
'    Response.Write "    <tr class='tdbg'>" & vbCrLf
'    Response.Write "      <td width='40%' class='tdbg5'><strong>商户编号：</strong><br>请填入您在上述在线支付平台申请的商户编号</td>" & vbCrLf
'    Response.Write "      <td><input name='PayOnlineShopID' type='text' id='PayOnlineShopID' value='" & rsConfig("PayOnlineShopID") & "' size='30' maxlength='50'></td>" & vbCrLf
'    Response.Write "    </tr>" & vbCrLf
'    Response.Write "    <tr class='tdbg'>" & vbCrLf
'    Response.Write "      <td width='40%' class='tdbg5'><strong>MD5私钥：</strong><br>请填入您在上述在线支付平台中设置的MD5私钥<br>部分在线支付平台不需要此项</td>" & vbCrLf
'    Response.Write "      <td><input name='PayOnlineKey' type='password' id='PayOnlineKey' value='" & rsConfig("PayOnlineKey") & "' size='30' maxlength='255'></td>" & vbCrLf
'    Response.Write "    </tr>" & vbCrLf
'    Response.Write "    <tr class='tdbg'>" & vbCrLf
'    Response.Write "      <td width='40%' class='tdbg5'><strong>手续费率：</strong></td>" & vbCrLf
'    Response.Write "      <td>" & vbCrLf
'    Response.Write "        <input name='PayOnlineRate' type='text' id='PayOnlineRate' value='" & rsConfig("PayOnlineRate") & "' size='6' maxlength='6' style='text-align:center'>%<br>" & vbCrLf
'    Response.Write "        <input name='PayOnlinePlusPoundage' type='checkbox' value='1' " & IsRadioChecked(rsConfig("PayOnlinePlusPoundage"), True) & "> 手续费由付款人额外支付" & vbCrLf
'    Response.Write "      </td>" & vbCrLf
'    Response.Write "    </tr>" & vbCrLf

    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>订单编号前缀：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Prefix_OrderFormNum' type='text' id='Prefix_OrderFormNum' value='" & rsConfig("Prefix_OrderFormNum") & "' size='6' maxlength='4'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>在线支付单编号前缀：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Prefix_PaymentNum' type='text' id='Prefix_PaymentNum' value='" & rsConfig("Prefix_PaymentNum") & "' size='6' maxlength='4'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>我所在的国家：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='Country' type='text' id='Country' value='" & rsConfig("Country") & "' size='15' maxlength='30'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>我所在的省份：</strong></td>" & vbCrLf
    Response.Write "      <td><select name='Province'>" & GetProvince(rsConfig("Province")) & "</select></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>我所在的城市或地区：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='City' type='text' id='City' value='" & rsConfig("City") & "' size='15' maxlength='30'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>我所在地区的邮政编码：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='PostCode' type='text' id='PostCode' value='" & rsConfig("PostCode") & "' size='10' maxlength='10'> <font color='red'>用于自动计算订单的运费</font></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>确认订单时站内短信/Email通知内容：</strong><br>支持HTML代码，可用标签详见下面的标签说明</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfOrderConfirm' cols='60' rows='4'>" & rsConfig("EmailOfOrderConfirm") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>收到银行汇款后站内短信/Email通知内容：</strong><br>支持HTML代码，可用标签详见下面的标签说明</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfReceiptMoney' cols='60' rows='4'>" & rsConfig("EmailOfReceiptMoney") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>退款后站内短信/Email通知内容：</strong><br>支持HTML代码，可用标签详见下面的标签说明</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfRefund' cols='60' rows='4'>" & rsConfig("EmailOfRefund") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>开发票后站内短信/Email通知内容：</strong><br>支持HTML代码，可用标签详见下面的标签说明</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfInvoice' cols='60' rows='4'>" & rsConfig("EmailOfInvoice") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>发出货物后站内短信/Email通知内容：</strong><br>支持HTML代码，可用标签详见下面的标签说明</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfDeliver' cols='60' rows='4'>" & rsConfig("EmailOfDeliver") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>发送卡号后站内短信/Email通知内容：</strong><br>支持HTML代码，可用标签详见下面的标签说明<br>特别标签：<br>{$CardInfo}：购买的卡号及密码信息</td>" & vbCrLf
    Response.Write "      <td><textarea name='EmailOfSendCard' cols='60' rows='4'>" & rsConfig("EmailOfSendCard") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>通知内容中的可用标签及含义：</strong></td>" & vbCrLf
    Response.Write "      <td><textarea name='Labels' cols='60' rows='4' ReadOnly>"
    Response.Write "{$OrderFormID}：订单ID" & vbCrLf
    Response.Write "{$OrderFormNum}：订单编号" & vbCrLf
    Response.Write "{$ContacterName}：收货人姓名" & vbCrLf
    Response.Write "{$OrderInfo}：订单信息" & vbCrLf
    Response.Write "{$MoneyTotal}：订单总金额" & vbCrLf
    Response.Write "{$MoneyReceipt}：订单已收款" & vbCrLf
    Response.Write "{$MoneyNeedPay}：需要支付金额" & vbCrLf
    Response.Write "{$InputTime}：订购时间" & vbCrLf
    Response.Write "{$UserName}：会员用户名" & vbCrLf
    Response.Write "{$Address}：收货人地址" & vbCrLf
    Response.Write "{$ZipCode}：收货人邮编" & vbCrLf
    Response.Write "{$Mobile}：收货人手机" & vbCrLf
    Response.Write "{$Phone}：收货人电话" & vbCrLf
    Response.Write "{$Email}：收货人Email" & vbCrLf
    Response.Write "{$PaymentType}：付款方式" & vbCrLf
    Response.Write "{$DeliverType}：送货方式" & vbCrLf
    Response.Write "{$OrderStatus}：订单状态" & vbCrLf
    Response.Write "{$PayStatus}：付款情况" & vbCrLf
    Response.Write "{$DeliverStatus}：物流状态" & vbCrLf
    Response.Write "{$Charge_Deliver}：运费" & vbCrLf
    Response.Write "{$PresentMoney}：返还现金券" & vbCrLf
    Response.Write "{$PresentExp}：赠送积分" & vbCrLf
    Response.Write "{$Charge_Deliver}：运费" & vbCrLf
    Response.Write "{$ExpressCompany}：物流公司名称" & vbCrLf
    Response.Write "{$ExpressNumber}：快递单号" & vbCrLf
    Response.Write "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否允许游客留言：</strong><br>若选择否，则游客或未登录用户不能签写留言</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='GuestBook_EnableVisitor' value='1' " & IsRadioChecked(rsConfig("GuestBook_EnableVisitor"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='GuestBook_EnableVisitor' value='0' " & IsRadioChecked(rsConfig("GuestBook_EnableVisitor"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否启用留言板验证码功能：</strong><br>用户签写留言时需要填写系统随机生成的验证码，此功能有利于预防他人恶意群发垃圾留言</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='GuestBookCheck' value='1' " & IsRadioChecked(rsConfig("GuestBookCheck"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='GuestBookCheck' value='0' " & IsRadioChecked(rsConfig("GuestBookCheck"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否公开访客IP：</strong><br>若选择是，则浏览者可以看到留言人的相关信息</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='GuestBook_ShowIP' value='1' " & IsRadioChecked(rsConfig("GuestBook_ShowIP"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='GuestBook_ShowIP' value='0' " & IsRadioChecked(rsConfig("GuestBook_ShowIP"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>必须在指定类别中留言：</strong><br>若选择在否，则留言人可以让所签写的留言不属于任何类别。</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='GuestBook_IsAssignSort' value='1' " & IsRadioChecked(rsConfig("GuestBook_IsAssignSort"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='GuestBook_IsAssignSort' value='0' " & IsRadioChecked(rsConfig("GuestBook_IsAssignSort"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否启用屏蔽垃圾广告功能：</strong><br>若选择是，下面要屏蔽的关键字的设置才有效。</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='GuestBook_EnableManageRubbish' value='1' " & IsRadioChecked(rsConfig("GuestBook_EnableManageRubbish"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='GuestBook_EnableManageRubbish' value='0' " & IsRadioChecked(rsConfig("GuestBook_EnableManageRubbish"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "   <tr class='tdbg'>     " & vbCrLf
    Response.Write "     <td width='40%' class='tdbg5'>               <strong>要屏蔽的关键字</strong>：<br>" & vbCrLf
    Response.Write "      (注：添加多个限制关键字，请用回车分隔。<br>如果用户提交的留言内容中含有要屏蔽的关键字，则会提示禁止留言！)<br> </td> " & vbCrLf
    Response.Write "     <td><textarea name='LockRubbish' cols='50' rows='8' id='LockRubbish'>" & vbCrLf
    Dim rsLockRubbish, arrLockRubbish
    rsLockRubbish = Trim(rsConfig("GuestBook_ManageRubbish"))

    If InStr(rsLockRubbish, "$$$") > 0 Then
        arrLockRubbish = Split(Trim(rsLockRubbish), "$$$")
        For i = 0 To UBound(arrLockRubbish)
            Response.Write arrLockRubbish(i)
            If i < UBound(arrLockRubbish) Then Response.Write Chr(10)
        Next
    Else
        Response.Write rsLockRubbish
    End If

    Response.Write "</textarea>" & vbCrLf
    Response.Write "     </td>    </tr>   " & vbCrLf

    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Dim GuestBook_MaxPerPage, MaxPerPage
    GuestBook_MaxPerPage = Array(20, 8, 6, 5)
    If Trim(rsConfig("GuestBook_MaxPerPage")) <> "" And Not IsNull(rsConfig("GuestBook_MaxPerPage")) Then
        MaxPerPage = Split(Trim(rsConfig("GuestBook_MaxPerPage")), "|||")
        If UBound(MaxPerPage) = 3 Then GuestBook_MaxPerPage = MaxPerPage
    End If
    Response.Write "      <td width='40%' class='tdbg5'><strong>留言讨论区方式每页显示多少条：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='GuestBook_DiscussionMaxPerPage' type='text' id='GuestBook_DiscussionMaxPerPage' value='" & GuestBook_MaxPerPage(0) & "' size='6' maxlength='5'>  条</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>留言留言本方式每页显示多少条：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='GuestBook_GuestBookMaxPerPage' type='text' id='GuestBook_GuestBookMaxPerPage' value='" & GuestBook_MaxPerPage(1) & "' size='6' maxlength='5'>  条</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>留言回复页每页显示多少条：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='GuestBook_ReplyMaxPerPage' type='text' id='GuestBook_ReplyMaxPerPage' value='" & GuestBook_MaxPerPage(2) & "' size='6' maxlength='5'>  条</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>主题展开树每页显示多少条：</strong></td>" & vbCrLf
    Response.Write "      <td><input name='GuestBook_TreeMaxPerPage' type='text' id='GuestBook_TreeMaxPerPage' value='" & GuestBook_MaxPerPage(3) & "' size='6' maxlength='5'>  条</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>网站是否启用RSS功能：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableRss' value='1' " & IsRadioChecked(rsConfig("EnableRss"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableRss' value='0' " & IsRadioChecked(rsConfig("EnableRss"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr id='RssSetting' class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>Rss使用编码：</strong>：<br>" & vbCrLf
    Response.Write "        Rss使用的汉字编码</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='RssCodeType' value='1' " & IsRadioChecked(rsConfig("RssCodeType"), True) & "> GB2312&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='RssCodeType' value='0' " & IsRadioChecked(rsConfig("RssCodeType"), False) & "> UTF-8" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>网站是否启用WAP(手机访问）功能：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableWap' value='1' " & IsRadioChecked(rsConfig("EnableWap"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='EnableWap' value='0' " & IsRadioChecked(rsConfig("EnableWap"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr id='WapSetting' class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>WAP浏览LOGO</strong>：<br>" & vbCrLf
    Response.Write "        使用手机浏览时显示的网站LOGO。<br>建议使用大部分老式手机都支持的WBMP格式图片，<br>使用这种格式的图片需要在服务器上增加MIME类型<br>wbmp&nbsp;image/vnd.wap.wbmp<br>留空则默认显示网站名称</td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input name='WapLogo' type='text' id='WapLogo' value='" & rsConfig("WapLogo") & "' size='40'>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr id='WapSetting2' class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否启用评论功能</strong>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='EnableWapPl' value='1' " & IsRadioChecked(rsConfig("EnableWapPl"), True) & "> 启用&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input type='radio' name='EnableWapPl' value='0' " & IsRadioChecked(rsConfig("EnableWapPl"), False) & "> 禁用" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr id='WapSetting3' class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否启用附件显示</strong><br>" & vbCrLf
    Response.Write "        某些老式手机不能支持彩色图片显示，如考虑兼容性，建议关闭，目前彩屏版手机则无须关闭。" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowWapAppendix' value='1' " & IsRadioChecked(rsConfig("ShowWapAppendix"), True) & "> 启用&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input type='radio' name='ShowWapAppendix' value='0' " & IsRadioChecked(rsConfig("ShowWapAppendix"), False) & "> 禁用" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr id='WapSetting4' class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否启用手机商城</strong><br>" & vbCrLf
    Response.Write "        如启用此项，请强制用户注册时必须填写联系地址，邮政编码，联系电话这三项。" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowWapShop' value='1' " & IsRadioChecked(rsConfig("ShowWapShop"), True) & "> 启用&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input type='radio' name='ShowWapShop' value='0' " & IsRadioChecked(rsConfig("ShowWapShop"), False) & "> 禁用" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr id='WapSetting5' class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>是否启用WAP进行后台管理</strong>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='ShowWapManage' value='1' " & IsRadioChecked(rsConfig("ShowWapManage"), True) & "> 启用&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    Response.Write "        <input type='radio' name='ShowWapManage' value='0' " & IsRadioChecked(rsConfig("ShowWapManage"), False) & "> 禁用" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf

    Response.Write "  <tbody id='Tabs' style='display:none'>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>动易短信通的用户名：</strong><br>请填入您在 动易短信通平台 注册的用户名</td>" & vbCrLf
    Response.Write "      <td><input name='SMSUserName' type='text' id='SMSUserName' value='" & rsConfig("SMSUserName") & "' size='30' maxlength='50'> &nbsp;&nbsp;<a href='http://sms.powereasy.net/Register.aspx' target='_blank'><font color='red'>点此注册新用户</font></a> &nbsp;&nbsp;<a href='http://sms.powereasy.net/' target='_blank'><font color='blue'>什么是动易短信通？</font></a></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>MD5密钥：</strong><br>请填入您在 动易短信通平台 中设置的MD5密钥</td>" & vbCrLf
    Response.Write "      <td><input name='SMSKey' type='password' id='SMSKey' value='" & rsConfig("SMSKey") & "' size='30' maxlength='50'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>客户提交订单时，系统是否自动发送手机短信通知管理员：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='SendMessageToAdminWhenOrder' value='1' " & IsRadioChecked(rsConfig("SendMessageToAdminWhenOrder"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='SendMessageToAdminWhenOrder' value='0' " & IsRadioChecked(rsConfig("SendMessageToAdminWhenOrder"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>管理员的小灵通或手机号码：</strong><br>每行输入一个号码。<br>可以输入多个号码，系统将同时发送到多个号码上</td>" & vbCrLf
    Response.Write "      <td><textarea name='Mobiles' cols='60' rows='4'>" & rsConfig("Mobiles") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>客户下订单时系统给管理员发送短信的内容：</strong><br>不支持HTML代码，可用标签详见下面的标签说明</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfOrder' cols='60' rows='4'>" & rsConfig("MessageOfOrder") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>客户在线支付成功后是否给客户发送手机短信，告知其卡号和密码：</strong></td>" & vbCrLf
    Response.Write "      <td>" & vbCrLf
    Response.Write "        <input type='radio' name='SendMessageToMemberWhenPaySuccess' value='1' " & IsRadioChecked(rsConfig("SendMessageToMemberWhenPaySuccess"), True) & "> 是 &nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "        <input type='radio' name='SendMessageToMemberWhenPaySuccess' value='0' " & IsRadioChecked(rsConfig("SendMessageToMemberWhenPaySuccess"), False) & "> 否" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>确认订单时手机短信通知内容：</strong><br>不支持HTML代码，可用标签详见下面的标签说明</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfOrderConfirm' cols='60' rows='4'>" & rsConfig("MessageOfOrderConfirm") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>收到银行汇款后手机短信通知内容：</strong><br>不支持HTML代码，可用标签详见下面的标签说明</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfReceiptMoney' cols='60' rows='4'>" & rsConfig("MessageOfReceiptMoney") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>退款后手机短信通知内容：</strong><br>不支持HTML代码，可用标签详见下面的标签说明</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfRefund' cols='60' rows='4'>" & rsConfig("MessageOfRefund") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>开发票后手机短信通知内容：</strong><br>不支持HTML代码，可用标签详见下面的标签说明</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfInvoice' cols='60' rows='4'>" & rsConfig("MessageOfInvoice") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>发出货物后手机短信通知内容：</strong><br>不支持HTML代码，可用标签详见下面的标签说明</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfDeliver' cols='60' rows='4'>" & rsConfig("MessageOfDeliver") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>发送卡号后手机短信通知内容：</strong><br>不支持HTML代码，可用标签详见下面的标签说明<br>特别标签：<br>{$CardInfo}：购买的卡号及密码信息</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfSendCard' cols='60' rows='4'>" & rsConfig("MessageOfSendCard") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>通知内容中的可用标签及含义：</strong></td>" & vbCrLf
    Response.Write "      <td><textarea name='Labels' cols='60' rows='4' ReadOnly>"
    Response.Write "{$OrderFormID}：订单ID" & vbCrLf
    Response.Write "{$OrderFormNum}：订单编号" & vbCrLf
    Response.Write "{$ContacterName}：收货人姓名" & vbCrLf
    Response.Write "{$OrderInfo}：订单信息" & vbCrLf
    Response.Write "{$MoneyTotal}：订单总金额" & vbCrLf
    Response.Write "{$MoneyReceipt}：订单已收款" & vbCrLf
    Response.Write "{$MoneyNeedPay}：需要支付金额" & vbCrLf
    Response.Write "{$InputTime}：订购时间" & vbCrLf
    Response.Write "{$UserName}：会员用户名" & vbCrLf
    Response.Write "{$Address}：收货人地址" & vbCrLf
    Response.Write "{$ZipCode}：收货人邮编" & vbCrLf
    Response.Write "{$Mobile}：收货人手机" & vbCrLf
    Response.Write "{$Phone}：收货人电话" & vbCrLf
    Response.Write "{$Email}：收货人Email" & vbCrLf
    Response.Write "{$PaymentType}：付款方式" & vbCrLf
    Response.Write "{$DeliverType}：送货方式" & vbCrLf
    Response.Write "{$OrderStatus}：订单状态" & vbCrLf
    Response.Write "{$PayStatus}：付款情况" & vbCrLf
    Response.Write "{$DeliverStatus}：物流状态" & vbCrLf
    Response.Write "{$Charge_Deliver}：运费" & vbCrLf
    Response.Write "{$PresentMoney}：返还现金券" & vbCrLf
    Response.Write "{$PresentExp}：赠送积分" & vbCrLf
    Response.Write "{$Charge_Deliver}：运费" & vbCrLf
    Response.Write "{$ExpressCompany}：物流公司名称" & vbCrLf
    Response.Write "{$ExpressNumber}：快递单号" & vbCrLf
    Response.Write "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'>&nbsp;</td>" & vbCrLf
    Response.Write "      <td> </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>给会员添加银行汇款记录时发送的手机短信内容：</strong><br>不支持HTML代码，可用标签：<br>{$UserName}：会员用户名<br>{$Balance}：资金余额<br>{$ReceiptDate}：到款日期<br>{$Money}：汇款金额<br>{$BankName}：汇入银行</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfAddRemit' cols='60' rows='4'>" & rsConfig("MessageOfAddRemit") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>给会员添加其他收入记录时发送的手机短信内容：</strong><br>不支持HTML代码，可用标签：<br>{$UserName}：会员用户名<br>{$Balance}：资金余额<br>{$Money}：收入金额<br>{$Reason}：原因</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfAddIncome' cols='60' rows='4'>" & rsConfig("MessageOfAddIncome") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>给会员添加支出记录时发送的手机短信内容：</strong><br>不支持HTML代码，可用标签：<br>{$UserName}：会员用户名<br>{$Balance}：资金余额<br>{$Money}：支出金额<br>{$Reason}：原因</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfAddPayment' cols='60' rows='4'>" & rsConfig("MessageOfAddPayment") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>给会员兑换点券时发送的手机短信内容：</strong><br>不支持HTML代码，可用标签：<br>{$UserName}：会员用户名<br>{$Balance}：资金余额<br>{$UserPoint}：可用点券<br>{$Money}：支出金额<br>{$Point}：得到的点券数</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfExchangePoint' cols='60' rows='4'>" & rsConfig("MessageOfExchangePoint") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>给会员奖励点券时发送的手机短信内容：</strong><br>不支持HTML代码，可用标签：<br>{$UserName}：会员用户名<br>{$UserPoint}：可用点券<br>{$Point}：增加的点券数<br>{$Reason}：奖励原因</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfAddPoint' cols='60' rows='4'>" & rsConfig("MessageOfAddPoint") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>给会员扣除点券时发送的手机短信内容：</strong><br>不支持HTML代码，可用标签：<br>{$UserName}：会员用户名<br>{$UserPoint}：可用点券<br>{$Point}：扣除的点券数<br>{$Reason}：扣除原因</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfMinusPoint' cols='60' rows='4'>" & rsConfig("MessageOfMinusPoint") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>给会员兑换有效期时发送的手机短信内容：</strong><br>不支持HTML代码，可用标签：<br>{$UserName}：会员用户名<br>{$Balance}：资金余额<br>{$ValidDays}：剩余天数<br>{$Money}：支出金额<br>{$Valid}：得到的有效期</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfExchangeValid' cols='60' rows='4'>" & rsConfig("MessageOfExchangeValid") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>给会员奖励有效期时发送的手机短信内容：</strong><br>不支持HTML代码，可用标签：<br>{$UserName}：会员用户名<br>{$ValidDays}：剩余天数<br>{$Valid}：得到的有效期<br>{$Reason}：奖励原因</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfAddValid' cols='60' rows='4'>" & rsConfig("MessageOfAddValid") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='40%' class='tdbg5'><strong>给会员扣除有效期时发送的手机短信内容：</strong><br>不支持HTML代码，可用标签：<br>{$UserName}：会员用户名<br>{$ValidDays}：剩余天数<br>{$Valid}：扣除的有效期<br>{$Reason}：扣除原因</td>" & vbCrLf
    Response.Write "      <td><textarea name='MessageOfMinusValid' cols='60' rows='4'>" & rsConfig("MessageOfMinusValid") & "</textarea></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </tbody>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</td></tr></table>" & vbCrLf
    
    Response.Write "<table width='100%' border='0'>" & vbCrLf
    Response.Write "    <tr>" & vbCrLf
    Response.Write "      <td height='40' align='center'>" & vbCrLf
    Response.Write "        <input name='FileExt_SiteIndex_Old' type='hidden' id='FileExt_SiteIndex_Old' value='" & rsConfig("FileExt_SiteIndex") & "'>" & vbCrLf
    Response.Write "        <input name='FileExt_SiteSpecial_Old' type='hidden' id='FileExt_SiteSpecial_Old' value='" & rsConfig("FileExt_SiteSpecial") & "'>" & vbCrLf
    Response.Write "        <input name='Modules_Old' type='hidden' id='Modules_Old' value='" & rsConfig("Modules") & "'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveConfig'>" & vbCrLf
    Response.Write "        <input name='cmdSave' type='submit' id='cmdSave' value=' 保存设置 '>" & vbCrLf
    Response.Write "      </td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    Response.Write "    <form name='mysitekeyform' id='mysitekeyform' method='POST' action='http://www.powereasy.net/genuine/CheckSite.asp?CheckType=SiteKey' target='_blank'>" & vbCrLf
    Response.Write "  <input type='hidden' id='SiteKey' name='SiteKey' value=''>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    rsConfig.Close
    Set rsConfig = Nothing
End Sub

Sub SaveConfig()
    Dim sqlConfig, rsConfig, iSiteKey, FoundErr
    FoundErr = False

    If Trim(Request("AdminDir")) = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>后台管理目录不能为空</li>"
    End If

    If Trim(Request("ADDir")) = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>网站广告目录不能为空</li>"
    End If

    If PE_CLng(Trim(Request("PhotoObject"))) > 0 Then
        If PE_CLng(Trim(Request("Watermark_Type"))) = 0 Then
            If Trim(Request("Watermark_Text")) <> "" Then
                If Trim(Request("Watermark_Text_FontColor")) = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>定义文字水印的颜色不能为空</li>"
                End If
            End If
        Else
            If Trim(Request("Watermark_Images_FileName")) = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>定义水印图片路径不能为空</li>"
            Else
                If Trim(Request("Watermark_Images_BackgroundColor")) = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>定义去除水印图片背景色不能为空</li>"
                End If
                If Not fso.FileExists(Server.MapPath(Trim(Request("Watermark_Images_FileName")))) Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>定义水印图片中图片路径不对,或指定路径中的图片不存在。</li>"
                End If
            End If
        End If
    End If
    Dim arrLockIP, arrIpW, arrIpB, i, arrLockIPCut
    arrLockIP = Split(Trim(Request("LockIPWhite")), vbCrLf)
    For i = 0 To UBound(arrLockIP)
        If Not (arrLockIP(i) = "" Or IsNull(arrLockIP(i))) And InStr(Trim(arrLockIP(i)), "----") > 0 Then
                arrLockIPCut = Split(Trim(arrLockIP(i)), "----")
                If Not isIP(Trim(arrLockIPCut(0))) Or Not isIP(Trim(arrLockIPCut(1))) Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "请填写正确网站黑白名单中的IP地址！"
                    Exit For
                End If
                If i = 0 Then
                    arrIpW = EncodeIP(Trim(arrLockIPCut(0))) & "----" & EncodeIP(Trim(arrLockIPCut(1)))
                Else
                    arrIpW = arrIpW & "$$$" & EncodeIP(Trim(arrLockIPCut(0))) & "----" & EncodeIP(Trim(arrLockIPCut(1)))
                End If
        End If
    Next
    arrLockIP = Split(Trim(Request("LockIPBlack")), vbCrLf)
    For i = 0 To UBound(arrLockIP)
        If Not (arrLockIP(i) = "" Or IsNull(arrLockIP(i))) And InStr(Trim(arrLockIP(i)), "----") > 0 Then
            arrLockIPCut = Split(Trim(arrLockIP(i)), "----")
            If Not isIP(Trim(arrLockIPCut(0))) Or Not isIP(Trim(arrLockIPCut(1))) Then
                FoundErr = True
                ErrMsg = ErrMsg & "请填写正确网站黑白名单中的IP地址！"
                Exit For
            End If
            If i = 0 Then
                arrIpB = EncodeIP(Trim(arrLockIPCut(0))) & "----" & EncodeIP(Trim(arrLockIPCut(1)))
            Else
                arrIpB = arrIpB & "$$$" & EncodeIP(Trim(arrLockIPCut(0))) & "----" & EncodeIP(Trim(arrLockIPCut(1)))
            End If
        End If
    Next

    If FoundErr = True Then
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    sqlConfig = "select * from PE_Config"
    Set rsConfig = Server.CreateObject("ADODB.Recordset")
    rsConfig.Open sqlConfig, Conn, 1, 3

    If rsConfig.BOF And rsConfig.EOF Then
        rsConfig.addnew
    End If

    rsConfig("SiteName") = Trim(Request("SiteName"))
    rsConfig("SiteTitle") = Trim(Request("SiteTitle"))
    rsConfig("SiteUrl") = Trim(Request("SiteUrl"))
    rsConfig("InstallDir") = InstallDir
    rsConfig("LogoUrl") = Trim(Request("LogoUrl"))
    rsConfig("BannerUrl") = Trim(Request("BannerUrl"))
    rsConfig("WebmasterName") = Trim(Request("WebmasterName"))
    rsConfig("WebmasterEmail") = Trim(Request("WebmasterEmail"))
    rsConfig("Copyright") = Trim(Request("Copyright"))
    rsConfig("Meta_Keywords") = Trim(Request("Meta_Keywords"))
    rsConfig("Meta_Description") = Trim(Request("Meta_Description"))
    rsConfig("SiteKey") = Trim(Request("SaveSiteKey"))

    rsConfig("ShowSiteChannel") = PE_CBool(Trim(Request("ShowSiteChannel")))
    rsConfig("ShowAdminLogin") = PE_CBool(Trim(Request("ShowAdminLogin")))
    rsConfig("EnableSaveRemote") = PE_CBool(Trim(Request("EnableSaveRemote")))
    rsConfig("EnableLinkReg") = PE_CBool(Trim(Request("EnableLinkReg")))
    rsConfig("EnableCountFriendSiteHits") = PE_CBool(Trim(Request("EnableCountFriendSiteHits")))
    rsConfig("EnableSoftKey") = PE_CBool(Trim(Request("EnableSoftKey")))
    rsConfig("IsCustom_Content") = PE_CBool(Trim(Request("IsCustom_Content")))
    rsConfig("objName_FSO") = Trim(Request("objName_FSO"))
    rsConfig("AdminDir") = Trim(Request("AdminDir"))
    rsConfig("ADDir") = Trim(Request("ADDir"))
    rsConfig("AnnounceCookieTime") = PE_CLng(Trim(Request("AnnounceCookieTime")))
    rsConfig("HitsOfHot") = PE_CLng(Trim(Request("HitsOfHot")))
    rsConfig("Modules") = ReplaceBadChar(Trim(Request("Modules")))
    rsConfig("FileExt_SiteIndex") = PE_CLng(Trim(Request("FileExt_SiteIndex")))
    rsConfig("FileExt_SiteSpecial") = PE_CLng(Trim(Request("FileExt_SiteSpecial")))
    rsConfig("SiteUrlType") = PE_CLng(Trim(Request("SiteUrlType")))
    'rsConfig("LockIPType") = PE_CLng(Trim(Request("LockIPW"))) + PE_CLng(Trim(Request("LockIPB")))
    rsConfig("LockIPType") = PE_CLng(Trim(Request("LockIPType")))
    rsConfig("LockIP") = arrIpW & "|||" & arrIpB

    rsConfig("EnableUserReg") = PE_CBool(Trim(Request("EnableUserReg")))
    rsConfig("EmailCheckReg") = PE_CBool(Trim(Request("EmailCheckReg")))
    rsConfig("AdminCheckReg") = PE_CBool(Trim(Request("AdminCheckReg")))
    rsConfig("EnableMultiRegPerEmail") = PE_CBool(Trim(Request("EnableMultiRegPerEmail")))
    rsConfig("EnableCheckCodeOfLogin") = PE_CBool(Trim(Request("EnableCheckCodeOfLogin")))
    rsConfig("EnableCheckCodeOfReg") = PE_CBool(Trim(Request("EnableCheckCodeOfReg")))
    rsConfig("EnableQAofReg") = PE_CBool(Trim(Request("EnableQAofReg")))
    rsConfig("QAofReg") = Trim(Request("RegQuestion1")) & " $$$" & Trim(Request("RegAnswer1")) & " $$$" & Trim(Request("RegQuestion2")) & " $$$" & Trim(Request("RegAnswer2")) & " $$$" & Trim(Request("RegQuestion3")) & " $$$" & Trim(Request("RegAnswer3"))

    rsConfig("UserNameLimit") = PE_CLng(Trim(Request("UserNameLimit")))
    rsConfig("UserNameMax") = PE_CLng(Trim(Request("UserNameMax")))
    rsConfig("UserName_RegDisabled") = Trim(Request("UserName_RegDisabled"))
    rsConfig("RegFields_MustFill") = ReplaceBadChar(Trim(Request("RegFields_MustFill")))

    rsConfig("PresentExp") = PE_CLng(Trim(Request("PresentExp")))
    rsConfig("PresentMoney") = PE_CDbl(Trim(Request("PresentMoney")))
    rsConfig("PresentPoint") = PE_CLng(Trim(Request("PresentPoint")))
    rsConfig("PresentValidNum") = PE_CLng(Trim(Request("PresentValidNum")))
    rsConfig("PresentValidUnit") = PE_CLng(Trim(Request("PresentValidUnit")))
    rsConfig("PresentExpPerLogin") = PE_CLng(Trim(Request("PresentExpPerLogin")))
    rsConfig("MoneyExchangePoint") = PE_CDbl(Trim(Request("MoneyExchangePoint")))
    rsConfig("MoneyExchangeValidDay") = PE_CDbl(Trim(Request("MoneyExchangeValidDay")))
    rsConfig("UserExpExchangePoint") = PE_CDbl(Trim(Request("UserExpExchangePoint")))
    rsConfig("UserExpExchangeValidDay") = PE_CDbl(Trim(Request("UserExpExchangeValidDay")))
    rsConfig("PointName") = Trim(Request("PointName"))
    rsConfig("PointUnit") = Trim(Request("PointUnit"))
    rsConfig("EmailOfRegCheck") = Trim(Request("EmailOfRegCheck"))
    rsConfig("ShowAnonymous") = PE_CBool(Trim(Request("ShowAnonymous")))
	
    rsConfig("MailObject") = Trim(Request("MailObject"))
    rsConfig("MailServer") = Trim(Request("MailServer"))
    rsConfig("MailServerUserName") = Trim(Request("MailServerUserName"))
    rsConfig("MailServerPassWord") = Trim(Request("MailServerPassWord"))
    rsConfig("MailDomain") = Trim(Request("MailDomain"))
    
    rsConfig("PhotoObject") = PE_CLng(Trim(Request("PhotoObject")))
    rsConfig("Thumb_DefaultWidth") = PE_CLng(Trim(Request("Thumb_DefaultWidth")))
    rsConfig("Thumb_DefaultHeight") = PE_CLng(Trim(Request("Thumb_DefaultHeight")))
    rsConfig("Thumb_Arithmetic") = PE_CLng(Trim(Request("Thumb_Arithmetic")))
    rsConfig("Thumb_BackgroundColor") = Trim(Request("Thumb_BackgroundColor"))
    rsConfig("PhotoQuality") = PE_CLng(Trim(Request("PhotoQuality")))

    rsConfig("Watermark_Type") = PE_CLng(Trim(Request("Watermark_Type")))
    rsConfig("Watermark_Text") = Trim(Request("Watermark_Text"))
    rsConfig("Watermark_Text_FontName") = Trim(Request("Watermark_Text_FontName"))
    rsConfig("Watermark_Text_FontSize") = PE_CLng(Trim(Request("Watermark_Text_FontSize")))
    rsConfig("Watermark_Text_FontColor") = Trim(Request("Watermark_Text_FontColor"))
    rsConfig("Watermark_Text_Bold") = PE_CBool(Trim(Request("Watermark_Text_Bold")))
    rsConfig("Watermark_Images_FileName") = Trim(Request("Watermark_Images_FileName"))
    rsConfig("Watermark_Images_Transparence") = PE_CLng(Trim(Request("Watermark_Images_Transparence")))
    rsConfig("Watermark_Images_BackgroundColor") = Trim(Request("Watermark_Images_BackgroundColor"))
    rsConfig("Watermark_Position_X") = PE_CLng(Trim(Request("Watermark_Position_X")))
    rsConfig("Watermark_Position_Y") = PE_CLng(Trim(Request("Watermark_Position_Y")))
    rsConfig("Watermark_Position") = PE_CLng(Trim(Request("Watermark_Position")))
    
    rsConfig("SearchInterval") = PE_CLng(Trim(Request("SearchInterval")))
    rsConfig("SearchResultNum") = PE_CLng(Trim(Request("SearchResultNum")))
    rsConfig("MaxPerPage_SearchResult") = PE_CLng(Trim(Request("MaxPerPage_SearchResult")))
    rsConfig("SearchContent") = PE_CBool(Trim(Request("SearchContent")))
    
    rsConfig("EnableGuestBuy") = PE_CBool(Trim(Request("EnableGuestBuy")))
    rsConfig("IncludeTax") = PE_CBool(Trim(Request("IncludeTax")))
    rsConfig("TaxRate") = PE_CLng(Trim(Request("TaxRate")))

'    rsConfig("PayOnlineProvider") = Trim(Request("PayOnlineProvider"))
'    rsConfig("PayOnlineShopID") = Trim(Request("PayOnlineShopID"))
'    rsConfig("PayOnlineKey") = Trim(Request("PayOnlineKey"))
'    rsConfig("PayOnlineRate") = CDbl(Trim(Request("PayOnlineRate")))
'
'    If Trim(Request("PayOnlinePlusPoundage")) = "1" Then
'        rsConfig("PayOnlinePlusPoundage") = True
'    Else
'        rsConfig("PayOnlinePlusPoundage") = False
'    End If

    rsConfig("Prefix_OrderFormNum") = Trim(Request("Prefix_OrderFormNum"))
    rsConfig("Prefix_PaymentNum") = Trim(Request("Prefix_PaymentNum"))

    rsConfig("Country") = Trim(Request("Country"))
    rsConfig("Province") = Trim(Request("Province"))
    rsConfig("City") = Trim(Request("City"))
    rsConfig("PostCode") = Trim(Request("PostCode"))
    rsConfig("EmailOfOrderConfirm") = Trim(Request("EmailOfOrderConfirm"))
    rsConfig("EmailOfSendCard") = Trim(Request("EmailOfSendCard"))
    rsConfig("EmailOfReceiptMoney") = Trim(Request("EmailOfReceiptMoney"))
    rsConfig("EmailOfRefund") = Trim(Request("EmailOfRefund"))
    rsConfig("EmailOfInvoice") = Trim(Request("EmailOfInvoice"))
    rsConfig("EmailOfDeliver") = Trim(Request("EmailOfDeliver"))
    rsConfig("ShowUserModel") = PE_CBool(Trim(Request("ShowUserModel")))	
    
    rsConfig("GuestBook_EnableVisitor") = PE_CBool(Trim(Request("GuestBook_EnableVisitor")))
    rsConfig("GuestBookCheck") = PE_CBool(Trim(Request("GuestBookCheck")))
    rsConfig("GuestBook_EnableManageRubbish") = PE_CBool(Trim(Request("GuestBook_EnableManageRubbish")))
    Dim arrLockRubbish, arrRubbish
    arrLockRubbish = Split(Trim(Request("LockRubbish")), vbCrLf)
    For i = 0 To UBound(arrLockRubbish)
        If Not (arrLockRubbish(i) = "" Or IsNull(arrLockRubbish(i))) Then
            If i = 0 Then
                arrRubbish = Trim(arrLockRubbish(i))
            Else
                arrRubbish = arrRubbish & "$$$" & Trim(arrLockRubbish(i))
            End If
        End If
    Next
    rsConfig("GuestBook_ManageRubbish") = arrRubbish
    rsConfig("GuestBook_ShowIP") = PE_CBool(Trim(Request("GuestBook_ShowIP")))
    rsConfig("GuestBook_IsAssignSort") = PE_CBool(Trim(Request("GuestBook_IsAssignSort")))
    rsConfig("GuestBook_MaxPerPage") = PE_CLng(Trim(Request("GuestBook_DiscussionMaxPerPage"))) & "|||" & PE_CLng(Trim(Request("GuestBook_GuestBookMaxPerPage"))) & "|||" & PE_CLng(Trim(Request("GuestBook_ReplyMaxPerPage"))) & "|||" & PE_CLng(Trim(Request("GuestBook_TreeMaxPerPage")))

    rsConfig("EnableRss") = PE_CBool(Trim(Request("EnableRss")))
    rsConfig("RssCodeType") = PE_CBool(Trim(Request("RssCodeType")))

    rsConfig("EnableWap") = PE_CBool(Trim(Request("EnableWap")))

    If Trim(Request("WapLogo")) = "" Then
        rsConfig("WapLogo") = 0
    Else
        rsConfig("WapLogo") = Trim(Request("WapLogo"))
    End If

    rsConfig("EnableWapPl") = PE_CBool(Trim(Request("EnableWapPl")))
    rsConfig("ShowWapAppendix") = PE_CBool(Trim(Request("ShowWapAppendix")))
    rsConfig("ShowWapShop") = PE_CBool(Trim(Request("ShowWapShop")))
    rsConfig("ShowWapManage") = PE_CBool(Trim(Request("ShowWapManage")))
    
    rsConfig("SMSUserName") = Trim(Request("SMSUserName"))
    rsConfig("SMSKey") = Trim(Request("SMSKey"))
    rsConfig("SendMessageToAdminWhenOrder") = PE_CBool(Trim(Request("SendMessageToAdminWhenOrder")))
    rsConfig("SendMessageToMemberWhenPaySuccess") = PE_CBool(Trim(Request("SendMessageToMemberWhenPaySuccess")))
    rsConfig("Mobiles") = Trim(Request("Mobiles"))
    rsConfig("MessageOfOrder") = Trim(Request("MessageOfOrder"))
    rsConfig("MessageOfOrderConfirm") = Trim(Request("MessageOfOrderConfirm"))
    rsConfig("MessageOfSendCard") = Trim(Request("MessageOfSendCard"))
    rsConfig("MessageOfReceiptMoney") = Trim(Request("MessageOfReceiptMoney"))
    rsConfig("MessageOfRefund") = Trim(Request("MessageOfRefund"))
    rsConfig("MessageOfInvoice") = Trim(Request("MessageOfInvoice"))
    rsConfig("MessageOfDeliver") = Trim(Request("MessageOfDeliver"))

    rsConfig("MessageOfAddRemit") = Trim(Request("MessageOfAddRemit"))
    rsConfig("MessageOfAddIncome") = Trim(Request("MessageOfAddIncome"))
    rsConfig("MessageOfAddPayment") = Trim(Request("MessageOfAddPayment"))
    rsConfig("MessageOfExchangePoint") = Trim(Request("MessageOfExchangePoint"))
    rsConfig("MessageOfAddPoint") = Trim(Request("MessageOfAddPoint"))
    rsConfig("MessageOfMinusPoint") = Trim(Request("MessageOfMinusPoint"))
    rsConfig("MessageOfExchangeValid") = Trim(Request("MessageOfExchangeValid"))
    rsConfig("MessageOfAddValid") = Trim(Request("MessageOfAddValid"))
    rsConfig("MessageOfMinusValid") = Trim(Request("MessageOfMinusValid"))

    rsConfig.Update
    rsConfig.Close
    Set rsConfig = Nothing
    Dim strSql
    If FoundInArr(Request("Modules"), "Supply", ",") Then
        strSql = "Update PE_Channel Set Disabled=" & PE_False & " Where ModuleType=6"
        Conn.Execute (strSql)
    Else
        strSql = "Update PE_Channel Set Disabled=" & PE_True & " Where ModuleType=6"
        Conn.Execute (strSql)
    End If
    If FoundInArr(Request("Modules"), "Job", ",") Then
        strSql = "Update PE_Channel Set Disabled=" & PE_False & " Where ModuleType=8"
        Conn.Execute (strSql)
    Else
        strSql = "Update PE_Channel Set Disabled=" & PE_True & " Where ModuleType=8"
        Conn.Execute (strSql)
    End If
    If FoundInArr(Request("Modules"), "House", ",") Then
        strSql = "Update PE_Channel Set Disabled=" & PE_False & " Where ModuleType=7"
        Conn.Execute (strSql)
    Else
        strSql = "Update PE_Channel Set Disabled=" & PE_True & " Where ModuleType=7"
        Conn.Execute (strSql)
    End If
    
    Call WriteSuccessMsg("网站配置保存成功！", ComeUrl)

    Dim FileExt_SiteIndex, FileExt_SiteIndex_Old, FileExt_SiteSpecial, FileExt_SiteSpecial_Old
    FileExt_SiteIndex = PE_CLng(Trim(Request("FileExt_SiteIndex")))
    FileExt_SiteIndex_Old = PE_CLng(Trim(Request("FileExt_SiteIndex_Old")))
    FileExt_SiteSpecial = PE_CLng(Trim(Request("FileExt_SiteSpecial")))
    FileExt_SiteSpecial_Old = PE_CLng(Trim(Request("FileExt_SiteSpecial_Old")))
    
    If IsReload(FileExt_SiteIndex, FileExt_SiteIndex_Old) Or IsReload(FileExt_SiteSpecial, FileExt_SiteSpecial_Old) Or Trim(Request("Modules")) <> Trim(Request("Modules_Old")) Then
        Call ReloadLeft
    End If

End Sub

Function IsReload(FileExt, FileExt_Old)
    IsReload = False
    If FileExt <> FileExt_Old Then
        If FileExt = 4 Or FileExt_Old = 4 Then
            IsReload = True
        End If
    End If
End Function

Sub ReloadLeft()
    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "  parent.left.location.reload();" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub

Function ShowInstalled(strObject)
    If Not IsObjInstalled(strObject) Then
        ShowInstalled = "<font color='red'><b>×</b></font>"
    Else
        ShowInstalled = "<b>√</b>"
    End If
End Function

Function IsModulesSelected(Compare1, Compare2)
    If FoundInArr(Compare1, Compare2, ",") = True Then
        IsModulesSelected = " checked"
    Else
        IsModulesSelected = ""
    End If
End Function

Function IsRadioChecked(Compare1, Compare2)
    If Compare1 = Compare2 Then
        IsRadioChecked = " checked"
    Else
        IsRadioChecked = ""
    End If
End Function

Function IsOptionSelected(Compare1, Compare2)
    If Compare1 = Compare2 Then
        IsOptionSelected = " selected"
    Else
        IsOptionSelected = ""
    End If
End Function

Function IsMustFill(Compare1, Compare2)
    If FoundInArr(Compare1, Compare2, ",") = True Then
        IsMustFill = " checked"
    Else
        IsMustFill = ""
    End If
End Function
Function GetProvince(ProvinceName)
    Dim rsProvince, strProvince
    strProvince = "<option value=''>请选择省份</option>"
    Set rsProvince = Conn.Execute("select DISTINCT Province from PE_City")
    Do While Not rsProvince.EOF
        If rsProvince(0) = ProvinceName Then
            strProvince = strProvince & "<option value='" & rsProvince(0) & "' selected>" & rsProvince(0) & "</option>"
        Else
            strProvince = strProvince & "<option value='" & rsProvince(0) & "'>" & rsProvince(0) & "</option>"
        End If
        rsProvince.MoveNext
    Loop
    Set rsProvince = Nothing
    GetProvince = strProvince
End Function

Function ISdisplay(ByVal Compare1, ByVal Compare2)
    If Compare1 = Compare2 Then
        ISdisplay = " style='display:'"
    Else
        ISdisplay = " style='display:none'"
    End If
End Function
%>
