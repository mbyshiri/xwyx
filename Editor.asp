<!-- #include File="Start.asp" -->
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

Dim arrButtons(104), arrButtons2, strButtons, arrButtonOption, i, TemplateType, EditorContent, tContentID
Dim ChannelID,ShowType,rs
Dim Anonymous

'获取频道相关数据
ChannelID = PE_CLng(Trim(Request("ChannelID")))
tContentID = FilterJS(Request("tContentID"))
TemplateType = Trim(Request("TemplateType"))
Anonymous = PE_CLng(Request("Anonymous"))
If TemplateType = "" Then
    TemplateType = 1
Else
    TemplateType = PE_CLng(TemplateType)
End If

ShowType = PE_CLng(Trim(Request("ShowType")))

'调入按钮数组
arrButtons(0) = "yToolbar$$$"
arrButtons(1) = "/yToolbar$$$"
arrButtons(2) = "TBHandle$$$"
arrButtons(3) = "TBSep$$$"
arrButtons(101) = "TBGen$$$"
arrButtons(102) = "TBGen2$$$"
arrButtons(103) = "TBGen3$$$"
arrButtons(5) = "Btn$全部选择$format('selectall')$selectall.gif"
arrButtons(6) = "Btn$删除$format('delete')$delete.gif"
arrButtons(7) = "Btn$剪切$format('cut')$cut.gif"
arrButtons(8) = "Btn$复制$format('copy')$copy.gif"
arrButtons(9) = "Btn$粘贴$format('paste')$paste.gif"
arrButtons(10) = "Btn$从word中粘贴$insert('word')$wordpaste.gif"
arrButtons(11) = "Btn$撤消$format('undo')$undo.gif"
arrButtons(12) = "Btn$恢复$format('redo')$redo.gif"
arrButtons(13) = "Btn$查找 / 替换$findstr()$find.gif"
arrButtons(14) = "Btn$计算器$insert('calculator')$calculator.gif"
arrButtons(15) = "Btn$打印$format('Print')$print.gif"
arrButtons(16) = "Btn$查看帮助$insert('help')$help.gif"
arrButtons(17) = "Btn$左对齐$format('justifyleft')$aleft.gif"
arrButtons(18) = "Btn$居中$format('justifycenter')$acenter.gif"
arrButtons(19) = "Btn$右对齐$format('justifyright')$aright.gif"
arrButtons(20) = "Btn$两端对齐$format('JustifyFull')$JustifyFull.gif"
arrButtons(21) = "Btn$绝对或相对位置$format('absolutePosition')$abspos.gif"
arrButtons(22) = "Btn$删除文字格式$format('RemoveFormat')$clear.gif"
arrButtons(23) = "Btn$插入段落$format('insertparagraph')$paragraph.gif"
arrButtons(24) = "Btn$插入换行符号$insert('br')$chars.gif"
arrButtons(25) = "Btn$字体颜色$insert('fgcolor')$fgcolor.gif"
arrButtons(26) = "Btn$文字背景色$insert('fgbgcolor')$fgbgcolor.gif"
arrButtons(27) = "Btn$加粗$format('bold')$bold.gif"
arrButtons(28) = "Btn$斜体$format('italic')$italic.gif"
arrButtons(29) = "Btn$下划线$format('underline')$underline.gif"
arrButtons(30) = "Btn$删除线$format('StrikeThrough')$strikethrough.gif"
arrButtons(31) = "BtnMenu$更多文字格式$showToolMenu('font')$arrow.gif"
arrButtons(32) = "Btn$显示或隐藏表格虚线、按钮等显示样式$showBorders()$showBorders.gif"
arrButtons(33) = "Btn$图片左环绕$imgalign('left')$imgleft.gif"
arrButtons(34) = "Btn$图片右环绕$imgalign('right')$imgright.gif"
arrButtons(35) = "Btn$插入超级连接$insert('CreateLink')$url.gif"
arrButtons(36) = "Btn$取消超级链接$format('unLink')$nourl.gif"
arrButtons(37) = "Btn$插入普通水平线$format('InsertHorizontalRule')$line.gif"
arrButtons(38) = "Btn$插入特殊水平线$insert('hr')$sline.gif"
arrButtons(39) = "Btn$插入手动分页符$insert('page')$page.gif"
arrButtons(40) = "Btn$插入当前日期$insert('nowdate')$date.gif"
arrButtons(41) = "Btn$插入当前时间$insert('nowtime')$time.gif"
arrButtons(42) = "Btn$插入栏目框$insert('FIELDSET')$fieldset.gif"
arrButtons(43) = "Btn$插入网页$insert('iframe')$htm.gif"
arrButtons(44) = "Btn$插入Excel表格$insert('excel')$excel.gif"
arrButtons(45) = "Btn$插入表格$TableInsert()$table.gif"
arrButtons(46) = "BtnMenu$表格操作$showToolMenu('table')$arrow.gif"
arrButtons(47) = "Btn$插入下拉菜单$Insermenu('" & Now() & "')$menu.gif"
arrButtons(48) = "BtnMenu$更多表单控件$showToolMenu('form')$arrow.gif"
arrButtons(49) = "Btn$插入滚动文本$insert('insermarquee')$Marquee.gif"
arrButtons(50) = "BtnMenu$插入更多样式$showToolMenu('object')$arrow.gif"
arrButtons(51) = "Btn$插入表情符号$insert('inseremot')$Emot.gif"
arrButtons(52) = "Btn$插入特殊符号$Insertlr('editor_tsfh.asp',300,190," & (Now() - Date) * 24 * 60 * 60 * 1000 & ")$symbol.gif"
'arrButtons(53) = "Btn$插入公式$insert('InsertEQ')$eq.gif"
arrButtons(53) = "Btn$内容过滤$insert('FilterCode')$FilterCode.gif"
arrButtons(54) = "BtnMenu$公式操作$showToolMenu('gongshi')$arrow.gif"
arrButtons(55) = "Btn$插入图片，支持格式为：jpg、gif、bmp、png等$insert('pic')$img.gif"
arrButtons(56) = "Btn$批量上传图片，支持格式为：jpg、gif、bmp、png等$insert('batchpic')$pimg.gif"
arrButtons(57) = "Btn$插入flash多媒体文件$insert('swf')$flash.gif"
arrButtons(58) = "Btn$插入视频文件，支持格式为：avi、wmv、asf等$insert('wmv')$wmv.gif"
arrButtons(59) = "Btn$插入RealPlay文件，支持格式为：rm、ra、ram$insert('rm')$rm.gif"
arrButtons(60) = "Btn$上传附件$insert('fujian')$fujian.gif"
arrButtons(61) = "Btn$从上传文件中选择$insert('SelectUpFile')$SelectUpFile.gif"
arrButtons(62) = "Btn$插入标签$insert('Label')$label.gif"
arrButtons(63) = "Btn$图片单行居中$imgalign('center')$imgcenter.gif"
arrButtons(64) = "Btn$插入带标题的分页$insert('pagetitle')$pagetitle.gif"

arrButtons(65) = "Btn$显示文章标题等信息$SuperFunctionLabel('" & InstallDir & "Editor/editor_label.asp','GetArticleList','文章列表函数标签',1,'GetList',800,700)$LabelIco\GetArticleList.gif"
arrButtons(66) = "Btn$显示图片文章$SuperFunctionLabel('" & InstallDir & "Editor/editor_label.asp','GetPicArticle','显示图片文章标签',1,'GetPic',700,500)$LabelIco\GetPicArticle.gif"
arrButtons(67) = "Btn$显示幻灯片文章$SuperFunctionLabel('" & InstallDir &"Editor/editor_label.asp','GetSlidePicArticle','显示幻灯片文章标签',1,'GetSlide',700,500)$LabelIco\GetSlidePicArticle.gif"
arrButtons(68) = "Btn$文章自定义列表$SuperFunctionLabel('" & InstallDir &"Editor/editor_CustomListLabel.asp','CustomListLable','文章自定义列表标签',1,'GetArticleCustom',720,700)$LabelIco\GetArticleCustom.gif"
arrButtons(69) = "Btn$显示软件标题$SuperFunctionLabel('" & InstallDir &"Editor/editor_label.asp','GetSoftList','下载列表函数标签',2,'GetList',800,700)$LabelIco\GetSoftList.gif"
arrButtons(70) = "Btn$显示图片下载$SuperFunctionLabel('" & InstallDir &"Editor/editor_label.asp','GetPicSoft','显示图片下载标签',2,'GetPic',700,500)$LabelIco\GetPicSoft.gif"
arrButtons(71) = "Btn$显示幻灯片下载$SuperFunctionLabel('" & InstallDir &"Editor/editor_label.asp','GetSlidePicSoft','显示幻灯片下载标签',2,'GetSlide',700,500)$LabelIco\GetSlidePicSoft.gif"
arrButtons(72) = "Btn$下载自定义列表$SuperFunctionLabel('" & InstallDir &"Editor/editor_CustomListLabel.asp','CustomListLable','下载自定义列表标签',2,'GetSoftCustom',720,700)$LabelIco\GetSoftCustom.gif"
arrButtons(73) = "Btn$显示图片标题$SuperFunctionLabel('" & InstallDir &"Editor/editor_label.asp','GetPhotoList','图片列表函数标签',3,'GetList',800,700)$LabelIco\GetPhotoList.gif"
arrButtons(74) = "Btn$显示图片$SuperFunctionLabel('" & InstallDir &"Editor/editor_label.asp','GetPicPhoto','显示图片图文标签',3,'GetPic',700,550)$LabelIco\GetPicPhoto.gif"
arrButtons(75) = "Btn$显示图片幻灯片$SuperFunctionLabel('" & InstallDir &"Editor/editor_label.asp','GetSlidePicPhoto','显示幻灯片图片标签',3,'GetSlide',700,550)$LabelIco\GetSlidePicPhoto.gif"
arrButtons(77) = "Btn$显示商品标题$SuperFunctionLabel('" & InstallDir &"Editor/editor_label.asp','GetProductList','商城列表函数标签',5,'GetList',800,750)$LabelIco\GetProductList.gif"
arrButtons(78) = "Btn$显示商品图片$SuperFunctionLabel('" & InstallDir &"Editor/editor_label.asp','GetPicProduct','显示图片商城标签',5,'GetPic',700,600)$LabelIco\GetPicProduct.gif"
arrButtons(79) = "Btn$显示商品幻灯片$SuperFunctionLabel('" & InstallDir &"Editor/editor_label.asp','GetSlidePicProduct','显示幻灯片商城标签',5,'GetSlide',700,460)$LabelIco\GetSlidePicProduct.gif"
arrButtons(80) = "Btn$商品自定义列表$SuperFunctionLabel('" & InstallDir &"Editor/editor_CustomListLabel.asp','CustomListLable','商城自定义列表标签',5,'GetProductCustom',720,700)$LabelIco\GetProductCustom.gif"
arrButtons(81) = "Btn$网站logo$FunctionLabel('" & InstallDir &"Editor/Lable/PE_Logo.htm','240','140')$LabelIco\PE_logo.gif"
arrButtons(82) = "Btn$网站banner$FunctionLabel('" & InstallDir &"Editor/Lable/PE_Banner.htm','240','140')$LabelIco\PE_banner.gif"
arrButtons(83) = "Btn$弹出公告$FunctionLabel('" & InstallDir &"Editor/Lable/PE_AnnWin.htm','240','140')$LabelIco\PE_PopAnnouce.gif"
arrButtons(84) = "Btn$公告$FunctionLabel('" & InstallDir &"Editor/Lable/PE_Annouce2.htm','240','210')$LabelIco\PE_Annouce.gif"
arrButtons(85) = "Btn$友情$FunctionLabel('" & InstallDir &"Editor/Lable/PE_FSite2.htm','330','510')$LabelIco\PE_FriendSite.gif"
arrButtons(86) = "Btn$调查$InsertLabel('ShowVote')$LabelIco\PE_Vote.gif"
arrButtons(87) = "Btn$作者列表$FunctionLabel('" & InstallDir &"Editor/Lable/PE_Author_ShowList.htm','400','345')$LabelIco\PE_Author.gif"
arrButtons(88) = "Btn$厂商列表$FunctionLabel('" & InstallDir &"Editor/Lable/PE_ProducerList.htm','400','450')$LabelIco\PE_Producer.gif"
arrButtons(89) = "Btn$显示作品集排行$FunctionLabel('" & InstallDir &"Editor/Lable/PE_ShowBlogList.htm','400','400')$LabelIco\PE_Blog.gif"
arrButtons(90) = "Btn$显示专题列表$FunctionLabel('" & InstallDir &"Editor/Lable/PE_ShowSpecialList.htm','320','300')$LabelIco\PE_Special.gif"
arrButtons(91) = "Btn$显示注册用户$InsertLabel('ShowTopUser')$LabelIco\PE_user.gif"
arrButtons(92) = "Btn$登录$InsertLabel('ShowAdminLogin')$LabelIco\PE_AdminLogin.gif"
arrButtons(93) = "Btn$导航$InsertLabel('ShowPath')$LabelIco\PE_Path.gif"
arrButtons(94) = "Btn$版权$InsertLabel('Copyright')$LabelIco\PE_Copyright.gif"

Response.Write "<html>"
Response.Write "<head>"
Response.Write " <meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
Response.Write " <title>HTML在线编辑器</title>"
Response.Write " <link rel='STYLESHEET' type='text/css' href='Editor/editor.css'>"
Response.Write "</head>"
Response.Write "<body bgcolor='#FFFFFF' leftmargin='0' topmargin='0' onConTextMenu='event.returnValue=false;'>"
Response.Write "    <table border='0' cellpadding='0' cellspacing='0' width='100%' height='100%' align='center'>"
Response.Write "      <tr>"
Response.Write "       <td valign='top'>"

Select Case ShowType
Case 0   '文章
    strButtons = arrButtons(0) & "|" & arrButtons(2) & "|" & arrButtons(5) & "|" & arrButtons(3) & "|" & arrButtons(6) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(7) & "|" & arrButtons(8) & "|" & arrButtons(9) & "|"
    strButtons = strButtons & arrButtons(10) & "|" & arrButtons(3) & "|" & arrButtons(11) & "|" & arrButtons(12) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(13) & "|" & arrButtons(3) & "|" & arrButtons(14) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(15) & "|" & arrButtons(3) & "|" & arrButtons(16) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(17) & "|" & arrButtons(18) & "|" & arrButtons(19) & "|"
    strButtons = strButtons & arrButtons(20) & "|" & arrButtons(21) & "|" & arrButtons(3) & "|" & arrButtons(22) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(35) & "|" & arrButtons(36) &"|"  & arrButtons(1) & "|"
    strButtons = strButtons & arrButtons(0) & "|" & arrButtons(2) & "|" & arrButtons(101) & "|"  & arrButtons(102) & "|" & arrButtons(103) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(25) & "|" & arrButtons(26) & "|" & arrButtons(3) & "|" & arrButtons(27) & "|"
    strButtons = strButtons & arrButtons(28) & "|" & arrButtons(29) & "|" & arrButtons(30) & "|" & arrButtons(31) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(32) & "|" & arrButtons(3) & "|" & arrButtons(33) & "|"
    strButtons = strButtons & arrButtons(63) & "|" & arrButtons(34) & "|" & arrButtons(3) & "|" & arrButtons(24) & "|"
    strButtons = strButtons & arrButtons(1) & "|" & arrButtons(0) & "|" & arrButtons(2) & "|" & arrButtons(37) & "|"
    strButtons = strButtons & arrButtons(38) & "|" & arrButtons(3) & "|" & arrButtons(39) &"|" & "|" & arrButtons(64) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(41) & "|" & arrButtons(40) & "|" & arrButtons(43) & "|" & arrButtons(42) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(45) & "|" & arrButtons(46) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(47) & "|" & arrButtons(48) & "|" & arrButtons(49) & "|" & arrButtons(50) & "|"
    strButtons = strButtons & arrButtons(51) & "|" & arrButtons(3) & "|" & arrButtons(52) & "|" & arrButtons(53) & "|"
    strButtons = strButtons & arrButtons(54) & "|" & arrButtons(3) & "|" & arrButtons(55) & "|" & arrButtons(56) & "|"
    strButtons = strButtons & arrButtons(57) & "|" & arrButtons(58) & "|" & arrButtons(59) & "|" & arrButtons(60) & "|"
    strButtons = strButtons & arrButtons(1)
Case 1   '模板
    strButtons = arrButtons(0) & "|" & arrButtons(2) & "|" & arrButtons(65) & "|" & arrButtons(66) & "|" & arrButtons(67) & "|"
    strButtons = strButtons & arrButtons(68) & "|" & arrButtons(3) & "|" & arrButtons(69) & "|" & arrButtons(70) & "|"
    strButtons = strButtons & arrButtons(71) & "|" & arrButtons(72) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(73) & "|" & arrButtons(74) & "|" & arrButtons(75) & "|" & arrButtons(3) & "|" & arrButtons(77) & "|"
    strButtons = strButtons & arrButtons(78) & "|" & arrButtons(79) & "|" & arrButtons(80) & "|"& arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(81) & "|" & arrButtons(82) & "|" & arrButtons(3) & "|" & arrButtons(83) & "|"
    strButtons = strButtons & arrButtons(84) & "|" & arrButtons(85) & "|" & arrButtons(86) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(87) & "|" & arrButtons(88) & "|" & arrButtons(89) & "|" & arrButtons(90) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(91) & "|" & arrButtons(92) & "|" & arrButtons(93) & "|" & arrButtons(94) & "|"
    strButtons = strButtons & arrButtons(1) & "|" & arrButtons(0) & "|" & arrButtons(2) & "|" & arrButtons(5) & "|" & arrButtons(3) & "|" & arrButtons(6) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(7) & "|" & arrButtons(8) & "|" & arrButtons(9) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(11) & "|" & arrButtons(12) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(13) & "|" & arrButtons(3) & "|" & arrButtons(14) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(15) & "|" & arrButtons(3) & "|" & arrButtons(16) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(17) & "|" & arrButtons(18) & "|" & arrButtons(19) & "|" & arrButtons(20) & "|"
    strButtons = strButtons & arrButtons(21) & "|" & arrButtons(3) & "|" & arrButtons(33) & "|" & arrButtons(63) & "|"
    strButtons = strButtons & arrButtons(34) & "|" & arrButtons(3) & "|" & arrButtons(22) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(37) & "|" & arrButtons(38) & "|" & arrButtons(40) & "|" & arrButtons(41) & "|"
    strButtons = strButtons & arrButtons(52) & "|" & arrButtons(35) & "|" & arrButtons(36) & "|" & arrButtons(24) & "|"
    strButtons = strButtons & arrButtons(1) & "|" & arrButtons(0) & "|" & arrButtons(2) & "|" & arrButtons(101) & "|"  & arrButtons(102) & "|" & arrButtons(103) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(25) & "|" & arrButtons(26) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(27) & "|" & arrButtons(28) & "|" & arrButtons(29) & "|" & arrButtons(30) & "|"
    strButtons = strButtons & arrButtons(31) & "|" & arrButtons(3) & "|" & arrButtons(62) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(43) & "|" & arrButtons(45) & "|" & arrButtons(46) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(32) & "|" & arrButtons(47) & "|" & arrButtons(48) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(49) & "|" & arrButtons(50) & "|" & arrButtons(55) & "|" & arrButtons(57) & "|"
    strButtons = strButtons & arrButtons(58) & "|" & arrButtons(59) & "|" & arrButtons(60) & "|" & arrButtons(1)
Case 2   '留言 公告
    strButtons = strButtons & arrButtons(0) & "|" & arrButtons(2) & "|" & arrButtons(101) & "|"  & arrButtons(102) & "|" & arrButtons(103) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(25) & "|" & arrButtons(26) & "|" & arrButtons(3) & "|" & arrButtons(27) & "|"
    strButtons = strButtons & arrButtons(28) & "|" & arrButtons(29) & "|" & arrButtons(30) & "|" & arrButtons(31) & "|"
    strButtons = strButtons & arrButtons(16) & "|" & arrButtons(1) & "|"
    strButtons = strButtons & arrButtons(0) & "|" & arrButtons(2) & "|" & arrButtons(17) & "|" & arrButtons(18) & "|"
    strButtons = strButtons & arrButtons(19) & "|" & arrButtons(3) & "|" & arrButtons(22) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(35) & "|" & arrButtons(36) & "|" & arrButtons(43) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(45) & "|" & arrButtons(46) & "|" & arrButtons(49) & "|"
    strButtons = strButtons & arrButtons(50) & "|" & arrButtons(3) & "|" & arrButtons(51) & "|" & arrButtons(52) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(55) & "|" & arrButtons(57) & "|" & arrButtons(58) & "|"
    strButtons = strButtons & arrButtons(59) & "|" & arrButtons(60) & "|" & arrButtons(1)
Case 3   '说明框
    strButtons = strButtons & arrButtons(0) & "|" & arrButtons(2) & "|" & arrButtons(101) & "|"  & arrButtons(102) & "|" & arrButtons(103) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(25) & "|" & arrButtons(26) & "|" & arrButtons(3) & "|" & arrButtons(27) & "|"
    strButtons = strButtons & arrButtons(28) & "|" & arrButtons(29) & "|" & arrButtons(30) & "|" & arrButtons(31) & "|"
    strButtons = strButtons & arrButtons(22) & "|" & arrButtons(35) & "|" & arrButtons(36) & "|" & arrButtons(52) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(55) & "|" & arrButtons(57) & "|" & arrButtons(58) & "|"
    strButtons = strButtons & arrButtons(59) & "|" & arrButtons(60) & "|" & arrButtons(1)
Case 4  '商城
    strButtons = arrButtons(0) & "|" & arrButtons(2) & "|" & arrButtons(5) & "|" & arrButtons(3) & "|" & arrButtons(6) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(7) & "|" & arrButtons(8) & "|" & arrButtons(9) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(11) & "|" & arrButtons(12) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(13) & "|" & arrButtons(14) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(17) & "|" & arrButtons(18) & "|" & arrButtons(19) & "|" & arrButtons(20) & "|"
    strButtons = strButtons & arrButtons(21) & "|" & arrButtons(3) & "|" & arrButtons(33) & "|" & arrButtons(63) & "|"
    strButtons = strButtons & arrButtons(34) & "|" & arrButtons(3) & "|" & arrButtons(22) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(38) & "|" & arrButtons(40) & "|" & arrButtons(41) & "|"
    strButtons = strButtons & arrButtons(52) & "|" & arrButtons(35) & "|" & arrButtons(36) & "|" & arrButtons(24) & "|"
    strButtons = strButtons & arrButtons(1) & "|" & arrButtons(0) & "|" & arrButtons(2) & "|" & arrButtons(101) & "|"  & arrButtons(102) & "|" & arrButtons(103) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(25) & "|" & arrButtons(26) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(27) & "|" & arrButtons(28) & "|" & arrButtons(29) & "|" 
    strButtons = strButtons & arrButtons(31) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(45) & "|" & arrButtons(46) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(47) & "|" & arrButtons(48) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(55) & "|" & arrButtons(56) & "|" & arrButtons(57) & "|"
    strButtons = strButtons & arrButtons(58) & "|" & arrButtons(59) & "|" & arrButtons(60) & "|" & arrButtons(1)
Case 5 '供求模块 
    strButtons = strButtons & arrButtons(0) & "|" & arrButtons(2) & "|" & arrButtons(101) & "|"  & arrButtons(102) & "|" & arrButtons(103) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(25) & "|" & arrButtons(26) & "|" & arrButtons(3) & "|" & arrButtons(27) & "|"
    strButtons = strButtons & arrButtons(28) & "|" & arrButtons(29) & "|" & arrButtons(30) & "|" & arrButtons(31) & "|"
    strButtons = strButtons & arrButtons(16) & "|" & arrButtons(1) & "|"
    strButtons = strButtons & arrButtons(0) & "|" & arrButtons(2) & "|" & arrButtons(17) & "|" & arrButtons(18) & "|"
    strButtons = strButtons & arrButtons(19) & "|" & arrButtons(3) & "|" & arrButtons(22) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(35) & "|" & arrButtons(36) & "|" & arrButtons(43) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(45) & "|" & arrButtons(46) & "|" & arrButtons(49) & "|"
    strButtons = strButtons & arrButtons(50) & "|" & arrButtons(3) & "|" & arrButtons(51) & "|" & arrButtons(52) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(55) & "|" & arrButtons(57) & "|" & arrButtons(58) & "|"
    strButtons = strButtons & arrButtons(59) & "|" & arrButtons(60) & "|" & arrButtons(1)
Case 6   '说明框
    strButtons = strButtons & arrButtons(0) & "|" & arrButtons(2) & "|"  & arrButtons(102) & "|" & "|" & arrButtons(103) & "|" & arrButtons(3) & "|"
    strButtons = strButtons & arrButtons(25) & "|" & arrButtons(26) & "|" & arrButtons(3) & "|" & arrButtons(27) & "|"
    strButtons = strButtons & arrButtons(28) & "|" & arrButtons(29) & "|" & arrButtons(30) & "|" & arrButtons(31) & "|"
    strButtons = strButtons & arrButtons(22) & "|" & arrButtons(35) & "|" & arrButtons(36) & "|" & arrButtons(52) & "|"
    strButtons = strButtons & arrButtons(3) & "|" & arrButtons(55) & "|" & arrButtons(1)
End Select

arrButtons2 = Split(strButtons, "|")

For i = 0 To UBound(arrButtons2)
    If arrButtons2(i) <> "" Then
        arrButtonOption = Split(arrButtons2(i), "$")
        Select Case arrButtonOption(0)
        Case "yToolbar"
            Response.Write "<div class='yToolbar'>" & vbCrLf
        Case "/yToolbar"
            Response.Write "</div>" & vbCrLf
        Case "TBHandle"
            Response.Write "  <div class='TBHandle'></div>" & vbCrLf
        Case "Btn"
            Response.Write "  <div class='Btn' TITLE='" & arrButtonOption(1) & "' LANGUAGE='javascript' onclick=""" & arrButtonOption(2) & """><img class='Ico' src='editor/images/" & arrButtonOption(3) & "' WIDTH='18' HEIGHT='18'></div>" & vbCrLf
        Case "BtnMenu"
            Response.Write "  <div class='BtnMenu' TITLE='" & arrButtonOption(1) & "' LANGUAGE='javascript' onclick=""" & arrButtonOption(2) & """><img class='Ico' src='editor/images/" & arrButtonOption(3) & "' WIDTH='5' HEIGHT='18'></div>" & vbCrLf
        Case "TBSep"
            Response.Write "  <div class='TBSep'></div>" & vbCrLf
        Case "TBGen"
            Response.Write "<select ID=""formatSelect"" class=""TBGen"" "
            Response.Write "onchange=""format('FormatBlock',this[this.selectedIndex].value);this.selectedIndex=0"">"
            Response.Write "  <option selected>段落格式</option>"
            Response.Write "  <option VALUE=""&lt;P&gt;"">普通</option>"
            Response.Write "  <option VALUE=""&lt;PRE&gt;"">已编排格式</option>"
            Response.Write "  <option VALUE=""&lt;H1&gt;"">标题一</option>"
            Response.Write "  <option VALUE=""&lt;H2&gt;"">标题二</option>"
            Response.Write "  <option VALUE=""&lt;H3&gt;"">标题三</option>"
            Response.Write "  <option VALUE=""&lt;H4&gt;"">标题四</option>"
            Response.Write "  <option VALUE=""&lt;H5&gt;"">标题五</option>"
            Response.Write "  <option VALUE=""&lt;H6&gt;"">标题六</option>"
            Response.Write "  <option VALUE=""&lt;H7&gt;"">标题七</option>"
            Response.Write "</select>"
        Case "TBGen2"
            Response.Write "<select id=""FontName"" class=""TBGen"" onchange=""format('fontname',this[this.selectedIndex].value);this.selectedIndex=0"">"
            Response.Write "  <option selected>字体</option>"
            Response.Write "  <option value=""宋体"">宋体</option>"
            Response.Write "  <option value=""黑体"">黑体</option>"
            Response.Write "  <option value=""楷体_GB2312"">楷体</option>"
            Response.Write "  <option value=""仿宋_GB2312"">仿宋</option>"
            Response.Write "  <option value=""隶书"">隶书</option>"
            Response.Write "  <option value=""幼圆"">幼圆</option>"
            Response.Write "  <option value=""Arial"">Arial</option>"
            Response.Write "  <option value=""Arial Black"">Arial Black</option>"
            Response.Write "  <option value=""Arial Narrow"">Arial Narrow</option>"
            Response.Write "  <option value=""Brush ScriptMT"">Brush Script MT</option>"
            Response.Write "  <option value=""Century Gothic"">Century Gothic</option>"
            Response.Write "  <option value=""Comic Sans MS"">Comic Sans MS</option>"
            Response.Write "  <option value=""Courier"">Courier</option>"
            Response.Write "  <option value=""Courier New"">Courier New</option>"
            Response.Write "  <option value=""MS Sans Serif"">MS Sans Serif</option>"
            Response.Write "  <option value=""Script"">Script</option>"
            Response.Write "  <option value=""System"">System</option>"
            Response.Write "  <option value=""Times New Roman"">Times New Roman</option>"
            Response.Write "  <option value=""Verdana"">Verdana</option>"
            Response.Write "  <option value=""WideLatin"">Wide Latin</option>"
            Response.Write "  <option value=""Wingdings"">Wingdings</option>"
            Response.Write "</select>"
        Case "TBGen3"
            Response.Write "<select id=""FontSize"" class=""TBGen"" onchange=""format('fontsize',this[this.selectedIndex].value);this.selectedIndex=0"">"
            Response.Write "  <option selected>字号</option>"
            Response.Write "  <option value=""7"">一号</option>"
            Response.Write "  <option value=""6"">二号</option>"
            Response.Write "  <option value=""5"">三号</option>"
            Response.Write "  <option value=""4"">四号</option>"
            Response.Write "  <option value=""3"">五号</option>"
            Response.Write "  <option value=""2"">六号</option>"
            Response.Write "  <option value=""1"">七号</option>"
            Response.Write "</select>"
        End Select
    End If
Next

Response.Write "</td></tr>"
Response.Write "  <tr>"
Response.Write "   <td valign='top' height='100%'>"
Response.Write "     <table border=0 cellpadding=0 cellspacing=0 width='100%' height='100%'>"
Response.Write "     <tr><td height='100%'>"
Response.Write "       <iframe style='font-size:12px' ID='HtmlEdit'  MARGINHEIGHT='1' MARGINWIDTH='1' style='width=100%; height=100%;' scrolling='yes' ></iframe>"
Response.Write "     </td></tr>"
Response.Write "     </table>"
Response.Write "   </td>"
Response.Write "  </tr>"
Response.Write "  <tr>"
Response.Write "   <td valign='top' height='25'>"
Response.Write "     <table border='0' cellpadding='0' cellspacing='0' width='100%' height='20' align='center'>"
Response.Write "      <tr>"
If ShowType <> 1 Then
    Response.Write "       <td valign='top' width='265' >"
    Response.Write "         <img id=setMode0 src='Editor/images/Editor2.gif' width='59' height='20' onclick=""setMode('EDIT')"">"
    Response.Write "         <img id=setMode1 src='Editor/images/html.gif' width='59' height='20' onclick=""setMode('CODE')"">"
    Response.Write "         <img id=setMode2 src='Editor/images/browse.gif' width='59' height='20' onclick=""setMode('VIEW')"">"
    Response.Write "         <img id=setMode3 src='Editor/images/text.gif' width='59' height='20' onclick=""setMode('TEXT')"">"
    Response.Write "       </td>" 
    Response.Write "       <td width='20' align='left'>"
    Response.Write "       <select name='Zoomname' id='doZoomid' onchange='doZoom(this[this.selectedIndex].value)'>"
    Response.Write "         <option value='10'>10%</option>"
    Response.Write "         <option value='25'>25%</option>"
    Response.Write "         <option value='50'>50%</option>"
    Response.Write "         <option value='75'>75%</option>"
    Response.Write "         <option value='100' selected>100%</option>"
    Response.Write "         <option value='150'>150%</option>"
    Response.Write "         <option value='200'>200%</option>"
    Response.Write "         <option value='500'>500%</option>"
    Response.Write "       </select>"
    Response.Write "       </td>"
Else
    Response.Write "       <td id='ShowObject' width='90%'></td>"
End If
Response.Write "       <td valign='top' align='right'>"
Response.Write "         <img  src='Editor/images/sizeplus.gif' width='20' height='20' onclick='sizeChange(200)'>"
Response.Write "         <img  src='Editor/images/sizeminus.gif' width='20' height='20' onclick='sizeChange(-200)'>"
Response.Write "       </td>"
Response.Write "       <td width='30'></td>"
Response.Write "     </tr>"
Response.Write "     </table>"
Response.Write "        <div id='HtmlEdit_Temp_HTML' style='VISIBILITY: hidden; OVERFLOW: hidden; POSITION: absolute; WIDTH: 1px; HEIGHT: 1px'></div>"
Response.Write "       </td>"
Response.Write "      </tr>"
Response.Write "      <input type='hidden' ID='ContentEdit' value=''>"
Response.Write "      <input type='hidden' ID='ModeEdit' value=''>"
Response.Write "      <input type='hidden' ID='ContentLoad' value=''>"
Response.Write "      <input type='hidden' ID='ContentFlag' value='0'>"
Response.Write "     </table>"
Response.Write "    </td>"
Response.Write "   </tr>"
Response.Write "</table>"

%>
<script language="VBScript">

Function Resumeblank(ByVal Content)
    if Content="" then 
        Resumeblank=Content 
        Exit Function
    end if
    Dim strHtml, strHtml2, Num, Numtemp, Strtemp, i
    strHtml = Replace(Content, "<DIV", "<div")
    strHtml = Replace(strHtml, "</DIV>", "</div>")
    strHtml = Replace(strHtml, "<TABLE", "<table")
    strHtml = Replace(strHtml, "</TABLE>", vbCrLf & "</table>" & vbCrLf)
    strHtml = Replace(strHtml, "<TBODY>", "")
    strHtml = Replace(strHtml, "</TBODY>", "" & vbCrLf)
    strHtml = Replace(strHtml, "<TR", "<tr")
    strHtml = Replace(strHtml, "</TR>", vbCrLf & "</tr>" & vbCrLf)
    strHtml = Replace(strHtml, "<TD", "<td")
    strHtml = Replace(strHtml, "</TD>", "</td>")
    strHtml = Replace(strHtml, "<"&"!--", vbCrLf & "<"&"!--")
    strHtml = Replace(strHtml, "<SELECT", vbCrLf & "<Select")
    strHtml = Replace(strHtml, "</SELECT>", vbCrLf & "</Select>")
    strHtml = Replace(strHtml, "<OPTION", vbCrLf & "  <Option")
    strHtml = Replace(strHtml, "</OPTION>", "</Option>")
    strHtml = Replace(strHtml, "<INPUT", vbCrLf & "  <Input")
    strHtml = Replace(strHtml, "<" & "script", vbCrLf & "<"&"script")
    strHtml = Replace(strHtml, "&amp;", "&")
    strHtml = Replace(strHtml, "{$--", vbCrLf & "<"&"!--$")
    strHtml = Replace(strHtml, "--}", "$--"&">")
    arrContent = Split(strHtml, vbCrLf)
    For i = 0 To UBound(arrContent)
        Numtemp = False
        If InStr(arrContent(i), "<table") > 0 Then
            Numtemp = True
            If Strtemp <> "<table" And Strtemp <> "</table>" Then
                Num = Num + 2
            End If
            Strtemp = "<table"
        ElseIf InStr(arrContent(i), "<tr") > 0 Then
            Numtemp = True
            If Strtemp <> "<tr" And Strtemp <> "</tr>" Then
                Num = Num + 2
            End If
            Strtemp = "<tr"
        ElseIf InStr(arrContent(i), "<td") > 0 Then
            Numtemp = True
            If Strtemp <> "<td" And Strtemp <> "</td>" Then
                Num = Num + 2
            End If
            Strtemp = "<td"
        ElseIf InStr(arrContent(i), "</table>") > 0 Then
            Numtemp = True
            If Strtemp <> "</table>" And Strtemp <> "<table" Then
                Num = Num - 2
            End If
            Strtemp = "</table>"
        ElseIf InStr(arrContent(i), "</tr>") > 0 Then
            Numtemp = True
            If Strtemp <> "</tr>" And Strtemp <> "<tr" Then
                Num = Num - 2
            End If
            Strtemp = "</tr>"
        ElseIf InStr(arrContent(i), "</td>") > 0 Then
            Numtemp = True
            If Strtemp <> "</td>" And Strtemp <> "<td" Then
                Num = Num - 2
            End If
            Strtemp = "</td>"
        ElseIf InStr(arrContent(i), "<"&"!--") > 0 Then
            Numtemp = True
        End If

        If Num < 0 Then Num = 0
        If Trim(arrContent(i)) <> "" Then
            If i = 0 Then
                strHtml2 = String(Num, " ") & arrContent(i)
            ElseIf Numtemp = True Then
                strHtml2 = strHtml2 & vbCrLf & String(Num, " ") & arrContent(i)
            Else
                strHtml2 = strHtml2 & vbCrLf & arrContent(i)
            End If
        End If
    Next
    Resumeblank = strHtml2
End Function
'==================================================
'函数名：ScriptHtml
'作  用：过滤html标记
'参  数：ConStr  ------ 要过滤的字符串
'参  数：TagName ------ 字符串种型
'参  数：FType   ------ 过滤的类型
'参  数：FontFilterText   ------ 过滤含有制定字符的标记
'==================================================
Function ScriptHtml(ByVal ConStr,ByVal TagName,ByVal FType,ByVal FontFilterText)
    Dim regEx, Match, Matches
    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Global = True
    Select Case FType
        Case 1
            regEx.Pattern = "<" & TagName & "([^>])*>"
            ConStr = regEx.Replace(ConStr, "")
        Case 2
            regEx.Pattern = "<" & TagName & "([^>])*>.*?</" & TagName & "([^>])*>"
            ConStr = regEx.Replace(ConStr, "")
        Case 3
            regEx.Pattern = "<" & TagName & "([^>])*>"
            ConStr = regEx.Replace(ConStr, "")
            regEx.Pattern = "</" & TagName & "([^>])*>"
            ConStr = regEx.Replace(ConStr, "")
        Case 4
            regEx.Pattern =  "<" & TagName & "([^>])*>.*?</" & TagName & "([^>])*>"
            Set Matches = regEx.Execute(ConStr)
            For Each Match In Matches
                If InStr(Match.Value, FontFilterText) > 0 Then
                    ConStr = Replace(ConStr, Match.Value, "")
                End If
            Next
    End Select
    ScriptHtml = ConStr
    Set regEx = Nothing
End Function
</script>
<script type="text/javascript">
// 系统初试化 和系统运用 函数组开始
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    SEP_PADDING = 5;
    HANDLE_PADDING = 7;
    window.onerror = ResumeError;
    // 改变模式：代码、编辑、文本、预览
    var sCurrMode = 'EDIT';
    var bEditMode = true;
    var yanchicss= false;
    ModeEdit.value = 'EDIT';

    // 连接对象
    // 浏览器版本检测
    var BrowserInfo = new Object() ;
    BrowserInfo.MajorVer = navigator.appVersion.match(/MSIE (.)/)[1] ;
    BrowserInfo.MinorVer = navigator.appVersion.match(/MSIE .\.(.)/)[1] ;
    BrowserInfo.IsIE55OrMore = BrowserInfo.MajorVer >= 6 || ( BrowserInfo.MajorVer >= 5 && BrowserInfo.MinorVer >= 5 ) ;

    var yToolbars =   new Array();
    var YInitialized = false;
    var bLoad=false;
    var pureText=true;
    var EditMode=true;
    var SourceMode=false;
    var PreviewMode=false;
    var CurrentMode=0;

    var sLinkFieldName ="<%=tContentID%>";
    var edithead="<html><head><META http-equiv=Content-Type content=text/html; charset=gb2312><link href='<%=InstallDir%>Skin/DefaultSkin.css' rel='stylesheet' type='text/css'></head>";
    <%if ShowType <> 1 Then%>
        edithead=edithead+"<body leftmargin=0 topmargin=0 style='background:url(<%=InstallDir%>powereasy/skin/blue/powereasy.gif)'>";
    <%end if %>

    var content
    //屏蔽错误
    function ResumeError() {
        return true;
    }
    function EditorEdit() {
    <%
    if ShowType = 1 and TemplateType <> 0  Then '0 自定义标签
        Response.Write "if (document.all){" & vbCrLf
        if TemplateType=1 then 
            Response.Write " parent.document.form1.EditorEdit.disabled=false;" & vbCrLf
            Response.Write " parent.document.form1.EditorMix.disabled=false;" & vbCrLf
        elseif TemplateType=2 then
            Response.Write " parent.document.form1.EditorEdit.disabled=false;" & vbCrLf
            Response.Write " parent.document.form1.EditorMix.disabled=false;" & vbCrLf
            Response.Write " parent.document.form1.EditorEdit2.disabled=false;" & vbCrLf
            Response.Write " parent.document.form1.EditorMix2.disabled=false;" & vbCrLf
        end if
        Response.write "}else{" &vbCrlf
        Response.write "    setTimeout(""EditorEdit()"",1000);" & vbCrLf
        Response.write "}" & vbCrlf
    End if
    %>
    }
    //程序初始化
    function document.onreadystatechange(){
        if (YInitialized) return;
        YInitialized = true;
        var i, s, curr;
        for (i=0; i<document.body.all.length; i++){
            curr=document.body.all[i];
            if (curr.className == 'yToolbar'){
                InitTB(curr);
                yToolbars[yToolbars.length] = curr;
            }
        }
        DoTemplate();
        oLinkField = parent.document.getElementsByName(sLinkFieldName)[0];
        if (ContentFlag.value=="0") {
            ContentEdit.value = oLinkField.value;
            ContentLoad.value = oLinkField.value;
            ModeEdit.value = 'EDIT'
            ContentFlag.value = "1";
        }
        
        window.onresize = DoTemplate;
        content=edithead +ContentEdit.value;
        EditorEdit();

        content = content.replace("[/textarea]", "</textarea>");
        
        HtmlEdit.document.open();
        HtmlEdit.document.write(content);
        HtmlEdit.document.close();  
        HtmlEdit.document.designMode='On';
        HtmlEdit.document.onkeydown = new Function('return onKeyDown(HtmlEdit.event);');
        HtmlEdit.document.onmouseup = new Function('return onMouseUp(HtmlEdit.event,<%=Clng(TemplateType)%>);');
        HtmlEdit.document.oncontextmenu=new Function('return showContextMenu(HtmlEdit.event);');
    }

    function yToolbarsCss(){
        if (document.all){
            var i, s, curr;
            for (i=0; i<document.body.all.length; i++){
                curr=document.body.all[i];
                if (curr.className == 'yToolbar')
                {
                    InitTB(curr);
                    yToolbars[yToolbars.length] = curr;
                }
            }
            DoTemplate();
        }else{
            setTimeout("yToolbarsCss()",1000);
        }
    }
    function InitBtn(btn){
        btn.onmouseover = BtnMouseOver;
        btn.onmouseout = BtnMouseOut;
        btn.onmousedown = BtnMouseDown;
        btn.onmouseup = BtnMouseUp;
        btn.ondragstart = YCancelEvent;
        btn.onselectstart = YCancelEvent;
        btn.onselect = YCancelEvent;
        btn.YUSERONCLICK = btn.onclick;
        btn.onclick = YCancelEvent;
        btn.YINITIALIZED = true;
        return true;
    }
    function InitBtnMenu(BtnMenu){
        BtnMenu.onmouseover = BtnMenuMouseOver;
        BtnMenu.onmouseout = BtnMenuMouseOut;
        BtnMenu.onmousedown = BtnMenuMouseDown;
        BtnMenu.onmouseup = BtnMenuMouseUp;
        BtnMenu.ondragstart = YCancelEvent;
        BtnMenu.onselectstart = YCancelEvent;
        BtnMenu.onselect = YCancelEvent;
        BtnMenu.YUSERONCLICK = BtnMenu.onclick;
        BtnMenu.onclick = YCancelEvent;
        BtnMenu.YINITIALIZED = true;
        return true;
    }
    function InitTB(y){
        if (!document.all){
            setTimeout("InitTB("+ y +")",1000);
            return;
        }
        y.TBWidth = 0;
        if (! PopulateTB(y)) return false;
        y.style.posWidth = y.TBWidth;
        return true;
    }
    function YCancelEvent(){
        event.returnValue=false;
        event.cancelBubble=true;
        return false;
    }
    function PopulateTB(y){
        var i, elements, element;
        elements = y.children;
        for (i=0; i<elements.length; i++) {
            element = elements[i];
            if (element.tagName == 'SCRIPT' || element.tagName == '!') continue;
            switch (element.className) {
            case 'Btn':
                if (element.YINITIALIZED == null)   {
                if (! InitBtn(element))
                    return false;
                }
                element.style.posLeft = y.TBWidth;
                y.TBWidth   += element.offsetWidth + 1;
                break;
            case 'BtnMenu':
                if (element.YINITIALIZED == null)   {
                if (! InitBtnMenu(element))
                    return false;
                }
                element.style.posLeft = y.TBWidth;
                y.TBWidth   += element.offsetWidth + 1;
                break;
            case 'TBGen':
                element.style.posLeft = y.TBWidth;
                y.TBWidth   += element.offsetWidth + 1;
                break;
            case 'TBSep':
                element.style.posLeft = y.TBWidth   + 2;
                y.TBWidth   += SEP_PADDING;
                break;
            case 'TBHandle':
                element.style.posLeft = 2;
                y.TBWidth   += element.offsetWidth + HANDLE_PADDING;
                break;
            default:
            return false;
            }
        }
        y.TBWidth += 1;
        return true;
    }
    function TemplateTBs(){
        NumTBs = yToolbars.length;
        if (NumTBs == 0) return;
        var i;
        var ScrWid = (document.body.offsetWidth) - 6;
        var TotalLen = ScrWid;
        for (i = 0 ; i < NumTBs ; i++) {
            TB = yToolbars[i];
            if (TB.TBWidth > TotalLen) TotalLen = TB.TBWidth;
        }
        var PrevTB;
        var LastStart = 0;
        var RelTop = 0;
        var LastWid, CurrWid;
        var TB = yToolbars[0];
        TB.style.posTop = 0;
        TB.style.posLeft = 0;
        var Start = TB.TBWidth;
        for (i = 1 ; i < yToolbars.length ; i++) {
            PrevTB = TB;
            TB = yToolbars[i];
            CurrWid = TB.TBWidth;
            if ((Start + CurrWid) > ScrWid) {
                Start = 0;
                LastWid = TotalLen - LastStart;
            }else {
                LastWid =PrevTB.TBWidth;
                RelTop -=TB.offsetHeight;
            }
            TB.style.posTop = RelTop;
            TB.style.posLeft = Start;
            PrevTB.style.width = LastWid;
            LastStart = Start;
            Start += CurrWid;
        }
        TB.style.width = TotalLen - LastStart;
        i--;
        TB = yToolbars[i];
        var TBInd = TB.sourceIndex;
        var A = TB.document.all;
        var item;
        for (i in A) {
            item = A.item(i);
            if (! item) continue;
            if (! item.style) continue;
            if (item.sourceIndex <= TBInd) continue;
            if (item.style.position == 'absolute') continue;
            item.style.posTop = RelTop;
        }
    }
    function DoTemplate(){
        TemplateTBs();
    }
    function BtnMouseOver(){
        if (event.srcElement.tagName != 'IMG') return false;
        var image = event.srcElement;
        var element = image.parentElement;
        if (image.className == 'Ico') element.className = 'BtnMouseOverUp';
        else if (image.className == 'IcoDown') element.className = 'BtnMouseOverDown';
        event.cancelBubble = true;
    }
    function BtnMouseOut(){
        if (event.srcElement.tagName != 'IMG') {
            event.cancelBubble = true;
            return false;
        }
        var image = event.srcElement;
        var element = image.parentElement;
        yRaisedElement = null;
        element.className = 'Btn';
        image.className = 'Ico';
        event.cancelBubble = true;
    }
    function BtnMouseDown(){
        if (event.srcElement.tagName != 'IMG') {
            event.cancelBubble = true;
            event.returnValue=false;
            return false;
        }
        var image = event.srcElement;
        var element = image.parentElement;
        element.className = 'BtnMouseOverDown';
        image.className = 'IcoDown';
        event.cancelBubble = true;
        event.returnValue=false;
        return false;
    }
    function BtnMouseUp(){
        if (event.srcElement.tagName != 'IMG') {
            event.cancelBubble = true;
            return false;
        }
        var image = event.srcElement;
        var element = image.parentElement;
        if (navigator.appVersion.match(/8./i)=='8.') 
		
        {
            if (element.YUSERONCLICK) eval(element.YUSERONCLICK + 'onclick(event)');   
        }
        else
        {
          if(document.documentMode === 5) {
          if (element.YUSERONCLICK) eval(element.YUSERONCLICK + 'onclick(event)');   
          }	
          else{  
            if (element.YUSERONCLICK) eval(element.YUSERONCLICK + 'anonymous()'); 
          }
		  
        }
        element.className = 'BtnMouseOverUp';
        image.className = 'Ico';
        event.cancelBubble = true;
        return false;
    }
    function BtnMenuMouseOver(){
      if (event.srcElement.tagName != 'IMG') return false;
      var image = event.srcElement;
      var element = image.parentElement;
      if (image.className == 'Ico') element.className = 'BtnMenuMouseOverUp';
      else if (image.className == 'IcoDown') element.className = 'BtnMenuMouseOverDown';
      event.cancelBubble = true;
    }
    function BtnMenuMouseOut(){
        if (event.srcElement.tagName != 'IMG') {
            event.cancelBubble = true;
            return false;
        }
        var image = event.srcElement;
        var element = image.parentElement;
        yRaisedElement = null;
        element.className = 'BtnMenu';
        image.className = 'Ico';
        event.cancelBubble = true;
    }
    function BtnMenuMouseDown(){
        if (event.srcElement.tagName != 'IMG') {
            event.cancelBubble = true;
            event.returnValue=false;
            return false;
        }
        var image = event.srcElement;
        var element = image.parentElement;
        element.className = 'BtnMenuMouseOverDown';
        image.className = 'IcoDown';
        event.cancelBubble = true;
        event.returnValue=false;
        return false;
    }
    function BtnMenuMouseUp(){
        if (event.srcElement.tagName != 'IMG') {
            event.cancelBubble = true;
            return false;
        }
        var image = event.srcElement;
        var element = image.parentElement;
        if (navigator.appVersion.match(/8./i)=='8.') 
		
        {
            if (element.YUSERONCLICK) eval(element.YUSERONCLICK + 'onclick(event)');   
        }
        else
        {
          if(document.documentMode === 5) {
          if (element.YUSERONCLICK) eval(element.YUSERONCLICK + 'onclick(event)');   
          }	
          else{  
            if (element.YUSERONCLICK) eval(element.YUSERONCLICK + 'anonymous()'); 
          }
		  
        }
        element.className = 'BtnMenuMouseOverUp';
        image.className = 'Ico';
        event.cancelBubble = true;
        return false;
    }
    // 系统初试化 和系统运用 函数组结速
    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // 表格处理定义函数组开始
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    // 表格相关全局变量
    var selectedTD
    var selectedTR
    var selectedTBODY
    var selectedTable
    // 显示隐藏表格
    var borderShown = "yes"

    // 插入表格
    function TableInsert(){
        if (!isTableSelected()){
            ShowDialog('Editor/editor_table.asp', 350, 410, true);
        }
    }
    // 修改表格属性
    function TableProp(){
        if (isTableSelected()||isCursorInTableCell()){
            ShowDialog('Editor/editor_table.asp?action=modify', 350, 410, true);
        }
    }
    // 修改单元格属性
    function TableCellProp(){
        if (isCursorInTableCell()){
            ShowDialog('Editor/editor_tablecell.asp', 350, 310, true);
        }
    }
    // 拆分单元格
    function TableCellSplit(){
        if (isCursorInTableCell()){
            ShowDialog('Editor/editor_tablecellsplit.asp', 200, 150, true);
        }
    }
    // 修改表格行属性
    function TableRowProp(){
        if (isCursorInTableCell()){
            ShowDialog('Editor/editor_tablecell.asp?action=row', 350, 310, true);
        }
    }
    // 插入行（在上方）
    function TableRowInsertAbove() {
        if (isCursorInTableCell()){
            var numCols = 0
            allCells = selectedTR.cells
            for (var i=0;i<allCells.length;i++) {
                numCols = numCols + allCells[i].getAttribute('colSpan')
            }
            var newTR = selectedTable.insertRow(selectedTR.rowIndex)
            for (i = 0; i < numCols; i++) {
                newTD = newTR.insertCell()
                newTD.innerHTML = "&nbsp;"
                if (borderShown == "yes") {
                    newTD.runtimeStyle.border = "1px dotted #330000"
                }
            }
        }    
    }
    // 插入行（在下方）
    function TableRowInsertBelow() {
        if (isCursorInTableCell()){
            var numCols = 0
            allCells = selectedTR.cells
            for (var i=0;i<allCells.length;i++) {
                numCols = numCols + allCells[i].getAttribute('colSpan')
            }
            var newTR = selectedTable.insertRow(selectedTR.rowIndex+1)
            for (i = 0; i < numCols; i++) {
                newTD = newTR.insertCell()
                newTD.innerHTML = "&nbsp;"
                
                if (borderShown == "yes") {
                    newTD.runtimeStyle.border = "1px dotted #BFBFBF"
                }
            }
        }
    }
    // 合并行（向下方）
    function TableRowMerge() {
        if (isCursorInTableCell()) {
            var rowSpanTD = selectedTD.getAttribute('rowSpan')
            allRows = selectedTable.rows
            if (selectedTR.rowIndex +1 != allRows.length) {
                var allCellsInNextRow = allRows[selectedTR.rowIndex+selectedTD.rowSpan].cells
                var addRowSpan = allCellsInNextRow[selectedTD.cellIndex].getAttribute('rowSpan')
                var moveTo = selectedTD.rowSpan
                if (!addRowSpan) addRowSpan = 1;
                selectedTD.rowSpan = selectedTD.rowSpan + addRowSpan
                allRows[selectedTR.rowIndex + moveTo].deleteCell(selectedTD.cellIndex)
            }
        }

    }
    // 拆分行
    function TableRowSplit(nRows){
        if (!isCursorInTableCell()) return;
        if (nRows<2) return;

        var addRows = nRows - 1;
        var addRowsNoSpan = addRows;

        var nsLeftColSpan = 0;
        for (var i=0; i<selectedTD.cellIndex; i++){
            nsLeftColSpan += selectedTR.cells[i].colSpan;
        }
        var allRows = selectedTable.rows;
        // rowspan>1时
        while (selectedTD.rowSpan > 1 && addRowsNoSpan > 0){
            var nextRow = allRows[selectedTR.rowIndex+selectedTD.rowSpan-1];
            selectedTD.rowSpan -= 1;

            var ncLeftColSpan = 0;
            var position = -1;
            for (var n=0; n<nextRow.cells.length; n++){
                ncLeftColSpan += nextRow.cells[n].getAttribute('colSpan');
                if (ncLeftColSpan>nsLeftColSpan){
                    position = n;
                    break;
                }
            }

            var newTD=nextRow.insertCell(position);
            newTD.innerHTML = "&nbsp;";

            if (borderShown == "yes") {
                newTD.runtimeStyle.border = "1px dotted #BFBFBF";
            }
                
            addRowsNoSpan -= 1;
        }
        // rowspan=1时
        for (var n=0; n<addRowsNoSpan; n++){
            var numCols = 0

            allCells = selectedTR.cells
            for (var i=0;i<allCells.length;i++) {
                numCols = numCols + allCells[i].getAttribute('colSpan')
            }

            var newTR = selectedTable.insertRow(selectedTR.rowIndex+1)

            // 上方行的rowspan达到这行
            for (var j=0; j<selectedTR.rowIndex; j++){
                for (var k=0; k<allRows[j].cells.length; k++){
                    if ((allRows[j].cells[k].rowSpan>1)&&(allRows[j].cells[k].rowSpan>=selectedTR.rowIndex-allRows[j].rowIndex+1)){
                        allRows[j].cells[k].rowSpan += 1;
                    }
                }
            }
            // 当前行
            for (i = 0; i < allCells.length; i++) {
                if (i!=selectedTD.cellIndex){
                    selectedTR.cells[i].rowSpan += 1;
                }else{
                    newTD = newTR.insertCell();
                    newTD.colSpan = selectedTD.colSpan;
                    newTD.innerHTML = "&nbsp;";

                    if (borderShown == "yes") {
                        newTD.runtimeStyle.border = "1px dotted #BFBFBF";
                    }
                }
            }
        }

    }
    // 删除行
    function TableRowDelete() {
        if (isCursorInTableCell()) {
            selectedTable.deleteRow(selectedTR.rowIndex)
        }
    }
    // 插入列（在左侧）
    function TableColInsertLeft() {
        if (isCursorInTableCell()) {
            moveFromEnd = (selectedTR.cells.length-1) - (selectedTD.cellIndex)
            allRows = selectedTable.rows
            for (i=0;i<allRows.length;i++) {
                rowCount = allRows[i].cells.length - 1
                position = rowCount - moveFromEnd
                if (position < 0) {
                    position = 0
                }
                newCell = allRows[i].insertCell(position)
                newCell.innerHTML = "&nbsp;"

                if (borderShown == "yes") {
                    newCell.runtimeStyle.border = "1px dotted #BFBFBF"
                }
            }
        }
    }

    // 插入列（在右侧）
    function TableColInsertRight() {
        if (isCursorInTableCell()) {
            moveFromEnd = (selectedTR.cells.length-1) - (selectedTD.cellIndex)
            allRows = selectedTable.rows
            for (i=0;i<allRows.length;i++) {
                rowCount = allRows[i].cells.length - 1
                position = rowCount - moveFromEnd
                if (position < 0) {
                    position = 0
                }
                newCell = allRows[i].insertCell(position+1)
                newCell.innerHTML = "&nbsp;"

                if (borderShown == "yes") {
                    newCell.runtimeStyle.border = "1px dotted #BFBFBF"
                }
            }    
        }
    }

    // 合并列
    function TableColMerge() {
        if (isCursorInTableCell()) {

            var colSpanTD = selectedTD.getAttribute('colSpan')
            allCells = selectedTR.cells

            if (selectedTD.cellIndex + 1 != selectedTR.cells.length) {
                var addColspan = allCells[selectedTD.cellIndex+1].getAttribute('colSpan')
                selectedTD.colSpan = colSpanTD + addColspan
                selectedTR.deleteCell(selectedTD.cellIndex+1)
            }    
        }

    }

    // 删除列
    function TableColDelete() {
        if (isCursorInTableCell()) {
            moveFromEnd = (selectedTR.cells.length-1) - (selectedTD.cellIndex)
            allRows = selectedTable.rows
            for (var i=0;i<allRows.length;i++) {
                endOfRow = allRows[i].cells.length - 1
                position = endOfRow - moveFromEnd
                if (position < 0) {
                    position = 0
                }
                    

                allCellsInRow = allRows[i].cells
                    
                if (allCellsInRow[position].colSpan > 1) {
                    allCellsInRow[position].colSpan = allCellsInRow[position].colSpan - 1
                } else { 
                    allRows[i].deleteCell(position)
                }
            }
        }
    }
    // 拆分列
    function TableColSplit(nCols){
        if (!isCursorInTableCell()) return;
        if (nCols<2) return;

        var addCols = nCols - 1;
        var addColsNoSpan = addCols;
        var newCell;

        var nsLeftColSpan = 0;
        var nsLeftRowSpanMoreOne = 0;
        for (var i=0; i<selectedTD.cellIndex; i++){
            nsLeftColSpan += selectedTR.cells[i].colSpan;
            if (selectedTR.cells[i].rowSpan > 1){
                nsLeftRowSpanMoreOne += 1;
            }
        }

        var allRows = selectedTable.rows
        // colSpan>1时
        while (selectedTD.colSpan > 1 && addColsNoSpan > 0) {
            newCell = selectedTR.insertCell(selectedTD.cellIndex+1);
            newCell.innerHTML = "&nbsp;"
            if (borderShown == "yes") {
                newCell.runtimeStyle.border = "1px dotted #BFBFBF"
            }
            selectedTD.colSpan -= 1;
            addColsNoSpan -= 1;
        }
        // colSpan=1时
        for (i=0;i<allRows.length;i++) {
            var ncLeftColSpan = 0;
            var position = -1;
            for (var n=0; n<allRows[i].cells.length; n++){
                ncLeftColSpan += allRows[i].cells[n].getAttribute('colSpan');
                if (ncLeftColSpan+nsLeftRowSpanMoreOne>nsLeftColSpan){
                    position = n;
                    break;
                }
            }
            
            if (selectedTR.rowIndex!=i){
                if (position!=-1){
                    allRows[i].cells[position+nsLeftRowSpanMoreOne].colSpan += addColsNoSpan;
                }
            }else{
                for (var n=0; n<addColsNoSpan; n++){
                    newCell = allRows[i].insertCell(selectedTD.cellIndex+1)
                    newCell.innerHTML = "&nbsp;"
                    newCell.rowSpan = selectedTD.rowSpan;

                    if (borderShown == "yes") {
                        newCell.runtimeStyle.border = "1px dotted #BFBFBF"
                    }
                }
            }
        }
    }
    // 是否选中表格
    function isTableSelected() {
        if (HtmlEdit.document.selection.type == "Control") {
            var oControlRange = HtmlEdit.document.selection.createRange();
            if (oControlRange(0).tagName.toUpperCase() == "TABLE") {
                selectedTable = HtmlEdit.document.selection.createRange()(0);
                return true;
            }    
        }
    } 
    // 光标是否在表格中
    function isCursorInTableCell() {
        if (HtmlEdit.document.selection.type != "Control") {
            var elem = HtmlEdit.document.selection.createRange().parentElement()
            while (elem.tagName.toUpperCase() != "TD" && elem.tagName.toUpperCase() != "TH"){
                elem = elem.parentElement
                    if (elem == null)
                    break
            }
            if (elem) {
                selectedTD = elem
                selectedTR = selectedTD.parentElement
                selectedTBODY =  selectedTR.parentElement
                selectedTable = selectedTBODY.parentElement
                return true
            }
        }
    }
    // 表格处理定义函数组结速
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // 右键菜单定义函数组开始
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // 菜单常量
    var sMenuHr="<tr><td align=center valign=middle height=2><TABLE border=0 cellpadding=0 cellspacing=0 width=128 height=2><tr><td height=1 class=HrShadow><\/td><\/tr><tr><td height=1 class=HrHighLight><\/td><\/tr><\/TABLE><\/td><\/tr>";
    var sMenu1="<TABLE border=0 cellpadding=0 cellspacing=0 class=Menu2 width=150><tr><td width=18 valign=bottom align=center style='background:url(Editor/images/contextmenu.gif);background-positionY: 35%; background-repeat:no-repeat;'><\/td><td width=132 class=RightBg><TABLE border=0 cellpadding=0 cellspacing=0>";
    var sMenu2="<\/TABLE><\/td><\/tr><\/TABLE>";
    // 菜单
    var oPopupMenu = null;
    if  (BrowserInfo.IsIE55OrMore){
        oPopupMenu = window.createPopup();
    }

    // 取菜单行
    function getMenuRow(s_Disabled, s_Event, s_Image, s_Html) {
        var s_MenuRow = "";
        s_MenuRow = "<tr><td align=center valign=middle><TABLE border=0 cellpadding=0 cellspacing=0 width=132><tr "+s_Disabled+"><td valign=middle height=20 class=MouseOut onMouseOver=this.className='MouseOver'; onMouseOut=this.className='MouseOut';";
        if (s_Disabled==""){
            s_MenuRow += " onclick=\"parent."+s_Event+";parent.oPopupMenu.hide();\"";
        }
        s_MenuRow += ">"
        if (s_Image !=""){
            s_MenuRow += "&nbsp;<img border=0 src='Editor/Images/"+s_Image+"' width=18 height=18 align=absmiddle "+s_Disabled+">&nbsp;";
        }else{
            s_MenuRow += "&nbsp;";
        }
        s_MenuRow += s_Html+"<\/td><\/tr><\/TABLE><\/td><\/tr>";
        return s_MenuRow;
    }
    // 取标准的format菜单行
    function getFormatMenuRow(menu, html, image){
        var s_Disabled = "";
        if (!HtmlEdit.document.queryCommandEnabled(menu)){
            s_Disabled = "disabled";
        }
        var s_Event = "format('"+menu+"')";
        var s_Image = menu+".gif";
        if (image){
            s_Image = image;
        }
        return getMenuRow(s_Disabled, s_Event, s_Image, html)
    }
    // 暂时解决一下 不能通用右键的问题
    function getFormatMenuRow2(menu, html, image){
        var s_Disabled = "";
        if (!HtmlEdit.document.queryCommandEnabled(menu)){
            s_Disabled = "disabled";
        }
        var s_Event = "format2('"+menu+"')";
        var s_Image = menu+".gif";
        if (image){
            s_Image = image;
        }
        return getMenuRow(s_Disabled, s_Event, s_Image, html)
    }
    //表格菜单
    function tableMenu(){
        if (!bEditMode) return false;
        var sMenu = ""
        var width = 150;
        var height = 0;

        var lefter = event.clientX;
        var leftoff = event.offsetX
        var topper = event.clientY;
        var topoff = event.offsetY;

        var oPopDocument = oPopupMenu.document;
        var oPopBody = oPopupMenu.document.body;

        sMenu += getTableMenuRow("TableInsert");
        sMenu += getTableMenuRow("TableProp");
        sMenu += sMenuHr;
        sMenu += getTableMenuRow("TableCell");
        height = 306;
    }
    // 取表格菜单行
    function getTableMenuRow(what){
        var s_Menu = "";
        var s_Disabled = "disabled";
        switch(what){
        case "TableInsert":
            if (!isTableSelected()) s_Disabled="";
            s_Menu += getMenuRow(s_Disabled, "TableInsert()", "table_cr.gif", "插入表格...")
            break;
        case "TableProp":
            if (isTableSelected()||isCursorInTableCell()) s_Disabled="";
            s_Menu += getMenuRow(s_Disabled, "TableProp()", "table_sx.gif", "表格属性...")
            break;
        case "TableCell":
            if (isCursorInTableCell()) s_Disabled="";
            s_Menu += getMenuRow(s_Disabled, "TableCellProp()", "table_sx2.gif", "单元格属性...")
            s_Menu += getMenuRow(s_Disabled, "TableCellSplit()", "table_cf.gif", "拆分单元格...")
            s_Menu += sMenuHr;
            s_Menu += getMenuRow(s_Disabled, "TableRowProp()", "table_sxh.gif", "表格行属性...")
            s_Menu += getMenuRow(s_Disabled, "TableRowInsertAbove()", "table_tr.gif", "插入行（在上方）");
            s_Menu += getMenuRow(s_Disabled, "TableRowInsertBelow()", "table_trx.gif", "插入行（在下方）");
            s_Menu += getMenuRow(s_Disabled, "TableRowMerge()", "table_hbx.gif", "合并行（向下方）");
            s_Menu += getMenuRow(s_Disabled, "TableRowSplit(2)", "table_cfh.gif", "拆分行");
            s_Menu += getMenuRow(s_Disabled, "TableRowDelete()", "table_trdel.gif", "删除行");
            s_Menu += sMenuHr;
            s_Menu += getMenuRow(s_Disabled, "TableColInsertLeft()", "table_td.gif", "插入列（在左侧）");
            s_Menu += getMenuRow(s_Disabled, "TableColInsertRight()", "table_tdr.gif", "插入列（在右侧）");
            s_Menu += getMenuRow(s_Disabled, "TableColMerge()", "table_hby.gif", "合并列（向右侧）");
            s_Menu += getMenuRow(s_Disabled, "TableColSplit(2)", "table_cf.gif", "拆分列");
            s_Menu += getMenuRow(s_Disabled, "TableColDelete()", "table_tddel.gif", "删除列");
            break;
        }
        return s_Menu;
    }
    // 右键是否在编辑状态
    function isyou(){
        var range = HtmlEdit.document.selection.createRange();
        var RangeType = HtmlEdit.document.selection.type;
        if (RangeType == "Text"){
            return true;
        }  
    }
    // 右键调用类型
    function youjiantype(){
        if (youjian=true){
            return true;
        }  
    }

    // 右键菜单
    function showContextMenu(event){

        if (!bEditMode) return false;
        var width = 150;
        var height = 0;
        var lefter = event.clientX;
        var topper = event.clientY;

        var oPopDocument = oPopupMenu.document;
        var oPopBody = oPopupMenu.document.body;

        var sMenu="";
        
        sMenu += getFormatMenuRow2("cut", "剪切");
        sMenu += getFormatMenuRow2("copy", "复制");
        sMenu += getFormatMenuRow2("paste", "常规粘贴");
        sMenu += getFormatMenuRow2("delete", "删除");
        <% if ShowType = 1 then %>
        sMenu += sMenuHr;
        sMenu += getMenuRow("", "insert('Label')", "label.gif", "添加标签...");
        height +=22;
        if (isControlSelected("IMG")){
            sMenu += getMenuRow("", "insert('editLabel')", "label2.gif", "编辑标签...");
            height+=21
        }
        <%elseif ShowType = 0 then%>
            sMenu += sMenuHr;
            sMenu += getMenuRow("","insert('page')","page.gif","添加分页标签");
            sMenu += getMenuRow("","insert('pagetitle')","pagetitle.gif","插入带标题的分页");
            sMenu += getMenuRow("","insert('copypagetitle')","pagetitle.gif","复制成带标题的分页");
            sMenu += getMenuRow("","insert('calljsad')","Jscript.gif","添加广告JS调用");
            height += 80;
        <% End if %>
        height += 102;
        if (HtmlEdit.document.selection.type == "Control") {
            <% if ShowType = 1 then %>
                sMenu += sMenuHr;
                sMenu += getMenuRow("","insert('ReplaceLabel')","label2.gif","替换为标签");
                height +=22;
            <% End if %>
            sMenu += getMenuRow("", "insert('Attribute')", "label3.gif", "代码属性...");    
            height+= 19;
        }    
        if (sCurrMode=="EDIT"){

            if (isyou()){
        
                <%if ShowType = 0 then %>
                sMenu += getMenuRow("", "insert('title')", "article_title.gif", "设置为标题");
                sMenu += getMenuRow("", "insert('keyword')", "article_keyword.gif", "设置为关键字");
                sMenu += getMenuRow("","insert('Intro')","article_Intro.gif","设置为文章简介");
                sMenu += sMenuHr;
                height+=65;
                <%elseif ShowType= 4 then %>
                sMenu+=  getMenuRow("","insert('ProductName')","article_title.gif", "设置为商品名称");
                sMenu += getMenuRow("", "insert('keyword')", "article_keyword.gif", "设置为关键字");
                sMenu += getMenuRow("","insert('ProductIntro')","article_Intro.gif","设置为商品简介");
                sMenu += sMenuHr;
                height+=65;
                <% End if %>
                sMenu += getMenuRow("", "insert('fgcolor')", "fgcolor.gif", "文字颜色");
                sMenu += getMenuRow("", "insert('fgbgcolor')", "fgbgcolor.gif", "文字背景色");
                sMenu += getMenuRow("", "format('bold')", "bold.gif", "文字加粗");
                sMenu += getMenuRow("", "format('italic')", "italic.gif", "文字斜体");
                sMenu += getMenuRow("", "format('underline')", "underline.gif", "文字下划线");
                sMenu += getMenuRow("", "format('StrikeThrough')", "strikethrough.gif", "文字删除线");
                height += 119;
            }

            if (isCursorInTableCell()){
                sMenu += getTableMenuRow("TableProp");
                sMenu += getTableMenuRow("TableCell");
                sMenu += sMenuHr;
                height += 286;
            }

            if (isControlSelected("TABLE")){
                sMenu += getTableMenuRow("TableProp");
                sMenu += sMenuHr;
                height += 22;
            }

            if (isControlSelected("IMG")){

                sMenu += getMenuRow("", "insert('pic')", "img.gif", "图片属性...");    
                sMenu += sMenuHr;
                sMenu += getMenuRow("", "imgalign('left')", "imgleft.gif", "图片左环绕...");
                sMenu += getMenuRow("", "imgalign('center')", "imgcenter.gif", "图片单行居中...");
                sMenu += getMenuRow("", "imgalign('right')", "imgright.gif", "图片右环绕...");        
                sMenu += sMenuHr;
                sMenu += getMenuRow("", "zIndex('forward')", "forward.gif", "上移一层");
                sMenu += getMenuRow("", "zIndex('backward')", "backward.gif", "下移一层");
                sMenu += sMenuHr;
                height+= 127;
            }

        }
        sMenu += getFormatMenuRow2("selectall", "全选");
        sMenu += getMenuRow("", "findstr()", "find.gif", "查找替换...");
        height += 20;

        sMenu = sMenu1 + sMenu + sMenu2;

        oPopDocument.open();
        oPopDocument.write("<head><link href=Editor/editor_dialog.css type=\"text/css\" rel=\"stylesheet\"></head><body scroll=\"no\"  leftmargin='0' topmargin='0' style='body:margin:0px;border:0px'onConTextMenu=\"event.returnValue=false;\">"+sMenu);
        oPopDocument.close();

        height+=2;
        if(lefter+width > document.body.clientWidth) lefter=lefter-width;
        oPopupMenu.show(lefter, topper, width, height, HtmlEdit.document.body);
        return false;
    }

    // 右键下拉工具栏菜单
    function showToolMenu(menu){

        if (!bEditMode) return false;
        var sMenu = ""
        var width = 150;
        var height = 0;

        var lefter = event.clientX;
        var leftoff = event.offsetX
        var topper = event.clientY;
        var topoff = event.offsetY;

        var oPopDocument = oPopupMenu.document;
        var oPopBody = oPopupMenu.document.body;

        switch(menu){
        case "font":
             // 字体菜单
            sMenu += getFormatMenuRow("superscript", "上标", "sup.gif");
            sMenu += getFormatMenuRow("subscript", "下标", "sub.gif");
            sMenu += sMenuHr;
            sMenu += getMenuRow("", "insert('big')", "tobig.gif", "字体增大");
            sMenu += getMenuRow("", "insert('small')", "tosmall.gif", "字体减小");
            sMenu += sMenuHr;
            sMenu += getFormatMenuRow("insertorderedlist", "编号", "num.gif");
            sMenu += getFormatMenuRow("insertunorderedlist", "项目符号", "list.gif");
            sMenu += getFormatMenuRow("indent", "增加缩进量", "indent.gif");
            sMenu += getFormatMenuRow("outdent", "减少缩进量", "outdent.gif");
            sMenu += sMenuHr;
            sMenu += getFormatMenuRow("insertparagraph", "插入段落", "paragraph.gif");
            sMenu += getMenuRow("", "insert('br')", "chars.gif", "插入换行符");
            height = 206;
            break;
        case "paragraph":// 段落菜单
            
            sMenu += getFormatMenuRow("JustifyLeft", "左对齐", "JustifyLeft.gif");
            sMenu += getFormatMenuRow("JustifyCenter", "居中对齐", "JustifyCenter.gif");
            sMenu += getFormatMenuRow("JustifyRight", "右对齐", "JustifyRight.gif");
            sMenu += getFormatMenuRow("JustifyFull", "两端对齐", "JustifyFull.gif");
            sMenu += sMenuHr;
            sMenu += getFormatMenuRow("insertorderedlist", "编号", "insertorderedlist.gif");
            sMenu += getFormatMenuRow("insertunorderedlist", "项目符号", "insertunorderedlist.gif");
            sMenu += getFormatMenuRow("indent", "增加缩进量", "indent.gif");
            sMenu += getFormatMenuRow("outdent", "减少缩进量", "outdent.gif");
            sMenu += sMenuHr;
            sMenu += getFormatMenuRow("insertparagraph", "插入段落", "insertparagraph.gif");
            sMenu += getMenuRow("", "insert('br')", "br.gif", "插入换行符");
            height = 204;
            break;
        case "gongshi":// 公式编辑器
            sMenu += getMenuRow("","insert('InsertEQ')", "eq1.gif", "插入公式");
            sMenu += getMenuRow("","insert('InstallEQ')", "eq2.gif", "安装公式编辑器插件");
            height = 42;
            break;
        case "edit":        // 编辑菜单
            var s_Disabled = "";
            if (history.data.length <= 1 || history.position <= 0) s_Disabled = "disabled";
            sMenu += getMenuRow(s_Disabled, "goHistory(-1)", "undo.gif", "撤消")
            if (history.position >= history.data.length-1 || history.data.length == 0) s_Disabled = "disabled";
            sMenu += getMenuRow(s_Disabled, "goHistory(1)", "redo.gif", "恢复")
            sMenu += sMenuHr;
            sMenu += getFormatMenuRow("Cut", "剪切", "cut.gif");
            sMenu += getFormatMenuRow("Copy", "复制", "copy.gif");
            sMenu += getFormatMenuRow("Paste", "常规粘贴", "paste.gif");
            sMenu += getMenuRow("", "PasteText()", "pastetext.gif", "纯文本粘贴");
            sMenu += getMenuRow("", "PasteWord()", "pasteword.gif", "从Word中粘贴");
            sMenu += sMenuHr;
            sMenu += getFormatMenuRow("delete", "删除", "del.gif");
            sMenu += getFormatMenuRow("RemoveFormat", "删除文字格式", "removeformat.gif");
            sMenu += sMenuHr;
            sMenu += getFormatMenuRow("SelectAll", "全部选中", "selectall.gif");
            sMenu += getFormatMenuRow("Unselect", "取消选择", "unselect.gif");
            sMenu += sMenuHr;
            sMenu += getMenuRow("", "findReplace()", "findreplace.gif", "查找替换");
            height = 248;
            break;
        case "object":        // 对象效果菜单
            sMenu += getMenuRow("", "zIndex('forward')", "forward.gif", "上移一层");
            sMenu += getMenuRow("", "zIndex('backward')", "backward.gif", "下移一层");
            sMenu += sMenuHr;
            sMenu += getMenuRow("", "insert('quote')", "quote.gif", "引用样式");
            sMenu += getMenuRow("", "insert('code')", "code.gif", "代码样式");
            height = 86;
            break;
        case "table":        // 表格菜单
            sMenu += getTableMenuRow("TableInsert");
            sMenu += getTableMenuRow("TableProp");
            sMenu += sMenuHr;
            sMenu += getTableMenuRow("TableCell");
            height = 306;
            break;
        case "form":        // 表单菜单
            sMenu += getMenuRow("", "Insermenu('time')", "FormDropdown.gif", "转向菜单");
            sMenu += getFormatMenuRow("InsertInputText", "插入输入框", "FormText.gif");
            sMenu += getFormatMenuRow("InsertTextArea", "插入输入区", "FormTextArea.gif");
            sMenu += getFormatMenuRow("InsertInputRadio", "插入单选钮", "FormRadio.gif");
            sMenu += getFormatMenuRow("InsertInputCheckbox", "插入复选钮", "FormCheckBox.gif");
            sMenu += getFormatMenuRow("InsertSelectDropdown", "插入下拉框", "FormDropdown.gif");
            sMenu += getFormatMenuRow("InsertButton", "插入按钮", "FormButton.gif");
            height = 150;
            break;
        case "zoom":        // 缩放菜单
            for (var i=0; i<aZoomSize.length; i++){
                if (aZoomSize[i]==nCurrZoomSize){
                    sMenu += getMenuRow("", "doZoom("+aZoomSize[i]+")", "checked.gif", aZoomSize[i]+"%");
                }else{
                    sMenu += getMenuRow("", "doZoom("+aZoomSize[i]+")", "space.gif", aZoomSize[i]+"%");
                }
                height += 20;
            }
            break;
        }
        
        sMenu = sMenu1 + sMenu + sMenu2;
        
        oPopDocument.open();
        oPopDocument.write("<head><link href=Editor/editor_dialog.css type=\"text/css\" rel=\"stylesheet\"></head><body scroll=\"no\"  leftmargin='0' topmargin='0' style='body:margin:0px;border:0px'onConTextMenu=\"event.returnValue=false;\">"+sMenu);
        oPopDocument.close();

        height+=2;
        if(lefter+width > document.body.clientWidth) lefter=lefter-width;
        oPopupMenu.show(lefter - leftoff - 2, topper - topoff + 22, width, height, document.body);

        return false;
    }
    // 右键菜单定义函数组结速
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // 编辑器设置 函数组开始
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // 改变编辑区高度
    function sizeChange(size){
        if (!BrowserInfo.IsIE55OrMore){
            alert("此功能需要IE5.5版本以上的支持！");
            return false;
        }
        for (var i=0; i<parent.frames.length; i++){
            if (parent.frames[i].document==self.document){
                var obj=parent.frames[i].frameElement;
                var height = parseInt(obj.offsetHeight);
                if (height+size>=100){
                    obj.height=height+size;
                }
                break;
            }
        }
    }
    // 快捷键
    function onKeyDown(event){
        <% if ShowType = 1 then %>
        //加载Html标签导航
        UpdateToolbar();
        <% End if%>
        var key = String.fromCharCode(event.keyCode).toUpperCase();

        // F2:显示或隐藏指导方针
        if (event.keyCode==113){
            showBorders();
            return false;
        }
        if (event.ctrlKey){
            // Ctrl+Enter:提交
            if (event.keyCode==10){
                doSubmit();
                return false;
            }
            // Ctrl++:增加编辑区
            if (key=="+"){
                sizeChange(300);
                return false;
            }
            // Ctrl+-:减小编辑区
            if (key=="-"){
                sizeChange(-300);
                return false;
            }
            // Ctrl+1:设计模式
            if (key=="1"){
                setMode("EDIT");
                return false;
            }
            // Ctrl+2:代码模式
            if (key=="2"){
                setMode("CODE");
                return false;
            }
            // Ctrl+3:纯文本
            if (key=="3"){
                setMode("TEXT");
                return false;
            }
            // Ctrl+4:预览
            if (key=="4"){
                setMode("VIEW");
                return false;
            }
        }
        switch(sCurrMode){
        case "VIEW":
            return true;
            break;
        case "EDIT":
            if (event.ctrlKey){
                // Ctrl+D:从Word粘贴
                if (key == "D"){
                    insert('word');
                    return false;
                }
                // Ctrl+R:查找替换
                if (key == "R"){
                    findstr();
                    return false;
                }
                // Ctrl+Z:Undo
                if (key == "Z"){
                    format('undo');
                    return false;
                }
                // Ctrl+Y:Redo
                if (key == "Y"){
                    format('redo');
                    return false;
                }
            }
            if(!event.ctrlKey && event.keyCode != 90 && event.keyCode != 89) {
                if (event.keyCode == 32 || event.keyCode == 13){
                    saveHistory()
                }
            }
            return true;
            break;
        default:
            if (event.keyCode==13){
                var sel = HtmlEdit.document.selection.createRange();
                sel.pasteHTML("<BR>");
                event.cancelBubble = true;
                event.returnValue = false;
                sel.select();
                sel.moveEnd("character", 1);
                sel.moveStart("character", 1);
                sel.collapse(false);
                return false;
            }
            // 屏蔽事件
            if (event.ctrlKey){
                // Ctrl+B,I,U
                if ((key == "B")||(key == "I")||(key == "U")){
                    return false;
                }
            }
        }
    }
    //触发焦点事件
    function onMouseUp(event,TemplateType){
    <% if ShowType = 1 then %>
    //    alert (TemplateType);
        parent.setContent('set',TemplateType);
        //加载Html标签导航
        UpdateToolbar();
    <% end if%>
    }
    //Html 标签导航
    function UpdateToolbar(){
    <% if ShowType = 1 then %>
        var ancestors = null;
        ancestors=GetAllAncestors();
        ShowObject.innerHTML='&nbsp;';
        for (var i=ancestors.length;--i>=0;){
            var el = ancestors[i];
            if (!el) continue;
            var a=document.createElement("span");
            a.href="#";
            a.el=el;
            a.editor=this;
            if (i==0){
                a.className='AncestorsMouseUp';
                EditControl=a.el;
            }
            else a.className='AncestorsStyle';
            a.onmouseover=function()
            {
                if (this.className=='AncestorsMouseUp') this.className='AncestorsMouseUpOver';
                else if (this.className=='AncestorsStyle') this.className='AncestorsMouseOver';
            };
            a.onmouseout=function()
            {
                if (this.className=='AncestorsMouseUpOver') this.className='AncestorsMouseUp';
                else if (this.className=='AncestorsMouseOver') this.className='AncestorsStyle';
            };
            a.onmousedown=function(){this.className='AncestorsMouseDown';};
            a.onmouseup=function(){this.className='AncestorsMouseUpOver';};
            a.ondragstart=YCancelEvent;
            a.onselectstart=YCancelEvent;
            a.onselect=YCancelEvent;
            a.onclick=function()
            {
                this.blur();
                SelectNodeContents(this);
                return false;
            };
            if (el.tagName.toLowerCase()!='tbody'){
                var txt='<'+el.tagName.toLowerCase();
                a.title=el.style.cssText;
                if (el.id) txt += "#" + el.id;
                if (el.className) txt += "." + el.className;
                txt=txt+'>';
                a.appendChild(document.createTextNode(txt));        
                ShowObject.appendChild(a);
            }
        }
        <%End if%>
    }
    function GetAllAncestors(){
        var p = GetParentElement();
        var a = [];
        while (p && (p.nodeType==1)&&(p.tagName.toLowerCase()!='body'))
        {
            a.push(p);
            p=p.parentNode;
        }
        a.push(HtmlEdit.document.body);
        return a;
    }
    function GetParentElement(){
        var sel=GetSelection();
        var range=CreateRange(sel);
        switch (sel.type)
        {
            case "Text":
            case "None":
                return range.parentElement();
            case "Control":
                return range.item(0);
            default:
                return HtmlEdit.document.body;
        }
    }
    function GetSelection(){
        return HtmlEdit.document.selection;
    }
    function CreateRange(sel){
        return sel.createRange();
    }
    function SelectNodeContents(Obj,pos){
        var node=Obj.el;
        EditControl=node;
        for (var i=0;i<ShowObject.children.length;i++)
        {
            if (ShowObject.children(i).className=='AncestorsMouseUp') ShowObject.children(i).className='AncestorsStyle';
        }
        HtmlEdit.focus();
        var range;
        var collapsed=(typeof pos!='undefined');
        range = HtmlEdit.document.body.createTextRange();
        range.moveToElementText(node);
        (collapsed) && range.collapse(pos);
        range.select();
    }
    // 显示无模式对话框
    function ShowDialog(url, width, height, optValidate){
        if (!    validateMode())    return;
        HtmlEdit.focus();
        var range = HtmlEdit.document.selection.createRange();
        var arr = showModalDialog(url, window, "dialogWidth:" + width + "px;dialogHeight:" + height + "px;help:no;scroll:yes;status:no");
        if (arr != null){
            range.pasteHTML(arr);
        }
      HtmlEdit.focus();
    }
    // 显示预览
    function cleanHtml(){
        var fonts = HtmlEdit.document.body.all.tags("FONT");
        var curr;
        for (var i = fonts.length - 1; i >= 0; i--) {
            curr = fonts[i];
            if (curr.style.backgroundColor == "#ffffff") curr.outerHTML    = curr.innerHTML;
        }
    }
    // 是否选中指定类型的控件
    function isControlSelected(tag){
        if (HtmlEdit.document.selection.type == "Control") {
            var oControlRange = HtmlEdit.document.selection.createRange();
            if (oControlRange(0).tagName.toUpperCase() == tag) {
                return true;
            }    
        }
        return false;
    }
    // 判断是否在编辑状态
    function validateMode(){
        if (EditMode) return true;
        alert("请先点编辑器下方的“编辑”按钮，进入“编辑”状态，然后再使用系统编辑功能!");
        HtmlEdit.focus();
        return false;
    }
    // 字体处理
    function format(what,opt){
        if (!validateMode()) return;
        if (opt=="removeFormat"){
            what=opt;
            opt=null;
        }
        if (opt==null) HtmlEdit.document.execCommand(what);
        else HtmlEdit.document.execCommand(what,"",opt);
        pureText = false;
        HtmlEdit.focus();
    }
    //暂时解决一下文本源码，复制粘贴的问题。
    function format2(what,opt){
        if (opt=="removeFormat"){
            what=opt;
            opt=null;
        }
        if (opt==null) HtmlEdit.document.execCommand(what);
        else HtmlEdit.document.execCommand(what,"",opt);
        pureText = false;
        HtmlEdit.focus();
    }
    // 修正Undo/Redo
    var history = new Object;
    history.data = [];
    history.position = 0;
    history.bookmark = [];
    // 保存历史
    function saveHistory() {
        if (bEditMode){
            if (history.data[history.position] != HtmlEdit.document.body.innerHTML){
                var nBeginLen = history.data.length;
                var nPopLen = history.data.length - history.position;
                for (var i=1; i<nPopLen; i++){
                    history.data.pop();
                    history.bookmark.pop();
                }

                history.data[history.data.length] = HtmlEdit.document.body.innerHTML;

                if (HtmlEdit.document.selection.type != "Control"){
                    history.bookmark[history.bookmark.length] = HtmlEdit.document.selection.createRange().getBookmark();
                } else {
                    var oControl = HtmlEdit.document.selection.createRange();
                    history.bookmark[history.bookmark.length] = oControl[0];
                }

                if (nBeginLen!=0){
                    history.position++;
                }
            }
        }
    }
    // 初始历史
    function initHistory() {
        history.data.length = 0;
        history.bookmark.length = 0;
        history.position = 0;
    }
    // 返回历史
    function goHistory(value) {
        saveHistory();
        // undo
        if (value == -1){
            if (history.position > 0){
                HtmlEdit.document.body.innerHTML = history.data[--history.position];
                setHistoryCursor();
            }
        // redo
        } else {
            if (history.position < history.data.length -1){
                HtmlEdit.document.body.innerHTML = history.data[++history.position];
                setHistoryCursor();
            }
        }
    }
    // 设置当前书签
    function setHistoryCursor() {
        if (history.bookmark[history.position]){
            r = HtmlEdit.document.body.createTextRange()
            if (history.bookmark[history.position] != "[object]"){
                if (r.moveToBookmark(history.bookmark[history.position])){
                    r.collapse(false);
                    r.select();
                }
            }
        }
    }
    // End Undo / Redo Fix
    function setMode(NewMode){
        if (!BrowserInfo.IsIE55OrMore){
            if ((NewMode=="CODE") || (NewMode=="EDIT") || (NewMode=="VIEW")){
                alert("HTML编辑模式需要IE5.5版本以上的支持！");
                return false;
            }
        }
        if (NewMode=="TEXT"){
            if (sCurrMode==ModeEdit.value){
                if (!confirm("警告！切换到纯文本模式会丢失您所有的HTML格式，您确认切换吗？")){
                    return false;
                }
            }
        }
        var sBody = "";
        switch(sCurrMode){
        case "CODE":
            if (NewMode=="TEXT"){
                HtmlEdit_Temp_HTML.innerHTML = HtmlEdit.document.body.innerText;
                sBody = HtmlEdit_Temp_HTML.innerText;
            }else{                
                sBody = HtmlEdit.document.body.innerText;
            }
            break;
        case "TEXT":
            sBody = HtmlEdit.document.body.innerText;
            sBody = HTMLEncode(sBody);
            break;
        case "EDIT":
            if (NewMode=="TEXT"){
                sBody = HtmlEdit.document.body.innerText;
            }else{
                sBody = HtmlEdit.document.body.innerHTML;
            }
            break;
        case "VIEW":
            if (NewMode=="TEXT"){
                sBody = HtmlEdit.document.body.innerText;
            }else{
                sBody = HtmlEdit.document.body.innerHTML;
            }
            break;
        default:
                
            sBody = ContentEdit.value;;
            break;        
        }
        sCurrMode = NewMode;
        ModeEdit.value = NewMode;
        setHTML(sBody);
    }
    // 替换特殊字符
    function HTMLEncode(text){
        text = text.replace(/&/g, "&amp;") ;
        text = text.replace(/"/g, "&quot;") ;
        text = text.replace(/</g, "&lt;") ;
        text = text.replace(/>/g, "&gt;") ;
        text = text.replace(/'/g, "&#146;") ;
        text = text.replace(/\ /g,"&nbsp;");
        text = text.replace(/\n/g,"<br>");
        text = text.replace(/\t/g,"&nbsp;&nbsp;&nbsp;&nbsp;");
        return text;
    }
    // 取编辑器的内容
    function getHTML(){
        var html;
        if((sCurrMode=="EDIT")||(sCurrMode=="VIEW")){
            html = HtmlEdit.document.body.innerHTML;
        }else{
            html = HtmlEdit.document.body.innerText;
        }
        if (sCurrMode!="TEXT"){
            if ((html.toLowerCase()=="<p>&nbsp;</p>")||(html.toLowerCase()=="<p></p>")){
                html = "";
            }
        }
        return html;
    }
    // 设置编辑器的内容
    function setHTML(html){
        ContentEdit.value = html;
        switch (sCurrMode){
        case "CODE":
            setMode0.src="Editor/images/Editor.gif";
            setMode1.src="Editor/images/html2.gif";
            setMode2.src="Editor/images/browse.gif";
            setMode3.src="Editor/images/Text.gif";
            HtmlEdit.document.designMode="on";
            HtmlEdit.document.open();
            HtmlEdit.document.write(edithead);
            HtmlEdit.document.write(Resumeblank(html));
            HtmlEdit.document.close();
            HtmlEdit.document.body.innerText=Resumeblank(html);    
            HtmlEdit.document.body.contentEditable="true";
            CurrentMode=1;
            EditMode=false;
            SourceMode=true;
            PreviewMode=false;
            bEditMode=true;
            break;
        case "EDIT":
            <%if ShowType <> 1 Then%>
                setMode0.src="Editor/images/Editor2.gif";
                setMode1.src="Editor/images/html.gif";
                setMode2.src="Editor/images/browse.gif";
                setMode3.src="Editor/images/Text.gif";
            <%End if%>
            HtmlEdit.document.designMode="on";
            HtmlEdit.document.open();
            HtmlEdit.document.write(edithead);
            HtmlEdit.document.write(html);
            HtmlEdit.document.close();    
            doZoom(nCurrZoomSize);
            CurrentMode=0;
            EditMode=true;
            SourceMode=false;
            PreviewMode=false;
            bEditMode=true;
            break;    
        case "TEXT":
            setMode0.src="Editor/images/Editor.gif";
            setMode1.src="Editor/images/html.gif";
            setMode2.src="Editor/images/browse.gif";
            setMode3.src="Editor/images/Text2.gif";
            HtmlEdit.document.designMode="on";
            HtmlEdit.document.open();
            HtmlEdit.document.write(edithead);
            HtmlEdit.document.write(Resumeblank(html));
            HtmlEdit.document.body.innerText=html;
            HtmlEdit.document.body.contentEditable="true";
            HtmlEdit.document.close();
            CurrentMode=1
            EditMode=false;
            SourceMode=true;
            PreviewMode=false;
            bEditMode=true;
            borderShown = "0";
            break;
        case "VIEW":
            setMode0.src="Editor/images/Editor.gif";
            setMode1.src="Editor/images/html.gif";
            setMode2.src="Editor/images/browse2.gif";
            setMode3.src="Editor/images/Text.gif";
            cleanHtml();
            CurrentMode=3;
            HtmlEdit.document.designMode="off";
            HtmlEdit.document.open();
            HtmlEdit.document.write(edithead);
            HtmlEdit.document.write(Resumeblank(html));
            HtmlEdit.document.body.contentEditable="false";
            HtmlEdit.document.close();
            EditMode=false;
            SourceMode=false;
            PreviewMode=false;
            bEditMode=false;
            break;
        default:
            alert("错误参数调用！");
            break;
        }

        HtmlEdit.document.onkeydown = new Function("return onKeyDown(HtmlEdit.event);");
        HtmlEdit.document.oncontextmenu=new Function("return showContextMenu(HtmlEdit.event);");
        HtmlEdit.document.onmouseup = new Function('return onMouseUp(HtmlEdit.event,<%=Clng(TemplateType)%>);');

        if (borderShown != "0" && EditMode) {
            borderShown = "0";
            showBorders();
        }
        initHistory();
        HtmlEdit.focus();
    }
    // 显示或隐藏指导方针
    var borderShown = 0;
    function showBorders() {
        if (!document.all){
            setTimeout("showBorders()",1000);
            return;
        }
        if (!validateMode()) return;
        
        var allForms = HtmlEdit.document.body.getElementsByTagName("FORM");
        var allInputs = HtmlEdit.document.body.getElementsByTagName("INPUT");
        var allTables = HtmlEdit.document.body.getElementsByTagName("TABLE");
        var allLinks = HtmlEdit.document.body.getElementsByTagName("A");

        // 表单
        for (a=0; a < allForms.length; a++) {
            if (borderShown == "0") {
                allForms[a].runtimeStyle.border = "1px dotted #FF0000"
            } else {
                allForms[a].runtimeStyle.cssText = ""
            }
        }

        // Input Hidden类
        for (b=0; b < allInputs.length; b++) {
            if (borderShown == "0") {
                if (allInputs[b].type.toUpperCase() == "HIDDEN") {
                    allInputs[b].runtimeStyle.border = "1px dashed #000000"
                    allInputs[b].runtimeStyle.width = "15px"
                    allInputs[b].runtimeStyle.height = "15px"
                    allInputs[b].runtimeStyle.backgroundColor = "#FDADAD"
                    allInputs[b].runtimeStyle.color = "#FDADAD"
                }
            } else {
                if (allInputs[b].type.toUpperCase() == "HIDDEN")
                    allInputs[b].runtimeStyle.cssText = ""
            }
        }

        // 表格
        for (i=0; i < allTables.length; i++) {
                if (borderShown == "0") {
                    allTables[i].runtimeStyle.border = "1px dotted #BFBFBF"
                } else {
                    allTables[i].runtimeStyle.cssText = ""
                }

                allRows = allTables[i].rows
                for (y=0; y < allRows.length; y++) {
                    allCellsInRow = allRows[y].cells
                        for (x=0; x < allCellsInRow.length; x++) {
                            if (borderShown == "0") {
                                allCellsInRow[x].runtimeStyle.border = "1px dotted #BFBFBF"
                            } else {
                                allCellsInRow[x].runtimeStyle.cssText = ""
                            }
                        }
                }
        }

        // 链接 A
        for (a=0; a < allLinks.length; a++) {
            if (borderShown == "0") {
                if (allLinks[a].href.toUpperCase() == "") {
                    allLinks[a].runtimeStyle.borderBottom = "1px dashed #000000"
                }
            } else {
                allLinks[a].runtimeStyle.cssText = ""
            }
        }

        if (borderShown == "0") {
            borderShown = "1"
        } else {
            borderShown = "0"
        }

        scrollUp()
    }

    // 返回页面最上部
    function scrollUp() {
        HtmlEdit.scrollBy(0,0);
    }
    // 保存验证
    function save()
    {
        if (CurrentMode==0){
        //编辑器嵌入其他网页时使用下面这一句（请将form1改成相应表单名）
        parent.myform.Content.value=HtmlEdit.document.body.innerHTML;
        //单独打开编辑器时使用下面这一句（请将form1改成相应表单名）  
        //  self.opener.form1.content.value+=HtmlEdit.document.body.innerHTML;
        }
        else if(CurrentMode==1){
        //编辑器嵌入其他网页时使用下面这一句（请将form1改成相应表单名）
        parent.myform.Content.value=HtmlEdit.document.body.innerText;
        //单独打开编辑器时使用下面这一句（请将form1改成相应表单名）  
        //  self.opener.form1.content.value+=HtmlEdit.document.body.innerText;
        }
        else{
            alert("预览状态不能保存！请先回到编辑状态后再保存");
        }
        HtmlEdit.focus();
    }
    // 检测当前是否在预览模式
    function isModeView(){
        if (sCurrMode=="VIEW"){
            alert("预览时不允许设置编辑区内容。");
            return true;
        }
        return false;
    }
    // 在当前文档位置插入.
    function insertHTML(html) {
        HtmlEdit.focus();
        if (isModeView()) return false;
        if (HtmlEdit.document.selection.type.toLowerCase() != "none"){
            HtmlEdit.document.selection.clear() ;
        }
        if (sCurrMode!="EDIT"){
            html=HTMLEncode(html);
        }
        HtmlEdit.document.selection.createRange().pasteHTML(html) ; 
    }
    //新加入功能
    //插入表单表单
    function Insergongneng(what){
        if (!    validateMode())    return;
        HtmlEdit.focus();
        var range = HtmlEdit.document.selection.createRange();
        var ran = HtmlEdit.document.selection.createRange("").text;
        switch(what){
        case "input":
            range.pasteHTML('<INPUT value='+ran+'>');
            break;
        case "textarea":
            range.pasteHTML('<TEXTAREA>'+ran+'</TEXTAREA>');
            break;
        case "radio":
            range.pasteHTML('<INPUT type=radio>');
            break;
        case "checkbox":
            range.pasteHTML('<INPUT type=checkbox>');
            break;
        case "bottom":
            range.pasteHTML('<BUTTON>'+ran+'</BUTTON>');
            break;
        }
        HtmlEdit.focus();
    }
    // 插入下拉菜单
    function Insermenu(id){
        HtmlEdit.focus();
        if (!    validateMode())    return;
        var range = HtmlEdit.document.selection.createRange();
        var ran = HtmlEdit.document.selection.createRange("").text;
        var arr = showModalDialog("Editor/editor_insmenu.asp?ChannelID=<%=ChannelID%>&id="+id, "", "dialogWidth:285pt;dialogHeight:186pt;help:0;status:0");

        if (arr != null){
            range.pasteHTML(arr);
        }
        HtmlEdit.focus();
    }
    // 插入特殊符号
    function Insertlr(filename,wwid,whei,myid){
        if (!    validateMode())    return;
        HtmlEdit.focus();
        var range = HtmlEdit.document.selection.createRange();
        var arr = showModalDialog("Editor/"+filename+"?ChannelID=<%=ChannelID%>&id="+myid, window, "dialogWidth:"+wwid+"pt;dialogHeight:"+whei+"pt;help:0;status:0");
        if (arr != null){
            range.pasteHTML(arr);
        }
        HtmlEdit.focus();
    }
    // 缩放操作
    var  nCurrZoomSize = 100;
    var  aZoomSize = new Array(10, 25, 50, 75, 100, 150, 200, 500);
    // 显示框架比例
    function doZoom(size) {
        HtmlEdit.document.body.runtimeStyle.zoom = size + "%";
        nCurrZoomSize = size;
    }
    // 图片属性 上下层
    function zIndex(action){
        var objReference    = null;
        var RangeType        = HtmlEdit.document.selection.type;
        if (RangeType != "Control") return;
        var selectedRange    = HtmlEdit.document.selection.createRange();
        for (var i=0; i<selectedRange.length; i++){
            objReference = selectedRange.item(i);
            if (action=='forward'){
                objReference.style.zIndex  +=1;
            }else{
                objReference.style.zIndex  -=1;
            }
            objReference.style.position='absolute';
        }
        HtmlEdit.content = false;
    }
    // 图片左右环绕
    function imgalign(align){

    if (!validateMode()) return;
    HtmlEdit.focus();

    var oControl;
    var oSeletion;
    var sRangeType;
    oSelection = HtmlEdit.document.selection.createRange();
    sRangeType = HtmlEdit.document.selection.type;

    if (sRangeType == "Control") {
        if (oSelection.item(0).tagName == "IMG"){
               
            oControl = oSelection.item(0)
            oControl.align = align;
        }
    }

    HtmlEdit.focus();

    }

    //普通标签
    function InsertLabel(label){
        HtmlEdit.focus();
        var range = HtmlEdit.document.selection.createRange();
        if (label=="ShowTopUser"){
            label=label+"("+prompt("请输入显示注册用户列表的数量.","5")+")"
        }
        range.pasteHTML("{$"+label+"}");
        HtmlEdit.focus();
    }
    //函数标签调用
    function FunctionLabel(url,width,height){
        HtmlEdit.focus();
        var range = HtmlEdit.document.selection.createRange();
        var label = showModalDialog(url, "", "dialogWidth:"+width+"px; dialogHeight:"+height+"px; help: no; scroll:no; status: no"); 
        if (label != null){
            range.pasteHTML(label);
        }
        HtmlEdit.focus();
    }
    //超级函数标签调用
    function SuperFunctionLabel(url,label,title,ModuleType,ChannelShowType,iwidth,iheight){
        HtmlEdit.focus();
        var range = HtmlEdit.document.selection.createRange();
        var label = showModalDialog(url+"?ChannelID=<%=ChannelID%>&Action=Add&LabelName="+label+"&Title="+title+"&ModuleType="+ModuleType+"&ChannelShowType="+ChannelShowType+"&InsertTemplate=0", "", "dialogWidth:"+iwidth+"px; dialogHeight:"+iheight+"px; help: no; scroll:yes; status: yes"); 
        if (label != null){
            range.pasteHTML(label);
        }
        HtmlEdit.focus();
    }
    // 编辑器设置 函数组结速
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    // 插入文件函 数组开始
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // 插入特殊对象
    function insert(what) {
        if (!validateMode()) return;
        HtmlEdit.focus();
        var range = HtmlEdit.document.selection.createRange();
        var RangeType = HtmlEdit.document.selection.type;
        switch(what){
        case "excel":        // 插入EXCEL表格
            insertHTML("<object classid='clsid:0002E510-0000-0000-C000-000000000046' id='Spreadsheet1' codebase='file:\\Bob\software\office2000\msowc.cab' width='100%' height='250'><param name='HTMLURL' value><param name='HTMLData' value='&lt;html xmlns:x=&quot;urn:schemas-microsoft-com:office:excel&quot;xmlns=&quot;http://www.w3.org/TR/REC-html40&quot;&gt;&lt;head&gt;&lt;style type=&quot;text/css&quot;&gt;&lt;!--tr{mso-height-source:auto;}td{black-space:nowrap;}.wc4590F88{black-space:nowrap;font-family:宋体;mso-number-format:General;font-size:auto;font-weight:auto;font-style:auto;text-decoration:auto;mso-background-source:auto;mso-pattern:auto;mso-color-source:auto;text-align:general;vertical-align:bottom;border-top:none;border-left:none;border-right:none;border-bottom:none;mso-protection:locked;}--&gt;&lt;/style&gt;&lt;/head&gt;&lt;body&gt;&lt;!--[if gte mso 9]&gt;&lt;xml&gt;&lt;x:ExcelWorkbook&gt;&lt;x:ExcelWorksheets&gt;&lt;x:ExcelWorksheet&gt;&lt;x:OWCVersion&gt;9.0.0.2710&lt;/x:OWCVersion&gt;&lt;x:Label Style='border-top:solid .5pt silver;border-left:solid .5pt silver;border-right:solid .5pt silver;border-bottom:solid .5pt silver'&gt;&lt;x:Caption&gt;Microsoft Office Spreadsheet&lt;/x:Caption&gt; &lt;/x:Label&gt;&lt;x:Name&gt;Sheet1&lt;/x:Name&gt;&lt;x:WorksheetOptions&gt;&lt;x:Selected/&gt;&lt;x:Height&gt;7620&lt;/x:Height&gt;&lt;x:Width&gt;15240&lt;/x:Width&gt;&lt;x:TopRowVisible&gt;0&lt;/x:TopRowVisible&gt;&lt;x:LeftColumnVisible&gt;0&lt;/x:LeftColumnVisible&gt; &lt;x:ProtectContents&gt;False&lt;/x:ProtectContents&gt; &lt;x:DefaultRowHeight&gt;210&lt;/x:DefaultRowHeight&gt; &lt;x:StandardWidth&gt;2389&lt;/x:StandardWidth&gt; &lt;/x:WorksheetOptions&gt; &lt;/x:ExcelWorksheet&gt;&lt;/x:ExcelWorksheets&gt; &lt;x:MaxHeight&gt;80%&lt;/x:MaxHeight&gt;&lt;x:MaxWidth&gt;80%&lt;/x:MaxWidth&gt;&lt;/x:ExcelWorkbook&gt;&lt;/xml&gt;&lt;![endif]--&gt;&lt;table class=wc4590F88 x:str&gt;&lt;col width=&quot;56&quot;&gt;&lt;tr height=&quot;14&quot;&gt;&lt;td&gt;&lt;/td&gt;&lt;/tr&gt;&lt;/table&gt;&lt;/body&gt;&lt;/html&gt;'> <param name='DataType' value='HTMLDATA'> <param name='AutoFit' value='0'><param name='DisplayColHeaders' value='-1'><param name='DisplayGridlines' value='-1'><param name='DisplayHorizontalScrollBar' value='-1'><param name='DisplayRowHeaders' value='-1'><param name='DisplayTitleBar' value='-1'><param name='DisplayToolbar' value='-1'><param name='DisplayVerticalScrollBar' value='-1'> <param name='EnableAutoCalculate' value='-1'> <param name='EnableEvents' value='-1'><param name='MoveAfterReturn' value='-1'><param name='MoveAfterReturnDirection' value='0'><param name='RightToLeft' value='0'><param name='ViewableRange' value='1:65536'></object>");
            break;
        case "nowdate":        // 插入当前系统日期
            var d = new Date();
            insertHTML(d.toLocaleDateString());
            break;
        case "nowtime":        // 插入当前系统时间
            var d = new Date();
            insertHTML(d.toLocaleTimeString());
            break;
        case "br":            // 插入换行符        
            insertHTML("<br>")
            break;
        case "code":        // 代码片段样式
            insertHTML('<table width=95% border="0" align="Center" cellpadding="6" cellspacing="0" style="border: 1px Dotted #CCCCCC; TABLE-LAYOUT: fixed"><tr><td bgcolor=#FDFDDF style="WORD-WRAP: break-word"><font style="color: #990000;font-weight:bold">以下是代码片段：</font><br>'+HTMLEncode(range.text)+'</td></tr></table>');
            break;
        case "quote":        // 引用片段样式
            insertHTML('<table width=95% border="0" align="Center" cellpadding="6" cellspacing="0" style="border: 1px Dotted #CCCCCC; TABLE-LAYOUT: fixed"><tr><td bgcolor=#F3F3F3 style="WORD-WRAP: break-word"><font style="color: #990000;font-weight:bold">以下是引用片段：</font><br>'+HTMLEncode(range.text)+'</td></tr></table>');
            break;
        case "big": // 字体变大
            insertHTML("<big>" + range.text + "</big>");
            break;
        case "small":    // 字体变小
            insertHTML("<small>" + range.text + "</small>");
            break;
        case "fgcolor": //字体颜色
            if (RangeType != "Text"){
                alert("请先选择一段文字！");
                return;
            }
            var arr = showModalDialog("Editor/editor_selcolor.asp?ChannelID=<%=ChannelID%>", "", "dialogWidth:18.5em; dialogHeight:17.5em; help: no; scroll: no; status: no");
            if (arr != null) format('forecolor', arr);
            else HtmlEdit.focus();
            break;
        case "fgbgcolor": //字体背景色
            if (RangeType != "Text"){
               alert("请先选择一段文字！");
               return;
            }
            var arr = showModalDialog("Editor/editor_selcolor.asp?ChannelID=<%=ChannelID%>", "", "dialogWidth:18.5em; dialogHeight:17.5em; help: no; scroll: no; status: no");
            if (arr != null){
                range.pasteHTML("<span style='background-color:"+arr+"'>"+range.text+"</span> ");
                range.select();
            }
            HtmlEdit.focus();
            break;
        case "hr": // 水平线
            var arr = showModalDialog("Editor/editor_inserthr.asp?ChannelID=<%=ChannelID%>", "", "dialogWidth:30em; dialogHeight:12em; help: no; scroll: no; status: no"); 
            if (arr != null){
                range.pasteHTML(arr);
            }
            HtmlEdit.focus();
            break;
        case "page": //分页
            if(range.text!=""){
               alert("请不要选择任何文本");
            }
            else{
               range.text="\n\n[NextPage]\n\n";
               parent.selectPaginationType();
            }
            break;
        case "word": //word粘贴
            HtmlEdit.document.execCommand("Paste",false);
            var editBody=HtmlEdit.document.body;
            for(var intLoop=0;intLoop<editBody.all.length;intLoop++){
                el=editBody.all[intLoop];
                el.removeAttribute("className","",0);
                el.removeAttribute("style","",0);
                el.removeAttribute("font","",0);
            }
            var html=HtmlEdit.document.body.innerHTML;
            html=html.replace(/<o:p>&nbsp;<\/o:p>/g,"");
            html=html.replace(/o:/g,"");
            html=html.replace(/<font>/g, "");
            html=html.replace(/<FONT>/g, "");
            html=html.replace(/<span>/g, "");
            html=html.replace(/<SPAN>/g, "");
            html=html.replace(/<SPAN lang=EN-US>/g, "");
            html=html.replace(/<P>/g, "");
            html=html.replace(/<\/P>/g, "");
            html=html.replace(/<\/SPAN>/g, "");
            HtmlEdit.document.body.innerHTML = html;
            format('selectall');
            format('RemoveFormat');
            break;
        case "calculator": // 计算器
            var arr = showModalDialog("Editor/editor_calculator.asp?ChannelID=<%=ChannelID%>", "", "dialogWidth:205px; dialogHeight:230px; status:0;help:0");
            if (arr != null){
                var ss;
                ss=arr.split("*")
                a=ss[0];
                b=ss[1];
                var str1;
                str1=""+a+""
                range.pasteHTML(str1);
            }
            HtmlEdit.focus();
            break;
        case "help": //帮助
            var arr = showModalDialog("Editor/editor_help.asp?ChannelID=<%=ChannelID%>", "", "dialogWidth:580px; dialogHeight:460px; help: no; scroll: no; status: no");
            break;
        case "FIELDSET": // 栏目框
            var arr = showModalDialog("Editor/editor_fieldset.asp?ChannelID=<%=ChannelID%>", "", "dialogWidth:25em; dialogHeight:12.5em; help: no; scroll: no; status: no");
            if (arr != null){
                range.pasteHTML(arr);
            }
            HtmlEdit.focus();
            break;
        case "iframe": //内联页
            var arr = showModalDialog("Editor/editor_insertiframe.asp?ChannelID=<%=ChannelID%>", "", "dialogWidth:30em; dialogHeight:12em; help: no; scroll: no; status: no");  
            if (arr != null){
                range.pasteHTML(arr);
            }
            HtmlEdit.focus();
            break;
        case "insermarquee": // 滚动文本
            var arr = showModalDialog("Editor/editor_marquee.asp?ChannelID=<%=ChannelID%>", "", "dialogWidth:275pt;dialogHeight:100pt;help:0;status:0");  
            if (arr != null){
                range.pasteHTML(arr);
            }
            HtmlEdit.focus();
            break;
        case "inseremot": // 插入表情
            var arr = showModalDialog("Editor/editor_emot.asp?ChannelID=<%=ChannelID%>", "", "dialogWidth:400px;dialogHeight:400px;help:0;status:0");  
            if (arr != null){
                range.pasteHTML(arr);
            }
            HtmlEdit.focus();
            break;
        case "calljsad": // 插入JS标签
            var arr = showModalDialog("Editor/editor_ad.asp?ChannelID=<%=ChannelID%>", "", "dialogWidth:200px;dialogHeight:200px;help:0;status:0");  
            if (arr != null){
                range.pasteHTML(arr);
            }
            HtmlEdit.focus();
            break;
        case "Label": // 插入标签
            var arr = showModalDialog("Editor/editor_tree.asp?ChannelID=<%=ChannelID%>", "", "dialogWidth:230pt;dialogHeight:500px;help:0;status:0");  
            if (arr != null){
                range.pasteHTML(arr);
            }
            HtmlEdit.focus();
            break;
        case "editLabel": // 编辑标签
            var oControl;
            var oSeletion;
            var sRangeType;
            var zzz="";
            oSelection = HtmlEdit.document.selection.createRange();
            sRangeType = HtmlEdit.document.selection.type;

            if (sRangeType == "Control") {
                if (oSelection.item(0).tagName == "IMG"){
                    oControl = oSelection.item(0);
                    zzz= oControl.zzz;
                }
                var arr = showModalDialog("Editor/editor_label.asp?ChannelID=<%=ChannelID%>&Action=Modify&Title=修改标签&editLabel="+zzz+"", window, "dialogWidth:" + 800 + "px;dialogHeight:" + 600 + "px;help:no;scroll:yes;status:no");
                if (arr != null){
                    oControl.zzz=arr
                }
            }else{
                alert("不能获取到html对象！");
            }
            HtmlEdit.focus();
            break;
        case "InsertEQ": // 公式
            var arr = showModalDialog("Editor/editor_inserteq.asp?ChannelID=<%=ChannelID%>", "", "dialogWidth:40em; dialogHeight:20em; status:0;help:0");
            if (arr != null){
                var ss;
                ss=arr.split("*")
                a=ss[0];
                b=ss[1];
                var str1;
                str1="<applet codebase='./' code='webeq3.ViewerControl' WIDTH=320 HEIGHT=100>"
                str1=str1+"<PARAM NAME='parser' VALUE='mathml'><param name='color' value='"+b+"'><PARAM NAME='size' VALUE='18'>"
                str1=str1+"<PARAM NAME=eq id=eq VALUE='"+a+"'></applet>"
                range.pasteHTML(str1);
            }
            HtmlEdit.focus();
            break;
        case "InstallEQ": // 安装公式
            window.open ("Editor/editor_inserteq.asp?ChannelID=<%=ChannelID%>&Action=Install", "", "height=200, width=300,left="+(screen.AvailWidth-300)/2+",top="+(screen.AvailHeight-200)/2+", toolbar=no, menubar=no, scrollbars=no, resizable=no,location=no, status=no")
            break;
        case "batchpic": //批量上传图片
            var arr = showModalDialog("Editor/editor_insertpic.asp?ChannelID=<%=ChannelID%>&ShowType=<%=ShowType%><%If Anonymous = 1 Then Response.Write "&Anonymous=1"%>", "", "dialogWidth:800px; dialogHeight:470px; help: no; scroll: yes; status: no");  
            if (arr != null){
                var ss=arr.split("$$$");
                range.pasteHTML(ss[0]);
                for(var i=1;i<=ss[1];i++){
                    if (ss[i+1]!="None"){
                        parent.AddItem(ss[i+1]);
                    }
                }
            }
            HtmlEdit.focus();
            break;
        case "pic": //上传图片
            var arr = showModalDialog("Editor/editor_Modifypic.asp?ChannelID=<%=ChannelID%>&ShowType=<%=ShowType%><%If Anonymous = 1 Then Response.Write "&Anonymous=1"%>", window, "dialogWidth:" + 500 + "px;dialogHeight:" + 540 + "px;help:no;scroll:yes;status:no");
            if (arr != null){
                var ss=arr.split("$$$");
                for(var i=1;i<=ss[0];i++){
                    if (ss[i]!=""){
                        parent.AddItem(ss[i]);
                    }
                }
            }
            HtmlEdit.focus();
            break;
        case "swf": //上传swf
            var arr = showModalDialog("Editor/editor_insertflash.asp?ChannelID=<%=ChannelID%>&ShowType=<%=ShowType%><%If Anonymous = 1 Then Response.Write "&Anonymous=1"%>", "", "dialogWidth:530px; dialogHeight:400px; help: no; scroll: yes; status: no"); 
            if (arr != null){
                var ss=arr.split("$$$");
                range.pasteHTML(ss[0]);
                if (ss[1]!="None"){
                    parent.AddItem(ss[1]);
                }
            }
            HtmlEdit.focus();
            break;
        case "wmv": //上传 wmv
            var arr = showModalDialog("Editor/editor_insertmedia.asp?ChannelID=<%=ChannelID%>&ShowType=<%=ShowType%><%If Anonymous = 1 Then Response.Write "&Anonymous=1"%>", "", "dialogWidth:530px; dialogHeight:500px; help: no; scroll: yes; status: no");
            if (arr != null){
                var ss=arr.split("$$$");
                range.pasteHTML(ss[0]);
                if (ss[1]!="None"){
                    parent.AddItem(ss[1]);
                }
            }
            HtmlEdit.focus();
            break;
        case "Attribute": //代码框编辑
            var arr = showModalDialog("Editor/editor_Attribute.asp?ChannelID=<%=ChannelID%>", window, "dialogWidth:" + 600 + "px;dialogHeight:" + 270 + "px;help:no;scroll:yes;status:no");
            showBorders();
            showBorders();
            HtmlEdit.focus();
            break;
        case "rm": //上传 rm
            var arr = showModalDialog("Editor/editor_insertrm.asp?ChannelID=<%=ChannelID%>&ShowType=<%=ShowType%><%If Anonymous = 1 Then Response.Write "&Anonymous=1"%>", "", "dialogWidth:500px; dialogHeight:500px; help: no; scroll: yes; status: no");  
            if (arr != null){
                var ss=arr.split("$$$");
                range.pasteHTML(ss[0]);
                if (ss[1]!="None"){
                    parent.AddItem(ss[1]);
                }
            }
            HtmlEdit.focus();
            break;
        case "fujian": //上传附件	
            var arr = showModalDialog("Editor/editor_insertfujian.asp?ChannelID=<%=ChannelID%>&ShowType=<%=ShowType%><%If Anonymous = 1 Then Response.Write "&Anonymous=1"%>", "", "dialogWidth:31em; dialogHeight:12em; help: no; scroll: no; status: no"); 
            if (arr != null){
                var ss=arr.split("$$$");
                range.pasteHTML(ss[0]);
                if (ss[1]!="None"){
                    parent.AddItem(ss[1]);
                }
            }
            HtmlEdit.focus();
            break;
        case "title":  // 设置标题
            if (RangeType != "Text"){
                alert("请先选择一段文字！");
                return;
            }
            parent.document.myform.Title.value=range.text;
            break;
        case "keyword" :// 设置关键字
            if (RangeType != "Text"){
                alert("请先选择一段文字！");
            }
            if (parent.document.myform.Keyword.value==""){
                parent.document.myform.Keyword.value=range.text;
            }
            else{
                parent.document.myform.Keyword.value+="|"+range.text;
            }
            break;
        case "ProductName":
            if (RangeType != "Text"){
                alert("请先选择一段文字！");
                return;
            }
            parent.document.myform.ProductName.value=range.text;
            break;
        case "Intro":
            if (RangeType != "Text"){
                alert("请先选择一段文字！");
                return;
            }
            parent.document.myform.Intro.value=range.text;
            break;
        case "ProductIntro":
            if (RangeType != "Text"){
                alert("请先选择一段文字！");
                return;
            }
            parent.document.myform.ProductIntro.value=range.text;
            break;
        case "ReplaceLabel":
            var oControl;
            var oSeletion;
            var sRangeType;
            oSelection = HtmlEdit.document.selection.createRange();
            sRangeType = HtmlEdit.document.selection.type;
            if (sRangeType == "Control") {
                oControl = oSelection.item(0);
                var arr = showModalDialog("Editor/editor_tree.asp?ChannelID=<%=ChannelID%>", "", "dialogWidth:230pt;dialogHeight:500px;help:0;status:0");  
                if (arr != null){
                    oControl.outerHTML=arr
                }
            }else{
                alert("不能获取到html对象！");
            }
            HtmlEdit.focus();
            break;
        case "CreateLink"://链接字体
            var arr = showModalDialog("Editor/editor_CreateLink.asp?ChannelID=<%=ChannelID%>&LinkName="+range.text+"", window, "dialogWidth:450px; dialogHeight:450px; help: no; scroll: no; status: no");
            if (arr != null){
                insertHTML(arr);
            }
            HtmlEdit.focus();
            break;
        case "pagetitle": //内容页的分页标签
            var arr=showModalDialog("Editor/editor_Pagetitle.asp?ChannelID=<%=ChannelID%>","","dialogWidth:400pt;dialogHeight:80px;help:0;status:0");
            
            if(arr!=null){
                range.pasteHTML(arr);
				parent.selectPaginationType();
            }
            HtmlEdit.focus();
            break;
        case "copypagetitle":
            if (RangeType != "Text"){
               alert("请先选择一段文字！");
               return;
            }else{
               range.text="[NextPage" + range.text + "]\n\n" + range.text + "";
               parent.selectPaginationType();
            }
            break;
        case "FilterCode":
            var arr=showModalDialog("Editor/editor_FilterCode.asp?ChannelID=<%=ChannelID%>","","dialogWidth:400pt;dialogHeight:340px;help:0;status:0");
            if(arr!=null){
                var ss=arr.split(",");
                var strhtml=HtmlEdit.document.body.innerHTML
                if (ss[0] == "true"){
                    strhtml = ScriptHtml(strhtml, "Iframe", 2,"")
                }
                if (ss[1] == "true"){
                    strhtml = ScriptHtml(strhtml, "Object", 2,"")
                }
                if (ss[2] == "true"){
                    strhtml = ScriptHtml(strhtml, "Script", 2,"")
                }
                if (ss[3] == "true"){
                    strhtml = ScriptHtml(strhtml, "Style", 2,"")
                }
                if (ss[4] == "true"){
                    strhtml = ScriptHtml(strhtml, "Div", 2,"")
                }
                if (ss[5] == "true"){
                    strhtml = ScriptHtml(strhtml, "Span", 2,"")
                }
                if (ss[6] == "true"){
                    strhtml = ScriptHtml(strhtml, "Table", 2,"")
                }
                if (ss[7] == "true"){
                    strhtml = ScriptHtml(strhtml, "Table", 3,"")
                    strhtml = ScriptHtml(strhtml, "Tbody", 3,"")
                    strhtml = ScriptHtml(strhtml, "Tr", 3,"")
                    strhtml = ScriptHtml(strhtml, "Td", 3,"")
                    strhtml = ScriptHtml(strhtml, "Th", 3,"")
                }
                if (ss[8] == "true"){
                    strhtml = ScriptHtml(strhtml, "IMG", 1,"")
                }
                if (ss[9] == "true"){
                    strhtml = ScriptHtml(strhtml, "Font", 3,"")
                }
                if (ss[10] == "true"){
                    strhtml = ScriptHtml(strhtml, "A", 3,"")
                }
                if (ss[11] == "true"){
                    strhtml = ScriptHtml(strhtml, "Font", 4,ss[12])
                }
                HtmlEdit.document.designMode="on";
                HtmlEdit.document.open();
                HtmlEdit.document.write(edithead);
                HtmlEdit.document.write(strhtml);
                HtmlEdit.document.close();    
                doZoom(nCurrZoomSize);
                CurrentMode=0;
                EditMode=true;
                SourceMode=false;
                PreviewMode=false;
                bEditMode=true;
            }
            break;
        default:
            alert("错误参数调用！");
            break;
        }
        range=null;
    }
    // 暂时解决一下通用性问题 查找
    function findstr(){
        var arr = showModalDialog("Editor/editor_find.asp?ChannelID=<%=ChannelID%>", window, "dialogWidth:320px; dialogHeight:170px; help: no; scroll: no; status: no");
    }
// 插入文件函数组 函数组结速
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
</script>
<%
Call CloseConn
%>