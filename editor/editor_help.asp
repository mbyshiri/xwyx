<%@language=vbscript codepage=936 %>
<%
option explicit
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>动易HTML在线编辑器使用帮助</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
TD {
	font-size: 9pt;
}
-->
</style>
</head>

<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellpadding="6" cellspacing="0">
  <tr> 
    <td align="center" bgcolor="#0066FF"><b><font color="#ffffff">动易HTML在线编辑器使用帮助</font></b></td>
  </tr>
  <tr> 
    <td>1、本HTML在线编辑器具有强大的图文混排功能！文字可自由设定颜色、段落格式、缩进量、文字上下标等格式，并可方便地删除文字格式。<br>
  2、可自动识别各种网址和Email地址，可增加或取消链接、插入特殊水平线等功能，支持从word中粘贴内容并去除冗余代码。<br>
        3、本编辑器支持以下贴图功能，可以在文档的任意位置插入图片及表格：
        <li>可直接复制网上的图片，然后粘贴到此编辑器中即可，编辑器会自动获得图片的URL地址。（强烈推荐使用此功能插入图片）</li>
        <li>使用“插入图片”功能按钮。可以在插入图片时指定图片的URL地址、图片大小、图片css效果、对齐方式、边框粗细等。（推荐使用）</li>
        <li>使用“批量上传图片”功能按钮，最多可以同时上传10张图片，并可预览图片效果。</li>
        <li>使用无组件上传功能上传文件功能，可上传本地图片、附件等，上传后会直接显示出来。具有从已上传的文件中选择的功能。</li>
        <li>在编辑HTML源代码状态下，手工输入图片代码。</li><br>
        4、插入表格时可指定各项参数，即时显示出效果并进行修改。<br>
        5、可插入flash多媒体、视频文件和Realplay文件，并使其能在网页中即时播放，支持插入栏目框和插入网页的功能！<br>
        6、具有便捷的右键编辑功能，可进行常规的复制、粘贴，也可对表格进行合并、插入、修改等操作。<br>
        7、具有“编辑”、“源代码”、“预览”、和“文本”四种视图模式，可扩大或缩小编辑区。<br>
        8、如果要手动书写源代码，请选中“源代码”视图模式。支持所有的HTML标签。。 
        书写完毕后，请选中“编辑”视图模式，HTML代码可立即显示实际效果。
        <br>
        <br>
    <font color="#FF0000">感谢：</font>本编辑器的部分功能如表格处理等借鉴了<a href="http://www.webasp.net" target="_blank">eWebEditor</a>的思路，在此特表示感谢。</td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#f5f5f5">原创：佚名&nbsp;&nbsp;&nbsp;&nbsp;修改：WEBBOY、雅虎、动力兔&nbsp;&nbsp;&nbsp;&nbsp;</td>
  </tr>
</table>
</body>
</html>
