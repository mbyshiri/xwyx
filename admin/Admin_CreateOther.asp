<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<!--#include file="../Include/PowerEasy.Common.Rss.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 1      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 1   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = ""   '其他权限

Dim strtmp, MaxPageCol, OutNum, XmlMaxPerPage, XmlOutNum, frequency, Priority, ArtPage, SoftPage, PhotoPage, ProductPage
Dim UOffset, Action2
Dim SubNode, BlogID, SiteLogoUrl, PriClassID, ClassField(5)
Dim strNoSee, strDefAuthor

Action2 = Trim(Request("Action2"))

If Right(SiteUrl, 1) <> "/" Then SiteUrl = SiteUrl & "/"
SiteLogoUrl = SiteUrl & LogoUrl

%>
<html><head><title>生成网站综合数据</title>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<link href='Admin_Style.css' rel='stylesheet' type='text/css'>
</head>
<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>
<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>
  <tr class='topbg'>
    <td height='22' colspan='2' align='center'><strong>生成网站综合数据</strong></td>
  </tr>
  <tr class='tdbg'>
    <td width='70' height='30'><strong>生成说明：</strong></td>
    <td>生成操作比较消耗系统资源及费时，每次生成时，请尽量减少要生成的文件量。</td>
  </tr>
</table>
<br>
<%
If Action2 = "" Then
%>
<table width='100%' border='0' align='center' cellpadding='3' cellspacing='1' class='border'>
    <tr><td class='title'>RSS生成操作</td></tr>
    <tr><td class='tdbg'>
        <table width='530' border='0' align='center' cellpadding='0' cellspacing='0'>
            <form name='formrss' method='post' action='Admin_CreateOther.asp'>
            <tr><td height='40'>
                生成网站首页的ＲＳＳ页面，当您禁用ＲＳＳ或网站首页为动态ＡＳＰ格式时，本功能无效。<br>
                <input name='Action2' type='hidden' id='Action2' value='CreateRss'>
                <input name='submit' type='submit' id='submit' value='开始生成>>'>
            </td></tr>
            </form>
        </table>
    </td></tr>
    <tr><td class='title'>Google地图生成操作</td></tr>
    <tr><td align='center' class='tdbg'>
        <table width='530' border='0' cellspacing='0' cellpadding='0'><a href='http://www.google.com/webmasters/sitemaps/login' target='_blank'><img src="images/GoogleSiteMaplogo.gif" border=0></a>生成符合GOOGLE规范的XML格式地图页面
            <form name='formxmlmap' method='post' action='Admin_CreateOther.asp'>
            <tr><td>
                总输出数量<input name='XmlOutNum' id='XmlOutNum' value='500' size=10 maxlength='5'>&nbsp;<font color=#888888>地图总输出数量</font><br>
                每页连接数<input name='XmlMaxPerPage' id='XmlMaxPerPage' value='100' size=10 maxlength='4'>&nbsp;<font color=#888888>每页连接数,GOOGLE规范要求不得大于５０００</font><br>
                &nbsp;&nbsp;时区偏移<input name='UOffset' id='UOffset' value='08' size=10 maxlength='2'>&nbsp;<font color=#888888>默认中国大陆为８</font><br>
                &nbsp;&nbsp;更新频率<SELECT name=frequency> <OPTION value=always>随时更新</OPTION> <OPTION value=hourly>每 小 时</OPTION> <OPTION value=daily>每天更新</OPTION> <OPTION value=weekly>每周更新</OPTION> <OPTION value=monthly selected>每月更新</OPTION> <OPTION value=yearly>每年更新</OPTION> <OPTION value=never>从不更新</OPTION></SELECT>&nbsp;<font color=#888888>根据站点内容更新情况自行选择</font><br>
                &nbsp;&nbsp;权&nbsp;&nbsp;&nbsp;&nbsp;重<input name='Priority' id='Priority' value='0.5' size=10 maxlength='3'>&nbsp;<font color=#888888>0-1.0之间,推荐使用默认值</font><br>
                <input name='Action2' type='hidden' id='Action2' value='CreateGoogleMap'>
                <input name='submit' type='submit' id='submit' value='开始生成>>'>
            </td></tr>
            </form>
        </table>
    </td></tr>
    <tr><td class='title'>BaiDu地图生成操作</td></tr>
    <tr><td align='center' class='tdbg'>
        <table width='530' border='0' cellspacing='0' cellpadding='0'><a href='http://news.baidu.com/newsop.html' target='_blank'><img src="images/BaiduSiteMaplogo.gif" border=0></a>生成符合百度规范的XML格式地图页面
            <form name='formxmlmap' method='post' action='Admin_CreateOther.asp'>
            <tr><td>
                总输出数量<input name='XmlOutNum' id='XmlOutNum' value='450' size=10 maxlength='5'>&nbsp;<font color=#888888>地图总输出数量</font><br>
                每页连接数<input name='XmlMaxPerPage' id='XmlMaxPerPage' value='90' size=10 maxlength='4'>&nbsp;<font color=#888888>每页连接数,百度规范要求不得大于100</font><br>
                &nbsp;&nbsp;更新频率<input name='frequency' id='frequency' value='1440' size=10 maxlength='6'>&nbsp;<font color=#888888>更新周期，以分钟为单位。搜索引擎将遵照此周期访问该页面，使页面上的新闻更及时地出现在百度新闻中</font><br>
                <input name='Action2' type='hidden' id='Action2' value='CreateBaiDuMap'>
                <input name='submit' type='submit' id='submit' value='开始生成>>'>
            </td></tr>
            </form>
        </table>
    </td></tr>
    <tr><td class='title'>其它类HTML格式地图生成操作</td></tr>
    <tr><td align='center' class='tdbg'>
        <table width='530' border='0' cellspacing='0' cellpadding='0'>
            <form name='formap' method='post' action='Admin_CreateOther.asp'>
            <tr><td>
                生成HTML格式的全站地图页面。<br>
                总输出数量<input name='OutNum' id='OutNum' value='500' size=8 maxlength='5'>&nbsp;<font color=#888888>ＨＴＭＬ地图总输出数量</font><br>
                每页连接数<input name='MaxPerPage' id='MaxPerPage' value='100' size=8 maxlength='3'>&nbsp;<font color=#888888>每页输出数量，不能大于１００</font><br>
                分页换行数<input name='MaxPageCol' id='MaxPageCol' value='27' size=8 maxlength='2'>&nbsp;<font color=#888888>地图分页连接每行显示数</font><br>
                <input name='Action2' type='hidden' id='Action2' value='CreateMap'>
                <input name='submit' type='submit' id='submit' value='开始生成>>'>
            </td></tr>
            </form>
        </table>
    </td></tr>
</table>
<%
Else
    Select Case Action2
    Case "CreateRss"
        If EnableRss = True Then
            Call GetRssIndex_file
            Response.Write "<br><br><a href='Admin_CreateOther.asp'>&lt;&lt; 返回生成管理</a>"
        Else
            Response.Write "<br><br><b>您已经禁用了RSS功能,页面未生成..........<a href='Admin_CreateOther.asp'>&lt;&lt; 返回生成管理</a></b>"
        End If
    Case "CreateMap"
        OutNum = Trim(Request("OutNum"))
        If OutNum = "" Or Not IsNumeric(OutNum) Then
            OutNum = 500
        Else
            OutNum = Int(OutNum)
        End If
        MaxPerPage = Int(Trim(Request("MaxPerPage")))
        If MaxPerPage = "" Or Not IsNumeric(MaxPerPage) Then
            MaxPerPage = 100
        Else
            MaxPerPage = Int(MaxPerPage)
        End If
        MaxPageCol = Int(Trim(Request("MaxPageCol")))
        If MaxPageCol = "" Or Not IsNumeric(MaxPageCol) Then
            MaxPageCol = 27
        Else
            MaxPageCol = Int(MaxPageCol)
        End If
        Call CreateMultiFolder(InstallDir & "SiteMap")

        Response.Write "<br><br><b>正在生成文章类Map页面.........."
        Call OutArticleMap
        Response.Write "</b>"

        Response.Write "<br><br><b>正在生成软件类Map页面.........."
        Call OutSoftMap
        Response.Write "</b>"

        Response.Write "<br><br><b>正在生成图片类Map页面.........."
        Call OutPhotoMap
        Response.Write "</b>"

        Response.Write "<br><br><b>正在生成商品类Map页面.........."
        Call OutProductMap
        Response.Write "</b>"
        Response.Write "<br><br><a href='Admin_CreateOther.asp'>&lt;&lt; 返回生成管理</a>"
    Case "CreateGoogleMap"
        XmlOutNum = Trim(Request("XmlOutNum"))
        If XmlOutNum = "" Or Not IsNumeric(XmlOutNum) Then
            XmlOutNum = 500
        Else
            XmlOutNum = Int(XmlOutNum)
        End If
        XmlMaxPerPage = Trim(Request("XmlMaxPerPage"))
        If XmlMaxPerPage = "" Or Not IsNumeric(XmlMaxPerPage) Then
            XmlMaxPerPage = 27
        Else
            XmlMaxPerPage = Int(XmlMaxPerPage)
        End If
        UOffset = Trim(Request("UOffset"))
        If UOffset = "" Or Not IsNumeric(UOffset) Then
            UOffset = 8
        Else
            UOffset = Int(UOffset)
        End If
        frequency = Trim(Request("frequency"))
        If frequency = "" Then frequency = "Monthly"
        Priority = Trim(Request("Priority"))
        If Priority = "" Then Priority = "0.5"
        
        Response.Write "<br><br><b>正在生成GOOGLE规范XML地图文章页面.........."
        Call OutXmlMap(1)
        Response.Write "</b>"

        Response.Write "<br><br><b>正在生成GOOGLE规范XML地图软件页面.........."
        Call OutXmlMap(2)
        Response.Write "</b>"

        Response.Write "<br><br><b>正在生成GOOGLE规范XML地图图片页面.........."
        Call OutXmlMap(3)
        Response.Write "</b>"
    
        Response.Write "<br><br><b>正在生成GOOGLE规范XML地图商品页面.........."
        Call OutXmlMap(5)
        Response.Write "</b>"

        Response.Write "<br><br><b>正在生成GOOGLE规范XML地图索引页面.........."
        Call OutXmlIndexMap
        Response.Write "</b>"
        Response.Write "<br><br><a href='Admin_CreateOther.asp'>&lt;&lt; 返回生成管理</a>"
    Case "CreateBaiDuMap"
        XmlOutNum = Trim(Request("XmlOutNum"))
        If XmlOutNum = "" Or Not IsNumeric(XmlOutNum) Then
            XmlOutNum = 100
        Else
            XmlOutNum = Int(XmlOutNum)
        End If
        XmlMaxPerPage = Trim(Request("XmlMaxPerPage"))
        If XmlMaxPerPage = "" Or Not IsNumeric(XmlMaxPerPage) Then
            XmlMaxPerPage = 27
        Else
            XmlMaxPerPage = Int(XmlMaxPerPage)
        End If
        frequency = Trim(Request("frequency"))
        If frequency = "" Or Not IsNumeric(frequency) Then
            frequency = 1440
        Else
            frequency = Int(frequency)
        End If
        
        Response.Write "<br><br><b>正在生成百度规范XML地图文章页面.........."
        Call OutBaiDuMap(1)
        Response.Write "</b>"

        Response.Write "<br><br><b>正在生成百度规范XML地图软件页面.........."
        Call OutBaiDuMap(2)
        Response.Write "</b>"

        Response.Write "<br><br><b>正在生成百度规范XML地图图片页面.........."
        Call OutBaiDuMap(3)
        Response.Write "</b>"
    
        Response.Write "<br><br><b>正在生成百度规范XML地图商品页面.........."
        Call OutBaiDuMap(5)
        Response.Write "</b>"

        Response.Write "<br><br><b>百度规范XML地图页面生成完毕,请<a href='http://news.baidu.com/newsop.html' target='_blank'>点击提交到百度</a>..........</b>"
        Response.Write "<br><br><a href='Admin_CreateOther.asp'>&lt;&lt; 返回生成管理</a>"
    Case Else
        Response.Write "<br><br><b>参数错误..........<a href='Admin_CreateOther.asp'>&lt;&lt; 返回生成管理</a></b>"
    End Select
    Set hf = Nothing
End If
%>
</body>
</html>
<!-- Powered by: PowerEasy 2006 -->
<%
Sub GetRssIndex_file()
    Dim strtmp, FileExt_SiteIndex
    XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))

    If EnableRss = True And FileExt_SiteIndex < 4 Then
        Response.Write "<b>正在生成首页RSS导航页面..........</b><br>"
        
        Set XMLDOM = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
        If RssCodeType = True Then
            strtmp = "<?xml version=""1.0"" encoding=""gb2312""?>"
            strtmp = strtmp & "<?xml-stylesheet type=""text/xsl"" href=""../rss.xsl"" version=""1.0""?>"
            Call ShowIndexRss(1)
            strtmp = strtmp & XMLDOM.documentElement.xml
        Else
            strtmp = "<?xml version=""1.0"" encoding=""UTF-8""?>"
            strtmp = strtmp & "<?xml-stylesheet type=""text/xsl"" href=""../rss.xsl"" version=""1.0""?>"
            Call ShowIndexRss(1)
            strtmp = strtmp & unicode(XMLDOM.documentElement.xml)
        End If
        Set Node = Nothing
        Set SubNode = Nothing
        Set XMLDOM = Nothing

        
        If IsNull(strtmp) Then
            Response.Write "<b>生成页面（" & InstallDir & "xml/Rss.xml）<font color=red>失败！</font></b>"
        Else
            If Not fso.FileExists(Server.MapPath(InstallDir & "xml/Rss.xml")) Then
                fso.CreateTextFile (Server.MapPath(InstallDir & "xml/Rss.xml"))
                
            End If
            If fso.FileExists(Server.MapPath(InstallDir & "xml/Rss.xml")) Then
                 Call WriteToFile(InstallDir & "xml/Rss.xml", strtmp)
            Else
                Response.Write "<b>生成页面失败，请检查您的FSO组件是否禁用.</b>"
            End If
            Response.Write "<b>生成页面（<a href='" & InstallDir & "xml/Rss.xml'>" & InstallDir & "xml/Rss.xml</a>）<font color=red>成功！</font></b>"
        End If
    Else
        Response.Write "<b>请检查您是否禁用了RSS功能或您的网站首页是asp模式，无须生成RSS导航页面. </b>"
    End If
End Sub

Sub OutArticleMap()
    Dim rsArticle, sqlArticle, rsChannel, strHTML, totalPut, totalPage, CurrentPage, i, j
    Dim iChannelDir, ChannelType, LinkUrl, preurl, UseCreateHTML, StructureType, FileNameType, FileExt_Item, ClassDir, ParentDir, ClassPurview, iAuthor
    Dim oldChannelID: oldChannelID = 0

    sqlArticle = "select top " & OutNum & " A.ArticleID,A.ChannelID,A.ClassID,A.Title,A.Author,A.UpdateTime,A.Elite,A.Status,A.InfoPoint,A.Deleted,A.LinkUrl,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Article A left join PE_Class C on A.ClassID=C.ClassID Where A.Status=3 and A.Deleted=" & PE_False & " order by A.ArticleID Desc"
    Set rsArticle = Server.CreateObject("adodb.recordset")
    rsArticle.Open sqlArticle, Conn, 1, 1
    If rsArticle.bof And rsArticle.EOF Then
        Response.Write "尚无内容!暂不生成页面!<br>"
    Else
        totalPut = rsArticle.recordcount
        If (totalPut Mod MaxPerPage) = 0 Then
            totalPage = totalPut \ MaxPerPage
        Else
            totalPage = totalPut \ MaxPerPage + 1
        End If
        i = 1
        CurrentPage = 1

        Do While Not rsArticle.EOF

            ClassDir = rsArticle(11)
            ParentDir = rsArticle(12)
            ClassPurview = rsArticle(13)

            If rsArticle(1) <> oldChannelID Then
                Set rsChannel = Conn.Execute("select Top 1 ChannelID,ChannelDir,ChannelType,LinkUrl,UseCreateHTML,StructureType,FileNameType,FileExt_Item from PE_Channel where ChannelID=" & rsArticle(1))
                If Not (rsChannel.bof And rsChannel.EOF) Then
                    iChannelDir = rsChannel("ChannelDir")
                    UseCreateHTML = rsChannel("UseCreateHTML")
                    StructureType = PE_CLng(rsChannel("StructureType"))
                    FileNameType = rsChannel("FileNameType")
                    FileExt_Item = rsChannel("FileExt_Item")
                    ChannelType = rsChannel("ChannelType")
                    LinkUrl = rsChannel("LinkUrl")
                End If
                rsChannel.Close
                If LinkUrl <> "" Then
                    preurl = LinkUrl
                Else
                    preurl = SiteUrl & iChannelDir
                End If
            End If

            iAuthor = rsArticle(4)
            If UseCreateHTML > 0 And (ClassPurview = 0  Or rsArticle(2) = -1) And rsArticle(8) = 0 Then
                strHTML = strHTML & "<li><a href='" & preurl & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle(5)) & GetItemFileName(FileNameType, iChannelDir, rsArticle(5), rsArticle(0)) & arrFileExt(FileExt_Item) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            Else
                strHTML = strHTML & "<li><a href='" & preurl & "/ShowArticle.asp?ArticleID=" & rsArticle(0) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            End If
            i = i + 1

            If i > MaxPerPage Then
                strtmp = "<html>" & vbCrLf
                strtmp = strtmp & "<head>" & vbCrLf
                strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
                strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
                strtmp = strtmp & "<link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
                strtmp = strtmp & "</head>" & vbCrLf
                strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
                strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> 全站文章索引 >> 第" & CurrentPage & "页:<br>" & vbCrLf
                strtmp = strtmp & strHTML & "<br><br>分页:"
                For j = 1 To totalPage
                    If CurrentPage = j Then
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " [" & j & "]<br>"
                        Else
                            strtmp = strtmp & " [" & j & "] "
                        End If
                    Else
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Article" & j & ".htm'>" & j & "</a><br>"
                        Else
                            strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Article" & j & ".htm'>" & j & "</a> "
                        End If
                    End If
                Next
                strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
                strtmp = strtmp & "</html>" & vbCrLf
                Call WriteToFile(InstallDir & "SiteMap/Article" & CurrentPage & ".htm", strtmp)
                Response.Write "<br> 生成页面（<a href='" & InstallDir & "SiteMap/Article" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Article" & CurrentPage & ".htm</a>）<font color=red>成功!</font>"
                CurrentPage = CurrentPage + 1
                i = 1
                strHTML = ""
            End If
            oldChannelID = rsArticle(1)
            rsArticle.movenext
        Loop
        Set rsChannel = Nothing

        strtmp = "<html>" & vbCrLf
        strtmp = strtmp & "<head>" & vbCrLf
        strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
        strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
        strtmp = strtmp & "<link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
        strtmp = strtmp & "</head>" & vbCrLf
        strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
        strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> 全站文章索引 >> 第" & CurrentPage & "页:<br>" & vbCrLf
        strtmp = strtmp & strHTML & "<br><br>分页:"
        For j = 1 To totalPage
            If CurrentPage = j Then
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " [" & j & "]<br>"
                Else
                    strtmp = strtmp & " [" & j & "] "
                End If
            Else
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Article" & j & ".htm'>" & j & "</a><br>"
                Else
                    strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Article" & j & ".htm'>" & j & "</a> "
                End If
            End If
        Next
        strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
        strtmp = strtmp & "</html>" & vbCrLf
        Call WriteToFile(InstallDir & "SiteMap/Article" & CurrentPage & ".htm", strtmp)
        Response.Write "<br> 生成页面（<a href='" & InstallDir & "SiteMap/Article" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Article" & CurrentPage & ".htm</a>）<font color=red>成功!</font>"
        strHTML = strHTML & "<br>" & vbCrLf
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub OutSoftMap()
    Dim rsArticle, sqlArticle, rsChannel, strHTML, totalPut, totalPage, CurrentPage, i, j
    Dim iChannelDir, ChannelType, LinkUrl, preurl, UseCreateHTML, StructureType, FileNameType, FileExt_Item, ClassDir, ParentDir, ClassPurview, iAuthor
    Dim oldChannelID: oldChannelID = 0

    sqlArticle = "select top " & OutNum & " A.SoftID,A.ChannelID,A.ClassID,A.SoftName,A.Author,A.UpdateTime,A.Elite,A.Status,A.Deleted,A.InfoPoint,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Soft A left join PE_Class C on A.ClassID=C.ClassID Where A.Status=3 and A.Deleted=" & PE_False & " order by A.SoftID Desc"
    Set rsArticle = Server.CreateObject("adodb.recordset")
    rsArticle.Open sqlArticle, Conn, 1, 1
    If rsArticle.bof And rsArticle.EOF Then
        Response.Write "尚无内容!暂不生成页面!<br>"
    Else
        totalPut = rsArticle.recordcount
        If (totalPut Mod MaxPerPage) = 0 Then
            totalPage = totalPut \ MaxPerPage
        Else
            totalPage = totalPut \ MaxPerPage + 1
        End If
        i = 1
        CurrentPage = 1

        Do While Not rsArticle.EOF

            ClassDir = rsArticle(10)
            ParentDir = rsArticle(11)
            ClassPurview = rsArticle(12)

            If rsArticle(1) <> oldChannelID Then
                Set rsChannel = Conn.Execute("select Top 1 ChannelID,ChannelDir,ChannelType,LinkUrl,UseCreateHTML,StructureType,FileNameType,FileExt_Item from PE_Channel where ChannelID=" & rsArticle(1))
                If Not (rsChannel.bof And rsChannel.EOF) Then
                    iChannelDir = rsChannel("ChannelDir")
                    UseCreateHTML = rsChannel("UseCreateHTML")
                    StructureType = PE_CLng(rsChannel("StructureType"))
                    FileNameType = rsChannel("FileNameType")
                    FileExt_Item = rsChannel("FileExt_Item")
                    ChannelType = rsChannel("ChannelType")
                    LinkUrl = rsChannel("LinkUrl")
                End If
                rsChannel.Close
                If LinkUrl <> "" Then
                    preurl = LinkUrl
                Else
                    preurl = SiteUrl & iChannelDir
                End If
            End If
        
            iAuthor = rsArticle(4)
            If UseCreateHTML > 0 And (ClassPurview = 0  Or rsArticle(2) = -1) Then
                strHTML = strHTML & "<li><a href='" & preurl & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle(5)) & GetItemFileName(FileNameType, iChannelDir, rsArticle(5), rsArticle(0)) & arrFileExt(FileExt_Item) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            Else
                strHTML = strHTML & "<li><a href='" & preurl & "/ShowSoft.asp?SoftID=" & rsArticle(0) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            End If

            i = i + 1
            If i > MaxPerPage Then
                strtmp = "<html>" & vbCrLf
                strtmp = strtmp & "<head>" & vbCrLf
                strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
                strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
                strtmp = strtmp & "<link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
                strtmp = strtmp & "</head>" & vbCrLf
                strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
                strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> 网站地图 >> 第" & CurrentPage & "页:<br>" & vbCrLf
                strtmp = strtmp & strHTML & "<br><br>分页:"
                For j = 1 To totalPage
                    If CurrentPage = j Then
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " [" & j & "]<br>"
                        Else
                            strtmp = strtmp & " [" & j & "] "
                        End If
                    Else
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Soft" & j & ".htm'>" & j & "</a><br>"
                        Else
                            strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Soft" & j & ".htm'>" & j & "</a> "
                        End If
                    End If
                Next
                strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
                strtmp = strtmp & "</html>" & vbCrLf
                Call WriteToFile(InstallDir & "SiteMap/Soft" & CurrentPage & ".htm", strtmp)
                Response.Write "<br> 生成页面（<a href='" & InstallDir & "SiteMap/Soft" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Soft" & CurrentPage & ".htm</a>）<font color=red>成功!</font>"
                CurrentPage = CurrentPage + 1
                i = 1
                strHTML = ""
            End If
            oldChannelID = rsArticle(1)
            rsArticle.movenext
        Loop
        Set rsChannel = Nothing

        strtmp = "<html>" & vbCrLf
        strtmp = strtmp & "<head>" & vbCrLf
        strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
        strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
        strtmp = strtmp & "<link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
        strtmp = strtmp & "</head>" & vbCrLf
        strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
        strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> 网站地图 >> 第" & CurrentPage & "页:<br>" & vbCrLf
        strtmp = strtmp & strHTML & "<br><br>分页:"
        For j = 1 To totalPage
            If CurrentPage = j Then
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " [" & j & "]<br>"
                Else
                    strtmp = strtmp & " [" & j & "] "
                End If
            Else
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Soft" & j & ".htm'>" & j & "</a><br>"
                Else
                    strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Soft" & j & ".htm'>" & j & "</a> "
                End If
            End If
        Next
        strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
        strtmp = strtmp & "</html>" & vbCrLf
        Call WriteToFile(InstallDir & "SiteMap/Soft" & CurrentPage & ".htm", strtmp)
        Response.Write "<br> 生成页面（<a href='" & InstallDir & "SiteMap/Soft" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Soft" & CurrentPage & ".htm</a>）<font color=red>成功!</font>"
    
        strHTML = strHTML & "<br>" & vbCrLf
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub OutPhotoMap()
    Dim rsArticle, sqlArticle, rsChannel, strHTML, totalPut, totalPage, CurrentPage, i, j
    Dim iChannelDir, ChannelType, LinkUrl, preurl, UseCreateHTML, StructureType, FileNameType, FileExt_Item, ClassDir, ParentDir, ClassPurview, iAuthor
    Dim oldChannelID: oldChannelID = 0

    sqlArticle = "select top " & OutNum & " A.PhotoID,A.ChannelID,A.ClassID,A.PhotoName,A.Author,A.UpdateTime,A.Status,A.Deleted,A.InfoPoint,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Photo A left join PE_Class C on A.ClassID=C.ClassID Where A.Status=3 and A.Deleted=" & PE_False & " order by A.PhotoID Desc"
    Set rsArticle = Server.CreateObject("adodb.recordset")
    rsArticle.Open sqlArticle, Conn, 1, 1
    If rsArticle.bof And rsArticle.EOF Then
        Response.Write "尚无内容!暂不生成页面!<br>"
    Else
        totalPut = rsArticle.recordcount
        If (totalPut Mod MaxPerPage) = 0 Then
            totalPage = totalPut \ MaxPerPage
        Else
            totalPage = totalPut \ MaxPerPage + 1
        End If
        i = 1
        CurrentPage = 1

        Do While Not rsArticle.EOF

            ClassDir = rsArticle(9)
            ParentDir = rsArticle(10)
            ClassPurview = rsArticle(11)

            If rsArticle(1) <> oldChannelID Then
                Set rsChannel = Conn.Execute("select Top 1 ChannelID,ChannelDir,ChannelType,LinkUrl, UseCreateHTML,StructureType,FileNameType,FileExt_Item from PE_Channel where ChannelID=" & rsArticle(1))
                If Not (rsChannel.bof And rsChannel.EOF) Then
                    iChannelDir = rsChannel("ChannelDir")
                    UseCreateHTML = rsChannel("UseCreateHTML")
                    StructureType = PE_CLng(rsChannel("StructureType"))
                    FileNameType = rsChannel("FileNameType")
                    FileExt_Item = rsChannel("FileExt_Item")
                    ChannelType = rsChannel("ChannelType")
                    LinkUrl = rsChannel("LinkUrl")
                End If
                rsChannel.Close
                If LinkUrl <> "" Then
                    preurl = LinkUrl
                Else
                    preurl = SiteUrl & iChannelDir
                End If
            End If
    
            iAuthor = rsArticle(4)
            If UseCreateHTML > 0 And (ClassPurview = 0  Or rsArticle(2) = -1) And rsArticle(8) = 0 Then
                strHTML = strHTML & "<li><a href='" & preurl & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle(5)) & GetItemFileName(FileNameType, iChannelDir, rsArticle(5), rsArticle(0)) & arrFileExt(FileExt_Item) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            Else
                strHTML = strHTML & "<li><a href='" & preurl & "/ShowPhoto.asp?PhotoID=" & rsArticle(0) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            End If
            i = i + 1
            If i > MaxPerPage Then
                strtmp = "<html>" & vbCrLf
                strtmp = strtmp & "<head>" & vbCrLf
                strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
                strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
                strtmp = strtmp & "<link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
                strtmp = strtmp & "</head>" & vbCrLf
                strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
                strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> 网站地图 >> 第" & CurrentPage & "页:<br>" & vbCrLf
                strtmp = strtmp & strHTML & "<br><br>分页:"
                For j = 1 To totalPage
                    If CurrentPage = j Then
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " [" & j & "]<br>"
                        Else
                            strtmp = strtmp & " [" & j & "] "
                        End If
                    Else
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Photo" & j & ".htm'>" & j & "</a><br>"
                        Else
                            strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Photo" & j & ".htm'>" & j & "</a> "
                        End If
                    End If
                Next
                strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
                strtmp = strtmp & "</html>" & vbCrLf
                Call WriteToFile(InstallDir & "SiteMap/Photo" & CurrentPage & ".htm", strtmp)
                Response.Write "<br> 生成页面（<a href='" & InstallDir & "SiteMap/Photo" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Photo" & CurrentPage & ".htm</a>）<font color=red>成功!</font>"
                CurrentPage = CurrentPage + 1
                i = 1
                strHTML = ""
            End If
            oldChannelID = rsArticle(1)
            rsArticle.movenext
        Loop
        Set rsChannel = Nothing

        strtmp = "<html>" & vbCrLf
        strtmp = strtmp & "<head>" & vbCrLf
        strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
        strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
        strtmp = strtmp & "<link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
        strtmp = strtmp & "</head>" & vbCrLf
        strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
        strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> 网站地图 >> 第" & CurrentPage & "页:<br>" & vbCrLf
        strtmp = strtmp & strHTML & "<br><br>分页:"
        For j = 1 To totalPage
            If CurrentPage = j Then
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " [" & j & "]<br>"
                Else
                    strtmp = strtmp & " [" & j & "] "
                End If
            Else
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Photo" & j & ".htm'>" & j & "</a><br>"
                Else
                    strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Photo" & j & ".htm'>" & j & "</a> "
                End If
            End If
        Next
        strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
        strtmp = strtmp & "</html>" & vbCrLf
        Call WriteToFile(InstallDir & "SiteMap/Photo" & CurrentPage & ".htm", strtmp)
        Response.Write "<br> 生成页面（<a href='" & InstallDir & "SiteMap/Photo" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Photo" & CurrentPage & ".htm</a>）<font color=red>成功!</font>"
        strHTML = strHTML & "<br>" & vbCrLf
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub OutProductMap()
    Dim rsArticle, sqlArticle, rsChannel, strHTML, totalPut, totalPage, CurrentPage, i, j
    Dim iChannelDir, ChannelType, LinkUrl, preurl, UseCreateHTML, StructureType, FileNameType, FileExt_Item, ClassDir, ParentDir, ClassPurview, iAuthor
    Dim oldChannelID: oldChannelID = 0

    sqlArticle = "select top " & OutNum & " A.ProductID,A.ChannelID,A.ClassID,A.ProductName,A.ProducerName,A.UpdateTime,A.EnableSale,A.Deleted,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Product A left join PE_Class C on A.ClassID=C.ClassID Where A.Deleted=" & PE_False & " and A.EnableSale=" & PE_True & " order by A.ProductID Desc"
    Set rsArticle = Server.CreateObject("adodb.recordset")
    rsArticle.Open sqlArticle, Conn, 1, 1
    If rsArticle.bof And rsArticle.EOF Then
        Response.Write "尚无内容!暂不生成页面!<br>"
    Else
        totalPut = rsArticle.recordcount
        If (totalPut Mod MaxPerPage) = 0 Then
            totalPage = totalPut \ MaxPerPage
        Else
            totalPage = totalPut \ MaxPerPage + 1
        End If
        i = 1
        CurrentPage = 1

        Do While Not rsArticle.EOF

            ClassDir = rsArticle(8)
            ParentDir = rsArticle(9)
            ClassPurview = rsArticle(10)

            If rsArticle(1) <> oldChannelID Then
                Set rsChannel = Conn.Execute("select Top 1 ChannelID,ChannelDir,ChannelType,LinkUrl,UseCreateHTML,StructureType,FileNameType,FileExt_Item from PE_Channel where ChannelID=" & rsArticle(1))
                If Not (rsChannel.bof And rsChannel.EOF) Then
                    iChannelDir = rsChannel("ChannelDir")
                    UseCreateHTML = rsChannel("UseCreateHTML")
                    StructureType = PE_CLng(rsChannel("StructureType"))
                    FileNameType = rsChannel("FileNameType")
                    FileExt_Item = rsChannel("FileExt_Item")
                    ChannelType = rsChannel("ChannelType")
                    LinkUrl = rsChannel("LinkUrl")
                End If
                rsChannel.Close
                If LinkUrl <> "" Then
                    preurl = LinkUrl
                Else
                    preurl = SiteUrl & iChannelDir
                End If
            End If
        
            iAuthor = rsArticle(4)
            If UseCreateHTML > 0 Then
                strHTML = strHTML & "<li><a href='" & preurl & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle(5)) & GetItemFileName(FileNameType, iChannelDir, rsArticle(5), rsArticle(0)) & arrFileExt(FileExt_Item) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            Else
                strHTML = strHTML & "<li><a href='" & preurl & "/ShowProduct.asp?ProductID=" & rsArticle(0) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            End If
            i = i + 1
            If i > MaxPerPage Then
                strtmp = "<html>" & vbCrLf
                strtmp = strtmp & "<head>" & vbCrLf
                strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
                strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
                strtmp = strtmp & "<link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
                strtmp = strtmp & "</head>" & vbCrLf
                strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
                strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> 网站地图 >> 第" & CurrentPage & "页:<br>" & vbCrLf
                strtmp = strtmp & strHTML & "<br><br>分页:"
                For j = 1 To totalPage
                    If CurrentPage = j Then
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " [" & j & "]<br>"
                        Else
                            strtmp = strtmp & " [" & j & "] "
                        End If
                    Else
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Product" & j & ".htm'>" & j & "</a><br>"
                        Else
                            strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Product" & j & ".htm'>" & j & "</a> "
                        End If
                    End If
                Next
                strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
                strtmp = strtmp & "</html>" & vbCrLf
                Call WriteToFile(InstallDir & "SiteMap/Product" & CurrentPage & ".htm", strtmp)
                Response.Write "<br> 生成页面（<a href='" & InstallDir & "SiteMap/Product" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Product" & CurrentPage & ".htm</a>）<font color=red>成功!</font>"
                CurrentPage = CurrentPage + 1
                i = 1
                strHTML = ""
            End If
            oldChannelID = rsArticle(1)
            rsArticle.movenext
        Loop
        Set rsChannel = Nothing

        strtmp = "<html>" & vbCrLf
        strtmp = strtmp & "<head>" & vbCrLf
        strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
        strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
        strtmp = strtmp & "<link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
        strtmp = strtmp & "</head>" & vbCrLf
        strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
        strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> 网站地图 >> 第" & CurrentPage & "页:<br>" & vbCrLf
        strtmp = strtmp & strHTML & "<br><br>分页:"
        For j = 1 To totalPage
            If CurrentPage = j Then
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " [" & j & "]<br>"
                Else
                    strtmp = strtmp & " [" & j & "] "
                End If
            Else
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Product" & j & ".htm'>" & j & "</a><br>"
                Else
                    strtmp = strtmp & " <a href='" & InstallDir & "SiteMap/Product" & j & ".htm'>" & j & "</a> "
                End If
            End If
        Next
        strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
        strtmp = strtmp & "</html>" & vbCrLf
        Call WriteToFile(InstallDir & "SiteMap/Product" & CurrentPage & ".htm", strtmp)
        Response.Write "<br> 生成页面（<a href='" & InstallDir & "SiteMap/Product" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Product" & CurrentPage & ".htm</a>）<font color=red>成功!</font>"
        strHTML = strHTML & "<br>" & vbCrLf
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub OutXmlMap(OutType)
    Dim rsArticle, sqlArticle, rsChannel, strHTML, totalPut, totalPage, CurrentPage, i, j
    Dim iChannelDir, ChannelType, LinkUrl, preurl, UseCreateHTML, StructureType, FileNameType, FileExt_Item, ClassDir, ParentDir, ClassPurview, AspName, OutFileName
    Dim oldChannelID: oldChannelID = 0
  
    Select Case OutType
    Case 1
        sqlArticle = "select top " & XmlOutNum & " A.ArticleID,A.ChannelID,A.ClassID,A.UpdateTime,A.Status,A.InfoPoint,A.Deleted,A.LinkUrl,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Article A left join PE_Class C on A.ClassID=C.ClassID Where A.Status=3 and A.Deleted=" & PE_False & " order by A.ArticleID Desc"
    Case 2
    sqlArticle = "select top " & XmlOutNum & " A.SoftID,A.ChannelID,A.ClassID,A.UpdateTime,A.Status,A.InfoPoint,A.Deleted,A.Hits,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Soft A left join PE_Class C on A.ClassID=C.ClassID Where A.Status=3 and A.Deleted=" & PE_False & " order by A.SoftID Desc"
    Case 3
    sqlArticle = "select top " & XmlOutNum & " A.PhotoID,A.ChannelID,A.ClassID,A.UpdateTime,A.Status,A.InfoPoint,A.Deleted,A.Hits,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Photo A left join PE_Class C on A.ClassID=C.ClassID Where A.Status=3 and A.Deleted=" & PE_False & " order by A.PhotoID Desc"
    Case 5
    sqlArticle = "select top " & XmlOutNum & " A.ProductID,A.ChannelID,A.ClassID,A.UpdateTime,A.EnableSale,A.Stocks,A.Deleted,A.Hits,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Product A left join PE_Class C on A.ClassID=C.ClassID Where A.Deleted=" & PE_False & " and A.EnableSale=" & PE_True & " order by A.ProductID Desc"
    End Select
    Set rsArticle = Server.CreateObject("adodb.recordset")
    rsArticle.Open sqlArticle, Conn, 1, 1
    If rsArticle.bof And rsArticle.EOF Then
        Response.Write "尚无内容!暂不生成页面!<br>"
    Else
        totalPut = rsArticle.recordcount
        If (totalPut Mod XmlMaxPerPage) = 0 Then
            totalPage = totalPut \ XmlMaxPerPage
        Else
            totalPage = totalPut \ XmlMaxPerPage + 1
        End If
        i = 1
        CurrentPage = 1

        Do While Not rsArticle.EOF

            ClassDir = rsArticle(8)
            ParentDir = rsArticle(9)
            ClassPurview = rsArticle(10)

            If rsArticle(1) <> oldChannelID Then
                Set rsChannel = Conn.Execute("select Top 1 ChannelID,ChannelDir,ChannelType,LinkUrl,UseCreateHTML,StructureType,FileNameType,FileExt_Item from PE_Channel where ChannelID=" & rsArticle(1))
                If Not (rsChannel.bof And rsChannel.EOF) Then
                    iChannelDir = rsChannel("ChannelDir")
                    UseCreateHTML = rsChannel("UseCreateHTML")
                    StructureType = PE_CLng(rsChannel("StructureType"))
                    FileNameType = rsChannel("FileNameType")
                    FileExt_Item = rsChannel("FileExt_Item")
                    ChannelType = rsChannel("ChannelType")
                    LinkUrl = rsChannel("LinkUrl")
                End If
                rsChannel.Close
                If LinkUrl <> "" Then
                    preurl = LinkUrl
                Else
                    preurl = SiteUrl & iChannelDir
                End If
            End If
            Select Case OutType
            Case 1
                AspName = "/ShowArticle.asp?ArticleID="
                OutFileName = "sitemap_article_"
            Case 2
                AspName = "/ShowSoft.asp?SoftID="
                OutFileName = "sitemap_Soft_"
            Case 3
                AspName = "/ShowPhoto.asp?PhotoID="
                OutFileName = "sitemap_Photo_"
            Case 5
                AspName = "/ShowProduct.asp?ProductID="
                OutFileName = "sitemap_Product_"
            End Select
            strHTML = strHTML & "<url>" & vbCrLf
            If OutType < 4 Then
                If UseCreateHTML > 0 And (ClassPurview = 0  Or rsArticle(2) = -1) And rsArticle(5) = 0  Then
                    strHTML = strHTML & "<loc>" & preurl & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle(3)) & GetItemFileName(FileNameType, iChannelDir, rsArticle(3), rsArticle(0)) & arrFileExt(FileExt_Item) & "</loc>" & vbCrLf
                Else
                    strHTML = strHTML & "<loc>" & preurl & AspName & rsArticle(0) & "</loc>" & vbCrLf
                End If
            ElseIf OutType = 5 Then
                If UseCreateHTML > 0 Then
                    strHTML = strHTML & "<loc>" & preurl & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle(3)) & GetItemFileName(FileNameType, iChannelDir, rsArticle(3), rsArticle(0)) & arrFileExt(FileExt_Item) & "</loc>" & vbCrLf
                Else
                    strHTML = strHTML & "<loc>" & preurl & AspName & rsArticle(0) & "</loc>" & vbCrLf
                End If
            End If
            strHTML = strHTML & "<lastmod>" & iso8601date(rsArticle(3), UOffset) & "</lastmod>" & vbCrLf
            strHTML = strHTML & "<changefreq>" & frequency & "</changefreq>" & vbCrLf
            strHTML = strHTML & "<priority>" & Priority & "</priority>" & vbCrLf
            strHTML = strHTML & "</url>" & vbCrLf
            i = i + 1

            If i > XmlMaxPerPage Then
                strtmp = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
                strtmp = strtmp & "<urlset xmlns=""http://www.google.com/schemas/sitemap/0.84"">" & vbCrLf
                strtmp = strtmp & strHTML
                strtmp = strtmp & "</urlset>" & vbCrLf
                Call WriteToFile(InstallDir & OutFileName & CurrentPage & ".xml", strtmp)
                Response.Write "<br> 生成页面（<a href='" & InstallDir & OutFileName & CurrentPage & ".xml' target='_blank'>" & SiteUrl & OutFileName & CurrentPage & ".xml</a>）<font color=red>成功!</font>"
                CurrentPage = CurrentPage + 1
                i = 1
                strHTML = ""
            End If
            oldChannelID = rsArticle(1)
            rsArticle.movenext
        Loop
        Set rsChannel = Nothing

        strtmp = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
        strtmp = strtmp & "<urlset xmlns=""http://www.google.com/schemas/sitemap/0.84"">" & vbCrLf
        strtmp = strtmp & strHTML
        strtmp = strtmp & "</urlset>" & vbCrLf
        Call WriteToFile(InstallDir & OutFileName & CurrentPage & ".xml", strtmp)
        Response.Write "<br> 生成页面（<a href='" & InstallDir & OutFileName & CurrentPage & ".xml' target='_blank'>" & SiteUrl & OutFileName & CurrentPage & ".xml</a>)<font color=red>成功!</font>"
        strHTML = strHTML & "<br>" & vbCrLf
    End If
    Select Case OutType
    Case 1
        ArtPage = totalPage
    Case 2
        SoftPage = totalPage
    Case 3
        PhotoPage = totalPage
    Case 5
        ProductPage = totalPage
    End Select
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub OutXmlIndexMap()
    Dim strtmp, j
    strtmp = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    strtmp = strtmp & "<sitemapindex xmlns=""http://www.google.com/schemas/sitemap/0.84"">" & vbCrLf
    If ArtPage > 0 Then
        For j = 1 To ArtPage
            strtmp = strtmp & "<sitemap>" & vbCrLf
            strtmp = strtmp & "<loc>" & SiteUrl & "sitemap_article_" & j & ".xml</loc>" & vbCrLf
            strtmp = strtmp & "<lastmod>" & iso8601date(Now(), UOffset) & "</lastmod>" & vbCrLf
            strtmp = strtmp & "</sitemap>" & vbCrLf
        Next
    End If
    If SoftPage > 0 Then
        For j = 1 To SoftPage
            strtmp = strtmp & "<sitemap>" & vbCrLf
            strtmp = strtmp & "<loc>" & SiteUrl & "sitemap_Soft_" & j & ".xml</loc>" & vbCrLf
            strtmp = strtmp & "<lastmod>" & iso8601date(Now(), UOffset) & "</lastmod>" & vbCrLf
            strtmp = strtmp & "</sitemap>" & vbCrLf
        Next
    End If
    If PhotoPage > 0 Then
        For j = 1 To PhotoPage
            strtmp = strtmp & "<sitemap>" & vbCrLf
            strtmp = strtmp & "<loc>" & SiteUrl & "sitemap_Photo_" & j & ".xml</loc>" & vbCrLf
            strtmp = strtmp & "<lastmod>" & iso8601date(Now(), UOffset) & "</lastmod>" & vbCrLf
            strtmp = strtmp & "</sitemap>" & vbCrLf
        Next
    End If
    If ProductPage > 0 Then
        For j = 1 To ProductPage
            strtmp = strtmp & "<sitemap>" & vbCrLf
            strtmp = strtmp & "<loc>" & SiteUrl & "sitemap_Product_" & j & ".xml</loc>" & vbCrLf
            strtmp = strtmp & "<lastmod>" & iso8601date(Now(), UOffset) & "</lastmod>" & vbCrLf
            strtmp = strtmp & "</sitemap>" & vbCrLf
        Next
    End If
    strtmp = strtmp & "</sitemapindex>" & vbCrLf
    Call WriteToFile(InstallDir & "sitemap_index.xml", strtmp)
    Response.Write "<br> 生成页面（<a href='" & InstallDir & "sitemap_index.xml' target='_blank'>" & InstallDir & "sitemap_index.xml</a>）<font color=red>成功!</font>，&nbsp;[<a href='http://www.google.com/webmasters/sitemaps/ping?sitemap=" & SiteUrl & "sitemap_index.xml' target='_blank'>点击这里提交到Google</a>]"
End Sub

Sub OutBaiDuMap(OutType)
    Dim rsArticle, sqlArticle, rsChannel, strHTML, totalPut, totalPage, CurrentPage, i, j
    Dim iChannelDir, ChannelType, LinkUrl, preurl, UseCreateHTML, StructureType, FileNameType, FileExt_Item, CUploadDir, ClassDir, ParentDir, ClassPurview, AspName, OutFileName
    Dim oldChannelID: oldChannelID = 0
  
    Select Case OutType
    Case 1
        sqlArticle = "select top " & XmlOutNum & " A.ArticleID,A.ChannelID,A.ClassID,A.UpdateTime,A.Status,A.InfoPoint,A.Deleted,A.LinkUrl,C.ClassDir,C.ParentDir,C.ClassPurview,A.Title,A.Author,A.CopyFrom,A.Keyword,A.Content,A.DefaultPicUrl from PE_Article A left join PE_Class C on A.ClassID=C.ClassID Where A.Status=3 and A.Deleted=" & PE_False & " order by A.ArticleID Desc"
    Case 2
        sqlArticle = "select top " & XmlOutNum & " A.SoftID,A.ChannelID,A.ClassID,A.UpdateTime,A.Status,A.InfoPoint,A.Deleted,A.Hits,C.ClassDir,C.ParentDir,C.ClassPurview,A.SoftName,A.Author,A.CopyFrom,A.Keyword,A.SoftIntro,A.SoftPicUrl from PE_Soft A left join PE_Class C on A.ClassID=C.ClassID Where A.Status=3 and A.Deleted=" & PE_False & " order by A.SoftID Desc"
    Case 3
        sqlArticle = "select top " & XmlOutNum & " A.PhotoID,A.ChannelID,A.ClassID,A.UpdateTime,A.Status,A.InfoPoint,A.Deleted,A.Hits,C.ClassDir,C.ParentDir,C.ClassPurview,A.PhotoName,A.Author,A.CopyFrom,A.Keyword,A.PhotoIntro,A.PhotoThumb from PE_Photo A left join PE_Class C on A.ClassID=C.ClassID Where A.Status=3 and A.Deleted=" & PE_False & " order by A.PhotoID Desc"
    Case 5
        sqlArticle = "select top " & XmlOutNum & " A.ProductID,A.ChannelID,A.ClassID,A.UpdateTime,A.EnableSale,A.Stocks,A.Deleted,A.Hits,C.ClassDir,C.ParentDir,C.ClassPurview,A.ProductName,A.ProducerName,A.TrademarkName,A.Keyword,A.ProductIntro,A.ProductThumb from PE_Product A left join PE_Class C on A.ClassID=C.ClassID Where A.Deleted=" & PE_False & " and A.EnableSale=" & PE_True & " order by A.ProductID Desc"
    End Select
    Set rsArticle = Server.CreateObject("adodb.recordset")
    rsArticle.Open sqlArticle, Conn, 1, 1
    If rsArticle.bof And rsArticle.EOF Then
        Response.Write "尚无内容!暂不生成页面!<br>"
    Else
        totalPut = rsArticle.recordcount
        If (totalPut Mod XmlMaxPerPage) = 0 Then
            totalPage = totalPut \ XmlMaxPerPage
        Else
            totalPage = totalPut \ XmlMaxPerPage + 1
        End If
        i = 1
        j = 1
        CurrentPage = 1

        Do While Not rsArticle.EOF
            ClassDir = rsArticle(8)
            ParentDir = rsArticle(9)
            ClassPurview = rsArticle(10)
            If rsArticle(1) <> oldChannelID Then
                Set rsChannel = Conn.Execute("select Top 1 ChannelID,ChannelDir,ChannelType,LinkUrl,UseCreateHTML,StructureType,FileNameType,FileExt_Item,UploadDir from PE_Channel where ChannelID=" & rsArticle(1))
                If Not (rsChannel.bof And rsChannel.EOF) Then
                    iChannelDir = rsChannel("ChannelDir")
                    UseCreateHTML = rsChannel("UseCreateHTML")
                    StructureType = PE_CLng(rsChannel("StructureType"))
                    FileNameType = rsChannel("FileNameType")
                    FileExt_Item = rsChannel("FileExt_Item")
                    ChannelType = rsChannel("ChannelType")
                    LinkUrl = rsChannel("LinkUrl")
                    CUploadDir = rsChannel("UploadDir")
                End If
                rsChannel.Close
                If LinkUrl <> "" Then
                    preurl = LinkUrl
                Else
                    preurl = SiteUrl & iChannelDir
                End If
            End If
            Select Case OutType
            Case 1
                AspName = "/ShowArticle.asp?ArticleID="
                OutFileName = "baidumap_article_"
            Case 2
                AspName = "/ShowSoft.asp?SoftID="
                OutFileName = "baidumap_Soft_"
            Case 3
                AspName = "/ShowPhoto.asp?PhotoID="
                OutFileName = "baidumap_Photo_"
            Case 5
                AspName = "/ShowProduct.asp?ProductID="
                OutFileName = "baidumap_Product_"
            End Select
            strHTML = strHTML & "<item>" & vbCrLf
            strHTML = strHTML & "<title>" & fhtml(rsArticle(11)) & "</title>" & vbCrLf

            If OutType < 4 Then
                If UseCreateHTML > 0 And (ClassPurview = 0 Or rsArticle(2) = -1) And rsArticle(5) = 0 Then
                    strHTML = strHTML & "<link>" & preurl & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle(3)) & GetItemFileName(FileNameType, iChannelDir, rsArticle(3), rsArticle(0)) & arrFileExt(FileExt_Item) & "</link>" & vbCrLf
                Else
                    strHTML = strHTML & "<link>" & preurl & AspName & rsArticle(0) & "</link>" & vbCrLf
                End If
            ElseIf OutType = 5 Then
                If UseCreateHTML > 0 Then
                    strHTML = strHTML & "<link>" & preurl & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle(3)) & GetItemFileName(FileNameType, iChannelDir, rsArticle(3), rsArticle(0)) & arrFileExt(FileExt_Item) & "</link>" & vbCrLf
                Else
                    strHTML = strHTML & "<link>" & preurl & AspName & rsArticle(0) & "</link>" & vbCrLf
                End If
            End If
            strHTML = strHTML & "<text>" & GetSubStr(fhtml(rsArticle(15)), 600, "") & "</text>" & vbCrLf
            If rsArticle(16) <> "" Then
                strHTML = strHTML & "<image>" & preurl & "/" & CUploadDir & "/" & rsArticle(16) & "</image>" & vbCrLf
            Else
                strHTML = strHTML & "<image/>" & vbCrLf
            End If
            strHTML = strHTML & "<keywords>" & fhtml(Replace(rsArticle(14), "|", " ")) & "</keywords>" & vbCrLf
            strHTML = strHTML & "<author>" & fhtml(rsArticle(12)) & "</author>" & vbCrLf
            strHTML = strHTML & "<source>" & fhtml(rsArticle(13)) & "</source>" & vbCrLf
            strHTML = strHTML & "<pubDate>" & rsArticle(3) & "</pubDate>" & vbCrLf
            strHTML = strHTML & "</item>" & vbCrLf
            i = i + 1
            j = j + 1

            If i > XmlMaxPerPage Then
                strtmp = "<?xml version=""1.0"" encoding=""GB2312""?>" & vbCrLf
                strtmp = strtmp & "<document>" & vbCrLf
                strtmp = strtmp & "<webSite>" & SiteUrl & "</webSite>" & vbCrLf
                strtmp = strtmp & "<webMaster>" & WebmasterEmail & "</webMaster>" & vbCrLf
                strtmp = strtmp & "<updatePeri>" & frequency & "</updatePeri>" & vbCrLf
                strtmp = strtmp & strHTML
                strtmp = strtmp & "</document>" & vbCrLf
                Call WriteToFile(InstallDir & OutFileName & CurrentPage & ".xml", strtmp)
                Response.Write "<br> 生成页面（<a href='" & InstallDir & OutFileName & CurrentPage & ".xml' target='_blank'>" & SiteUrl & OutFileName & CurrentPage & ".xml</a>）<font color=red>成功!</font>"
                CurrentPage = CurrentPage + 1
                i = 1
                strHTML = ""
            End If
            oldChannelID = rsArticle(1)
            If j > XmlOutNum Then
                Exit Do
            End If
            rsArticle.movenext
        Loop
        Set rsChannel = Nothing

        strtmp = "<?xml version=""1.0"" encoding=""GB2312""?>" & vbCrLf
        strtmp = strtmp & "<document>" & vbCrLf
        strtmp = strtmp & "<webSite>" & SiteUrl & "</webSite>" & vbCrLf
        strtmp = strtmp & "<webMaster>" & WebmasterEmail & "</webMaster>" & vbCrLf
        strtmp = strtmp & "<updatePeri>" & frequency & "</updatePeri>" & vbCrLf
        strtmp = strtmp & strHTML
        strtmp = strtmp & "</document>" & vbCrLf
        Call WriteToFile(InstallDir & OutFileName & CurrentPage & ".xml", strtmp)
        Response.Write "<br> 生成页面（<a href='" & InstallDir & OutFileName & CurrentPage & ".xml' target='_blank'>" & SiteUrl & OutFileName & CurrentPage & ".xml</a>)<font color=red>成功!</font>"
        strHTML = strHTML & "<br>" & vbCrLf
    End If
    Select Case OutType
    Case 1
        ArtPage = totalPage
    Case 2
        SoftPage = totalPage
    Case 3
        PhotoPage = totalPage
    Case 5
        ProductPage = totalPage
    End Select
    rsArticle.Close
    Set rsArticle = Nothing
End Sub


'**************************************************
'函数名：iso8601date
'作  用：时间格式转换
'**************************************************
Function iso8601date(dLocal, utcOffset)
    Dim d, d1
    d = DateAdd("H", -1 * utcOffset, dLocal)
    If Len(utcOffset) < 2 Then
        d1 = "0" & utcOffset
    Else
        d1 = utcOffset
    End If
    iso8601date = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2) & "T"
    iso8601date = iso8601date & (Right("0" & Hour(d), 2) & ":" & Right("0" & Minute(d), 2) & ":" & Right("0" & Second(d), 2))
    If utcOffset < 0 Then
        iso8601date = iso8601date & ("-" & d1 & ":00")
    Else
        iso8601date = iso8601date & ("+" & d1 & ":00")
    End If
End Function

'**************************************************
'函数名：fhtml
'作  用：强化过滤HTML标记
'**************************************************
Function fhtml(istr)
    istr = cuthtml(istr)
    fhtml = Replace(Replace(Replace(Replace(Replace(Replace(istr, " ", ""), "'", "&apos;"), """", "&quot;"), ">", "&gt;"), "<", "&lt;"), "&", "&amp;")
End Function

Function cuthtml(ByVal str)
    If IsNull(str) Or Trim(str) = "" Then
        cuthtml = ""
        Exit Function
    End If
    regEx.Pattern = "(\<.[^\<]*\>)"
    str = regEx.Replace(str, " ")
    regEx.Pattern = "(\<\/[^\<]*\>)"
    str = regEx.Replace(str, "")
    
    str = Replace(str, "'", "")
    str = Replace(str, Chr(34), "")
    cuthtml = str
End Function
%>
