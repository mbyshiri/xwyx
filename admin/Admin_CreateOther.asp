<!--#include file="Admin_Common.asp"-->
<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<!--#include file="../Include/PowerEasy.Common.Rss.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Const NeedCheckComeUrl = True   '�Ƿ���Ҫ����ⲿ����

Const PurviewLevel = 1      '0--����飬1--��������Ա��2--��ͨ����Ա
Const PurviewLevel_Channel = 1   '0--����飬1--Ƶ������Ա��2--��Ŀ�ܱ࣬3--��Ŀ����Ա
Const PurviewLevel_Others = ""   '����Ȩ��

Dim strtmp, MaxPageCol, OutNum, XmlMaxPerPage, XmlOutNum, frequency, Priority, ArtPage, SoftPage, PhotoPage, ProductPage
Dim UOffset, Action2
Dim SubNode, BlogID, SiteLogoUrl, PriClassID, ClassField(5)
Dim strNoSee, strDefAuthor

Action2 = Trim(Request("Action2"))

If Right(SiteUrl, 1) <> "/" Then SiteUrl = SiteUrl & "/"
SiteLogoUrl = SiteUrl & LogoUrl

%>
<html><head><title>������վ�ۺ�����</title>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<link href='Admin_Style.css' rel='stylesheet' type='text/css'>
</head>
<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>
<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>
  <tr class='topbg'>
    <td height='22' colspan='2' align='center'><strong>������վ�ۺ�����</strong></td>
  </tr>
  <tr class='tdbg'>
    <td width='70' height='30'><strong>����˵����</strong></td>
    <td>���ɲ����Ƚ�����ϵͳ��Դ����ʱ��ÿ������ʱ���뾡������Ҫ���ɵ��ļ�����</td>
  </tr>
</table>
<br>
<%
If Action2 = "" Then
%>
<table width='100%' border='0' align='center' cellpadding='3' cellspacing='1' class='border'>
    <tr><td class='title'>RSS���ɲ���</td></tr>
    <tr><td class='tdbg'>
        <table width='530' border='0' align='center' cellpadding='0' cellspacing='0'>
            <form name='formrss' method='post' action='Admin_CreateOther.asp'>
            <tr><td height='40'>
                ������վ��ҳ�ģңӣ�ҳ�棬�������ãңӣӻ���վ��ҳΪ��̬���ӣи�ʽʱ����������Ч��<br>
                <input name='Action2' type='hidden' id='Action2' value='CreateRss'>
                <input name='submit' type='submit' id='submit' value='��ʼ����>>'>
            </td></tr>
            </form>
        </table>
    </td></tr>
    <tr><td class='title'>Google��ͼ���ɲ���</td></tr>
    <tr><td align='center' class='tdbg'>
        <table width='530' border='0' cellspacing='0' cellpadding='0'><a href='http://www.google.com/webmasters/sitemaps/login' target='_blank'><img src="images/GoogleSiteMaplogo.gif" border=0></a>���ɷ���GOOGLE�淶��XML��ʽ��ͼҳ��
            <form name='formxmlmap' method='post' action='Admin_CreateOther.asp'>
            <tr><td>
                ���������<input name='XmlOutNum' id='XmlOutNum' value='500' size=10 maxlength='5'>&nbsp;<font color=#888888>��ͼ���������</font><br>
                ÿҳ������<input name='XmlMaxPerPage' id='XmlMaxPerPage' value='100' size=10 maxlength='4'>&nbsp;<font color=#888888>ÿҳ������,GOOGLE�淶Ҫ�󲻵ô��ڣ�������</font><br>
                &nbsp;&nbsp;ʱ��ƫ��<input name='UOffset' id='UOffset' value='08' size=10 maxlength='2'>&nbsp;<font color=#888888>Ĭ���й���½Ϊ��</font><br>
                &nbsp;&nbsp;����Ƶ��<SELECT name=frequency> <OPTION value=always>��ʱ����</OPTION> <OPTION value=hourly>ÿ С ʱ</OPTION> <OPTION value=daily>ÿ�����</OPTION> <OPTION value=weekly>ÿ�ܸ���</OPTION> <OPTION value=monthly selected>ÿ�¸���</OPTION> <OPTION value=yearly>ÿ�����</OPTION> <OPTION value=never>�Ӳ�����</OPTION></SELECT>&nbsp;<font color=#888888>����վ�����ݸ����������ѡ��</font><br>
                &nbsp;&nbsp;Ȩ&nbsp;&nbsp;&nbsp;&nbsp;��<input name='Priority' id='Priority' value='0.5' size=10 maxlength='3'>&nbsp;<font color=#888888>0-1.0֮��,�Ƽ�ʹ��Ĭ��ֵ</font><br>
                <input name='Action2' type='hidden' id='Action2' value='CreateGoogleMap'>
                <input name='submit' type='submit' id='submit' value='��ʼ����>>'>
            </td></tr>
            </form>
        </table>
    </td></tr>
    <tr><td class='title'>BaiDu��ͼ���ɲ���</td></tr>
    <tr><td align='center' class='tdbg'>
        <table width='530' border='0' cellspacing='0' cellpadding='0'><a href='http://news.baidu.com/newsop.html' target='_blank'><img src="images/BaiduSiteMaplogo.gif" border=0></a>���ɷ��ϰٶȹ淶��XML��ʽ��ͼҳ��
            <form name='formxmlmap' method='post' action='Admin_CreateOther.asp'>
            <tr><td>
                ���������<input name='XmlOutNum' id='XmlOutNum' value='450' size=10 maxlength='5'>&nbsp;<font color=#888888>��ͼ���������</font><br>
                ÿҳ������<input name='XmlMaxPerPage' id='XmlMaxPerPage' value='90' size=10 maxlength='4'>&nbsp;<font color=#888888>ÿҳ������,�ٶȹ淶Ҫ�󲻵ô���100</font><br>
                &nbsp;&nbsp;����Ƶ��<input name='frequency' id='frequency' value='1440' size=10 maxlength='6'>&nbsp;<font color=#888888>�������ڣ��Է���Ϊ��λ���������潫���մ����ڷ��ʸ�ҳ�棬ʹҳ���ϵ����Ÿ���ʱ�س����ڰٶ�������</font><br>
                <input name='Action2' type='hidden' id='Action2' value='CreateBaiDuMap'>
                <input name='submit' type='submit' id='submit' value='��ʼ����>>'>
            </td></tr>
            </form>
        </table>
    </td></tr>
    <tr><td class='title'>������HTML��ʽ��ͼ���ɲ���</td></tr>
    <tr><td align='center' class='tdbg'>
        <table width='530' border='0' cellspacing='0' cellpadding='0'>
            <form name='formap' method='post' action='Admin_CreateOther.asp'>
            <tr><td>
                ����HTML��ʽ��ȫվ��ͼҳ�档<br>
                ���������<input name='OutNum' id='OutNum' value='500' size=8 maxlength='5'>&nbsp;<font color=#888888>�ȣԣ̵ͣ�ͼ���������</font><br>
                ÿҳ������<input name='MaxPerPage' id='MaxPerPage' value='100' size=8 maxlength='3'>&nbsp;<font color=#888888>ÿҳ������������ܴ��ڣ�����</font><br>
                ��ҳ������<input name='MaxPageCol' id='MaxPageCol' value='27' size=8 maxlength='2'>&nbsp;<font color=#888888>��ͼ��ҳ����ÿ����ʾ��</font><br>
                <input name='Action2' type='hidden' id='Action2' value='CreateMap'>
                <input name='submit' type='submit' id='submit' value='��ʼ����>>'>
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
            Response.Write "<br><br><a href='Admin_CreateOther.asp'>&lt;&lt; �������ɹ���</a>"
        Else
            Response.Write "<br><br><b>���Ѿ�������RSS����,ҳ��δ����..........<a href='Admin_CreateOther.asp'>&lt;&lt; �������ɹ���</a></b>"
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

        Response.Write "<br><br><b>��������������Mapҳ��.........."
        Call OutArticleMap
        Response.Write "</b>"

        Response.Write "<br><br><b>�������������Mapҳ��.........."
        Call OutSoftMap
        Response.Write "</b>"

        Response.Write "<br><br><b>��������ͼƬ��Mapҳ��.........."
        Call OutPhotoMap
        Response.Write "</b>"

        Response.Write "<br><br><b>����������Ʒ��Mapҳ��.........."
        Call OutProductMap
        Response.Write "</b>"
        Response.Write "<br><br><a href='Admin_CreateOther.asp'>&lt;&lt; �������ɹ���</a>"
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
        
        Response.Write "<br><br><b>��������GOOGLE�淶XML��ͼ����ҳ��.........."
        Call OutXmlMap(1)
        Response.Write "</b>"

        Response.Write "<br><br><b>��������GOOGLE�淶XML��ͼ���ҳ��.........."
        Call OutXmlMap(2)
        Response.Write "</b>"

        Response.Write "<br><br><b>��������GOOGLE�淶XML��ͼͼƬҳ��.........."
        Call OutXmlMap(3)
        Response.Write "</b>"
    
        Response.Write "<br><br><b>��������GOOGLE�淶XML��ͼ��Ʒҳ��.........."
        Call OutXmlMap(5)
        Response.Write "</b>"

        Response.Write "<br><br><b>��������GOOGLE�淶XML��ͼ����ҳ��.........."
        Call OutXmlIndexMap
        Response.Write "</b>"
        Response.Write "<br><br><a href='Admin_CreateOther.asp'>&lt;&lt; �������ɹ���</a>"
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
        
        Response.Write "<br><br><b>�������ɰٶȹ淶XML��ͼ����ҳ��.........."
        Call OutBaiDuMap(1)
        Response.Write "</b>"

        Response.Write "<br><br><b>�������ɰٶȹ淶XML��ͼ���ҳ��.........."
        Call OutBaiDuMap(2)
        Response.Write "</b>"

        Response.Write "<br><br><b>�������ɰٶȹ淶XML��ͼͼƬҳ��.........."
        Call OutBaiDuMap(3)
        Response.Write "</b>"
    
        Response.Write "<br><br><b>�������ɰٶȹ淶XML��ͼ��Ʒҳ��.........."
        Call OutBaiDuMap(5)
        Response.Write "</b>"

        Response.Write "<br><br><b>�ٶȹ淶XML��ͼҳ���������,��<a href='http://news.baidu.com/newsop.html' target='_blank'>����ύ���ٶ�</a>..........</b>"
        Response.Write "<br><br><a href='Admin_CreateOther.asp'>&lt;&lt; �������ɹ���</a>"
    Case Else
        Response.Write "<br><br><b>��������..........<a href='Admin_CreateOther.asp'>&lt;&lt; �������ɹ���</a></b>"
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
        Response.Write "<b>����������ҳRSS����ҳ��..........</b><br>"
        
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
            Response.Write "<b>����ҳ�棨" & InstallDir & "xml/Rss.xml��<font color=red>ʧ�ܣ�</font></b>"
        Else
            If Not fso.FileExists(Server.MapPath(InstallDir & "xml/Rss.xml")) Then
                fso.CreateTextFile (Server.MapPath(InstallDir & "xml/Rss.xml"))
                
            End If
            If fso.FileExists(Server.MapPath(InstallDir & "xml/Rss.xml")) Then
                 Call WriteToFile(InstallDir & "xml/Rss.xml", strtmp)
            Else
                Response.Write "<b>����ҳ��ʧ�ܣ���������FSO����Ƿ����.</b>"
            End If
            Response.Write "<b>����ҳ�棨<a href='" & InstallDir & "xml/Rss.xml'>" & InstallDir & "xml/Rss.xml</a>��<font color=red>�ɹ���</font></b>"
        End If
    Else
        Response.Write "<b>�������Ƿ������RSS���ܻ�������վ��ҳ��aspģʽ����������RSS����ҳ��. </b>"
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
        Response.Write "��������!�ݲ�����ҳ��!<br>"
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
                strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ȫվ�������� >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
                strtmp = strtmp & strHTML & "<br><br>��ҳ:"
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
                Response.Write "<br> ����ҳ�棨<a href='" & InstallDir & "SiteMap/Article" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Article" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
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
        strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ȫվ�������� >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
        strtmp = strtmp & strHTML & "<br><br>��ҳ:"
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
        Response.Write "<br> ����ҳ�棨<a href='" & InstallDir & "SiteMap/Article" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Article" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
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
        Response.Write "��������!�ݲ�����ҳ��!<br>"
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
                strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ��վ��ͼ >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
                strtmp = strtmp & strHTML & "<br><br>��ҳ:"
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
                Response.Write "<br> ����ҳ�棨<a href='" & InstallDir & "SiteMap/Soft" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Soft" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
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
        strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ��վ��ͼ >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
        strtmp = strtmp & strHTML & "<br><br>��ҳ:"
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
        Response.Write "<br> ����ҳ�棨<a href='" & InstallDir & "SiteMap/Soft" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Soft" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
    
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
        Response.Write "��������!�ݲ�����ҳ��!<br>"
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
                strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ��վ��ͼ >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
                strtmp = strtmp & strHTML & "<br><br>��ҳ:"
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
                Response.Write "<br> ����ҳ�棨<a href='" & InstallDir & "SiteMap/Photo" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Photo" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
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
        strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ��վ��ͼ >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
        strtmp = strtmp & strHTML & "<br><br>��ҳ:"
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
        Response.Write "<br> ����ҳ�棨<a href='" & InstallDir & "SiteMap/Photo" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Photo" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
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
        Response.Write "��������!�ݲ�����ҳ��!<br>"
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
                strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ��վ��ͼ >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
                strtmp = strtmp & strHTML & "<br><br>��ҳ:"
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
                Response.Write "<br> ����ҳ�棨<a href='" & InstallDir & "SiteMap/Product" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Product" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
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
        strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ��վ��ͼ >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
        strtmp = strtmp & strHTML & "<br><br>��ҳ:"
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
        Response.Write "<br> ����ҳ�棨<a href='" & InstallDir & "SiteMap/Product" & CurrentPage & ".htm' target='_blank'>" & SiteUrl & "SiteMap/Product" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
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
        Response.Write "��������!�ݲ�����ҳ��!<br>"
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
                Response.Write "<br> ����ҳ�棨<a href='" & InstallDir & OutFileName & CurrentPage & ".xml' target='_blank'>" & SiteUrl & OutFileName & CurrentPage & ".xml</a>��<font color=red>�ɹ�!</font>"
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
        Response.Write "<br> ����ҳ�棨<a href='" & InstallDir & OutFileName & CurrentPage & ".xml' target='_blank'>" & SiteUrl & OutFileName & CurrentPage & ".xml</a>)<font color=red>�ɹ�!</font>"
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
    Response.Write "<br> ����ҳ�棨<a href='" & InstallDir & "sitemap_index.xml' target='_blank'>" & InstallDir & "sitemap_index.xml</a>��<font color=red>�ɹ�!</font>��&nbsp;[<a href='http://www.google.com/webmasters/sitemaps/ping?sitemap=" & SiteUrl & "sitemap_index.xml' target='_blank'>��������ύ��Google</a>]"
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
        Response.Write "��������!�ݲ�����ҳ��!<br>"
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
                Response.Write "<br> ����ҳ�棨<a href='" & InstallDir & OutFileName & CurrentPage & ".xml' target='_blank'>" & SiteUrl & OutFileName & CurrentPage & ".xml</a>��<font color=red>�ɹ�!</font>"
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
        Response.Write "<br> ����ҳ�棨<a href='" & InstallDir & OutFileName & CurrentPage & ".xml' target='_blank'>" & SiteUrl & OutFileName & CurrentPage & ".xml</a>)<font color=red>�ɹ�!</font>"
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
'��������iso8601date
'��  �ã�ʱ���ʽת��
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
'��������fhtml
'��  �ã�ǿ������HTML���
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
