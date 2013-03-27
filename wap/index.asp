<!--#include file="../Start.asp"-->
<!--#include file="../Include/PowerEasy.Cache.asp"-->
<!--#include file="../Include/PowerEasy.Channel.asp"-->
<!--#include file="../Include/PowerEasy.MD5.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim ReadMe, WapLocationUrl
ReadMe = Trim(Request("ReadMe"))
XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))
WapLocationUrl = SiteUrl & "/wap/index.asp"
WapDomain = XmlText("Wap", "Domain", WapLocationUrl)
If WapDomain <> WapLocationUrl And Right(WapDomain, 1) <> "/" Then
    WapDomain = WapDomain & "/"
End If

If ReadMe = "Yes" Then
%>
<html>
<title>WAP浏览器</title>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="160" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr valign="top"><td><img src="Images/WapBack01.gif" width="160" height="48"></td>
  </tr>
  <tr height="140">
    <td height="153" valign="middle" background="Images/WapBack02.gif">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="2" colspan="3"></td>
        </tr>
        <tr>
          <td width="30"></td>
          <td width="112" valign='top' style="font-size: 9pt;word-break:break-all;Width:fixed"><font color="#FFFFFF">温馨提示：本站已开通WAP服务，若您的手机支持WAP功能，可以使用手机访问：<br><% =WapDomain%></font></td>
          <td width="18">&nbsp;</td>
        </tr>
        <tr>
          <td height="2" colspan="3"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr><td><img src="Images/WapBack03.gif" width="160" height="56"></td></tr>
</table>
</body>
</html>
<%
Else
    Response.ContentType = "text/vnd.wap.wml; charset=utf-8"
    Call main
End If
Set XmlDoc = Nothing
Call CloseConn

'类私有变量
Private PhoneNumber, PhoneType, WapDomain, strHTML, Source, SiteLogo
Sub main()

    strHTML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    strHTML = strHTML & "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml1_1.1.xml"">" & vbCrLf
    strHTML = strHTML & "<wml>" & vbCrLf
    strHTML = strHTML & "<head>" & vbCrLf
    strHTML = strHTML & "<meta http-equiv=""Cache-control"" content=""max-age=0"" forua=""true""/>" & vbCrLf
    strHTML = strHTML & "<meta http-equiv=""Cache-control"" content=""must-revalidate"" forua=""true""/>" & vbCrLf
    strHTML = strHTML & "</head>" & vbCrLf
    strHTML = strHTML & "<template>"
    strHTML = strHTML & "<do type=""prev"" label=""" & XmlText("Wap", "BackBotton", "返回") & """>"
    strHTML = strHTML & "<prev/>"
    strHTML = strHTML & "</do>"
    strHTML = strHTML & "</template>"
    
    If WapLogo = "0" Then
        SiteLogo = "=<strong>" & SiteName & "</strong>="
    Else
        SiteLogo = "<img alt=""LOGO"" src=""" & WapLogo & """/>"
    End If

    '获得号码
    PhoneNumber = Request.ServerVariables("HTTP_X_UP_CALLING_LINE_ID")

    '获得手机型号

    PhoneType = Request.ServerVariables("HTTP_USER_AGENT")
    'If PhoneNumber = "" Then
    '    PhoneNumber = "非手机访问"
    'End If
    Source = Trim(Request("Source"))
    If FoundErr = True Then
        strHTML = strHTML & "<card id=""main"" title=""Welcome"">" & vbCrLf
        strHTML = strHTML & "<p>" & XmlText("Wap", "CloseEd", "本站已关闭WAP功能！") & "</p>" & vbCrLf
        strHTML = strHTML & "</card>" & vbCrLf
    Else
        If Source = "" Then
            Call ShowWap(0, 0, 0, 0)
        Else
            Source = ReplaceBadChar(Source)
            Call ProSource(Source)
        End If
    End If
    strHTML = strHTML & "</wml>" & vbCrLf
    Response.Write unicode(strHTML)
End Sub

'**************************************************
'函数名：ProSource
'作  用：参数处理
'**************************************************
Sub ProSource(ByVal iText)
    Dim StrRow, Mtype, ChannelID, ArticleID, ClassID
    StrRow = Split(iText, "|")
    Action = StrRow(0)
    ChannelID = StrRow(1)
    If ChannelID = "" Then
        ChannelID = 0
    Else
        ChannelID = PE_CLng(ChannelID)
    End If

    If ChannelID > 0 Then
        GetChannel (ChannelID)
    End If

    Select Case Action
    Case "ChannelList"
        Call ShowWap(ChannelID, 0, 0, 0)
    Case "ClassList"
        ClassID = StrRow(2)
        If ClassID = "" Then
            ClassID = 0
        Else
            ClassID = PE_CLng(ClassID)
        End If
        Call ShowWap(ChannelID, ClassID, 0, 0)
    Case "ShowArticle"
        ArticleID = StrRow(2)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call ShowArticle(ChannelID, ArticleID, StrRow(3))
    Case "ShowSoft"
        ArticleID = StrRow(2)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call ShowSoft(32, ChannelID, ArticleID)
    Case "ShowPhoto"
        ArticleID = StrRow(2)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call ShowPhoto(32, ChannelID, ArticleID)
    Case "ShowProduct"
        ArticleID = StrRow(2)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call ShowProduct(32, ChannelID, ArticleID)
    Case "AComment"
        Mtype = StrRow(2)
        If Mtype = "" Then
            Mtype = 1
        Else
            Mtype = PE_CLng(Mtype)
        End If
        ClassID = StrRow(3)
        If ClassID = "" Then
            ClassID = 0
        Else
            ClassID = PE_CLng(ClassID)
        End If
        ArticleID = StrRow(4)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call Comment(0, ArticleID, ChannelID, Mtype, ClassID)
    Case "AComment2"
        Mtype = StrRow(2)
        If Mtype = "" Then
            Mtype = 1
        Else
            Mtype = PE_CLng(Mtype)
        End If
        ClassID = StrRow(3)
        If ClassID = "" Then
            ClassID = 1
        Else
            ClassID = PE_CLng(ClassID)
        End If
        ArticleID = StrRow(4)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call Comment(1, ArticleID, ChannelID, Mtype, ClassID)
    Case "CommentSave"
        Mtype = StrRow(2)
        If Mtype = "" Then
            Mtype = 1
        Else
            Mtype = PE_CLng(Mtype)
        End If
        ArticleID = StrRow(3)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call CommentSave(ChannelID, Mtype, ArticleID, StrRow(4), StrRow(5), StrRow(6))
    Case "AFuJian"
        ArticleID = StrRow(2)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call Appendix(ChannelID, ArticleID)
    Case "dg"
        ArticleID = StrRow(2)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call dg(ChannelID, ArticleID)
    Case "dgacept"
        ArticleID = StrRow(2)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Dim iID
        If StrRow(3) = "" Then
            iID = 0
        Else
            iID = PE_CLng(StrRow(3))
        End If
        Call dgacept(ChannelID, ArticleID, iID, StrRow(4), StrRow(5), StrRow(6), StrRow(7))
    Case "getjynum"
        ArticleID = StrRow(2)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call getjynum(ChannelID, ArticleID, 1, "none", "none")
    Case "getjynum2"
        ArticleID = StrRow(2)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call getjynum(ChannelID, ArticleID, 2, StrRow(3), StrRow(4))
    Case "ManageLogin"
        Call ManageLogin(ChannelID, StrRow(2), StrRow(3))
    Case "ChannelManage"
        Call ChannelManage(ChannelID, StrRow(2), StrRow(3))
    Case "ArticlePass"
        ArticleID = StrRow(2)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call ArticlePass(ChannelID, ArticleID, StrRow(3), StrRow(4))
    Case "SoftPass"
        ArticleID = StrRow(2)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call SoftPass(ChannelID, ArticleID, StrRow(3), StrRow(4))
    Case "PhotoPass"
        ArticleID = StrRow(2)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call PhotoPass(ChannelID, ArticleID, StrRow(3), StrRow(4))
    Case "GuestPass"
        ArticleID = StrRow(2)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call GuestPass(ChannelID, ArticleID, StrRow(3), StrRow(4))
    Case "ProductPass"
        ArticleID = StrRow(2)
        If ArticleID = "" Then
            ArticleID = 0
        Else
            ArticleID = PE_CLng(ArticleID)
        End If
        Call ProductPass(ChannelID, ArticleID, StrRow(3), StrRow(4))
    End Select
End Sub


'**************************************************
'前台浏览部分开始
'**************************************************
Sub ShowWap(ByVal iChannelID, ByVal iClassID, ByVal iHot, ByVal iElite)
    Dim sqlChannel, rsChannel, sqlArticle, rsArticle, ModuleType, HitsOfHot
    strHTML = strHTML & "<card id=""main"" title=""" & SiteName & """>" & vbCrLf
    
    If iChannelID = 0 Then '如果是显示首页
        sqlChannel = "select ChannelID,OrderID,ChannelName,ChannelDir,ModuleType,ChannelType,Disabled from PE_Channel where Disabled = " & PE_False & " and ChannelType<2 order by OrderID"
        Set rsChannel = Conn.Execute(sqlChannel)
        If rsChannel.BOF And rsChannel.EOF Then
            strHTML = strHTML & "<p align=""center"">" & XmlText("BaseText", "ChannelErr", "找不到频道！")
        Else
            strHTML = strHTML & "<p align=""center"">" & SiteLogo & vbCrLf
            Do While Not rsChannel.EOF
                If rsChannel("ModuleType") = 1 Or rsChannel("ModuleType") = 2 Or rsChannel("ModuleType") = 3 Or rsChannel("ModuleType") = 5 Then
                    If rsChannel("ModuleType") = 5 Then
                        strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ChannelList|" & rsChannel("ChannelID") & """>" & rsChannel("ChannelName") & "</a>" & vbCrLf
                    Else
                        strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ChannelList|" & rsChannel("ChannelID") & """>" & rsChannel("ChannelName") & "</a>" & vbCrLf
                    End If
                End If
                rsChannel.MoveNext
            Loop
            If ShowWapManage = True Then strHTML = strHTML & "<br/>-----<br/><a href=""" & WapDomain & "?Source=ManageLogin|1|none|none"">" & XmlText("Wap", "ManageLogin", "-管理登录-") & "</a>" & vbCrLf
        End If
        strHTML = strHTML & "</p>" & vbCrLf
        rsChannel.Close
        Set rsChannel = Nothing
    Else
        strHTML = strHTML & "<p>" & XmlText("Wap", "News", "-最新更新-") & vbCrLf
        Set rsChannel = Conn.Execute("select ChannelName,ChannelDir,ModuleType,HitsOfHot,UploadDir from PE_Channel where ChannelID=" & iChannelID & " and Disabled = " & PE_False & " and ChannelType<2 order by OrderID")
        ChannelName = rsChannel("ChannelName")
        ChannelDir = rsChannel("ChannelDir")
        ModuleType = rsChannel("ModuleType")
        HitsOfHot = rsChannel("HitsOfHot")
        UploadDir = rsChannel("UploadDir")
        rsChannel.Close
        Set rsChannel = Nothing
        Select Case ModuleType
        Case 1
            sqlArticle = "select top 12 A.ArticleID,A.ChannelID,A.ClassID,A.Title,A.Hits,A.UpdateTime,A.Elite,A.Status,A.IncludePic,A.LinkUrl,A.Deleted,C.ClassPurview from PE_Article A inner join PE_Class C on A.ClassID=C.ClassID Where A.ChannelID=" & iChannelID & " and C.ClassPurview<2"
            If iClassID <> 0 Then sqlArticle = sqlArticle & " and A.ClassID=" & iClassID
            sqlArticle = sqlArticle & " and A.Status=3 and A.Deleted=" & PE_False
            If iHot = 1 Then
                sqlArticle = sqlArticle & " order by A.Hits Desc"
            ElseIf iElite = 1 Then
                sqlArticle = sqlArticle & " order by A.Elite " & PE_OrderType & ",A.UpdateTime Desc"
            Else
                sqlArticle = sqlArticle & " order by A.UpdateTime Desc"
            End If
            Set rsArticle = Conn.Execute(sqlArticle)
            If Not (rsArticle.BOF And rsArticle.EOF) Then
                Do While Not rsArticle.EOF
                    If rsArticle(9) = "" Then
                        strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ShowArticle|" & iChannelID & "|" & rsArticle(0) & "|0"">" & ReplaceText(GetSubStr(xml_nohtml(rsArticle(3)), 20, False), 2) & "</a>"
                        If rsArticle(8) > 0 Then strHTML = strHTML & "-图" & vbCrLf
                        If rsArticle(4) > HitsOfHot Then strHTML = strHTML & "-热" & vbCrLf
                        If rsArticle(6) = True Then strHTML = strHTML & "-荐" & vbCrLf
                    End If
                    rsArticle.MoveNext
                Loop
            Else
                strHTML = strHTML & "无文章" & vbCrLf
            End If
            rsArticle.Close
        Case 2
                sqlArticle = "select top 12 A.SoftID,A.ChannelID,A.ClassID,A.SoftName,A.Hits,A.UpdateTime,A.Elite,A.Status,A.Deleted,C.ClassPurview from PE_Soft A inner join PE_Class C on A.ClassID=C.ClassID Where A.ChannelID=" & iChannelID & " and C.ClassPurview<2"
                If iClassID <> 0 Then sqlArticle = sqlArticle & " and A.ClassID=" & iClassID
                sqlArticle = sqlArticle & " and A.Status=3 and A.Deleted=" & PE_False
                If iHot = 1 Then
                    sqlArticle = sqlArticle & " order by A.Hits Desc"
                ElseIf iElite = 1 Then
                    sqlArticle = sqlArticle & " order by A.Elite " & PE_OrderType & ",A.UpdateTime Desc"
                Else
                    sqlArticle = sqlArticle & " order by A.UpdateTime Desc"
                End If
                Set rsArticle = Conn.Execute(sqlArticle)
                If Not (rsArticle.BOF And rsArticle.EOF) Then
                    Do While Not rsArticle.EOF
                        strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ShowSoft|" & iChannelID & "|" & rsArticle(0) & """>" & GetSubStr(xml_nohtml(rsArticle(3)), 20, False) & "</a>"
                        If rsArticle(6) = True Then strHTML = strHTML & "-荐" & vbCrLf
                        If rsArticle(4) > HitsOfHot Then strHTML = strHTML & "-热" & vbCrLf
                        rsArticle.MoveNext
                    Loop
                Else
                    strHTML = strHTML & "无下载" & vbCrLf
                End If
                rsArticle.Close
        Case 3
                sqlArticle = "select top 12 A.PhotoID,A.ChannelID,A.ClassID,A.PhotoName,A.Hits,A.UpdateTime,A.Elite,A.Status,A.Deleted,C.ClassPurview from PE_Photo A inner join PE_Class C on A.ClassID=C.ClassID Where A.ChannelID=" & iChannelID & " and C.ClassPurview<2"
                If iClassID <> 0 Then sqlArticle = sqlArticle & " and A.ClassID=" & iClassID
                sqlArticle = sqlArticle & " and A.Status=3 and A.Deleted=" & PE_False
                If iHot = 1 Then
                    sqlArticle = sqlArticle & " order by A.Hits Desc"
                ElseIf iElite = 1 Then
                    sqlArticle = sqlArticle & " order by A.Elite " & PE_OrderType & ",A.UpdateTime Desc"
                Else
                    sqlArticle = sqlArticle & " order by A.UpdateTime Desc"
                End If
                Set rsArticle = Conn.Execute(sqlArticle)
                If Not (rsArticle.BOF And rsArticle.EOF) Then
                    Do While Not rsArticle.EOF
                        strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ShowPhoto|" & iChannelID & "|" & rsArticle(0) & """>" & GetSubStr(xml_nohtml(rsArticle(3)), 20, False) & "</a>"
                        If rsArticle(6) = True Then strHTML = strHTML & "-荐" & vbCrLf
                        If rsArticle(4) > HitsOfHot Then strHTML = strHTML & "-热" & vbCrLf
                        rsArticle.MoveNext
                    Loop
                Else
                    strHTML = strHTML & "无图片" & vbCrLf
                End If
                rsArticle.Close
        Case 5
                sqlArticle = "select top 12 ProductID,ChannelID,ClassID,ProductName,IsHot,IsElite,UpdateTime,Hits,Deleted from PE_Product Where ChannelID=" & iChannelID
                If iClassID <> 0 Then sqlArticle = sqlArticle & " and ClassID=" & iClassID
                sqlArticle = sqlArticle & " and Deleted=" & PE_False
                If iHot = 1 Then
                    sqlArticle = sqlArticle & " order by Hits Desc"
                ElseIf iElite = 1 Then
                    sqlArticle = sqlArticle & " order by IsElite " & PE_OrderType & ",UpdateTime Desc"
                Else
                    sqlArticle = sqlArticle & " order by UpdateTime Desc"
                End If
                Set rsArticle = Conn.Execute(sqlArticle)
                If Not (rsArticle.BOF And rsArticle.EOF) Then
                    Do While Not rsArticle.EOF
                        strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ShowProduct|" & iChannelID & "|" & rsArticle(0) & """>" & GetSubStr(xml_nohtml(rsArticle(3)), 20, False) & "</a>"
                        If rsArticle(4) = True Then strHTML = strHTML & "-热" & vbCrLf
                        If rsArticle(5) = True Then strHTML = strHTML & "-荐" & vbCrLf
                        rsArticle.MoveNext
                    Loop
                Else
                    strHTML = strHTML & "无商品" & vbCrLf
                End If
                rsArticle.Close
        End Select
        Set rsArticle = Nothing
        strHTML = strHTML & "</p>" & vbCrLf
        strHTML = strHTML & GetChildClass(iChannelID, iClassID)
    End If
    strHTML = strHTML & "</card>" & vbCrLf
End Sub

Function GetChildClass(ByVal iChannelID, ByVal iClassID)
    Dim rsClass, strtmp
    strtmp = "<p>-子栏目-" & vbCrLf
    If iClassID = 0 Then
        Set rsClass = Conn.Execute("select ClassID,ClassName,Child from PE_Class where ChannelID=" & iChannelID & " and ClassType=1 and ParentID =0")
        If Not (rsClass.BOF And rsClass.EOF) Then
            Do While Not rsClass.EOF
                strtmp = strtmp & "<br/>[<a href=""" & WapDomain & "?Source=ClassList|" & iChannelID & "|" & rsClass("ClassID") & """>" & rsClass("ClassName") & "</a>]" & vbCrLf
                rsClass.MoveNext
            Loop
        End If
        strtmp = strtmp & "<br/>[<a href=""" & WapDomain & """>首页</a>]" & vbCrLf
    Else
        Set rsClass = Conn.Execute("select ClassID,ClassName,Child from PE_Class where ParentID=" & iClassID & " and ClassType=1 order by OrderID")
        If rsClass.BOF And rsClass.EOF Then
            strtmp = strtmp & "<br/>[<a href=""" & WapDomain & """>首页</a>]" & vbCrLf
        Else
            Do While Not rsClass.EOF
                strtmp = strtmp & "<br/>[<a href=""" & WapDomain & "?Source=ClassList|" & iChannelID & "|" & rsClass("ClassID") & """>" & rsClass("ClassName") & "</a>]" & vbCrLf
                rsClass.MoveNext
            Loop
            strtmp = strtmp & "<br/>[<a href=""" & WapDomain & """>首页</a>]" & vbCrLf
        End If
    End If
    rsClass.Close
    Set rsClass = Nothing
    GetChildClass = strtmp & "</p>" & vbCrLf
End Function

'**************************************************
'函数名：ShowArticle
'作  用：显示文章内容
'**************************************************
Sub ShowArticle(ByVal iChannelID, ByVal iArticleID, ByVal iPage)
    Dim sqlArticle, rsArticle
    strHTML = strHTML & "<card id=""main"" title=""" & SiteName & """>" & vbCrLf
    If iArticleID = 0 Then
        strHTML = strHTML & "<p>无此文章！</p>" & vbCrLf
    Else
        sqlArticle = "select * from PE_Article Where ArticleID=" & iArticleID & " and Status=3 and Deleted=" & PE_False & " and InfoPoint=0"
        Set rsArticle = Conn.Execute(sqlArticle)
        If rsArticle.BOF And rsArticle.EOF Then
            strHTML = strHTML & "<p>收费文章，请登录网站浏览！" & vbCrLf
            strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ChannelList|" & iChannelID & """>返回</a></p>" & vbCrLf
        Else
            strHTML = strHTML & "<p>" & getpage(iChannelID, iArticleID, ReplaceText(Replace(xml_nohtml(rsArticle("Content")), "[NextPage]", ""), 1), iPage) & vbCrLf
            If EnableWapPl = True Then strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=AComment|" & rsArticle("ChannelID") & "|1|" & rsArticle("ClassID") & "|" & iArticleID & """>评论</a>" & vbCrLf
            If ShowWapAppendix = True Then
                If rsArticle("IncludePic") > 0 Then strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=AFuJian|" & iChannelID & "|" & iArticleID & """>图片</a>" & vbCrLf
            End If
            strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ChannelList|" & iChannelID & """>返回</a></p>" & vbCrLf
        End If
        rsArticle.Close
        Set rsArticle = Nothing
    End If
    strHTML = strHTML & "</card>" & vbCrLf
End Sub

'**************************************************
'函数名：ShowSoft
'作  用：显示下载内容
'**************************************************
Sub ShowSoft(ByVal iSize, ByVal iChannelID, ByVal iSoftID)
    Dim sqlSoft, rsSoft
    strHTML = strHTML & "<card id=""main"" title=""" & SiteName & """>" & vbCrLf
    If iSoftID = 0 Then
        strHTML = strHTML & "<p>无此下载！</p>" & vbCrLf
    Else
        sqlSoft = "select * from PE_Soft Where SoftID=" & iSoftID & " and Status=3 and Deleted=" & PE_False & " and InfoPoint=0"
        Set rsSoft = Conn.Execute(sqlSoft)
        If rsSoft.BOF And rsSoft.EOF Then
            strHTML = strHTML & "<p>收费软件，请登录网站下载！</p>" & vbCrLf
        Else
            strHTML = strHTML & "<p>名称:" & GetSubStr2(xml_nohtml(rsSoft("SoftName")), iSize) & "<br/>" & vbCrLf
            If Not IsNull(rsSoft("SoftVersion")) Then strHTML = strHTML & "版本:" & GetSubStr2(xml_nohtml(rsSoft("SoftVersion")), iSize) & "<br/>" & vbCrLf
            If Not IsNull(rsSoft("SoftIntro")) Then strHTML = strHTML & "简介:" & GetSubStr2(xml_nohtml(rsSoft("SoftIntro")), 80) & "<br/>" & vbCrLf
            strHTML = strHTML & GetDownloadUrlList(rsSoft("DownloadUrl"))
            If EnableWapPl = True Then strHTML = strHTML & "<a href=""" & WapDomain & "?Source=AComment|" & rsSoft("ChannelID") & "|2|" & rsSoft("ClassID") & "|" & iSoftID & """>评论</a>" & vbCrLf
            strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ChannelList|" & iChannelID & """>返回</a></p>" & vbCrLf
        End If
        rsSoft.Close
        Set rsSoft = Nothing
    End If
    strHTML = strHTML & "</card>" & vbCrLf
End Sub

'**************************************************
'函数名：ShowPhoto
'作  用：显示图片内容
'**************************************************
Sub ShowPhoto(ByVal iSize, ByVal iChannelID, ByVal iPhotoID)
    Dim sqlPhoto, rsPhoto
    strHTML = strHTML & "<card id=""main"" title=""" & SiteName & """>" & vbCrLf
    If iPhotoID = 0 Then
        strHTML = strHTML & "<p>无此图片！</p>" & vbCrLf
    Else
        sqlPhoto = "select * from PE_Photo Where PhotoID=" & iPhotoID & " and Status=3 and Deleted=" & PE_False & " and InfoPoint=0"
        Set rsPhoto = Conn.Execute(sqlPhoto)
        If rsPhoto.BOF And rsPhoto.EOF Then
            strHTML = strHTML & "<p>收费图片，请登录网站浏览！</p>" & vbCrLf
        Else
            strHTML = strHTML & "<p>" & GetSubStr2(xml_nohtml(rsPhoto("PhotoName")), iSize) & "<br/>" & vbCrLf
            If ShowWapAppendix = True Then
                If rsPhoto("PhotoThumb") > "" Then
                    If Left(rsPhoto("PhotoThumb"), 4) = "http" Then
                        strHTML = strHTML & "<img alt=""图片预览"" src=""" & rsPhoto("PhotoThumb") & """/><br/>" & vbCrLf
                    Else
                        strHTML = strHTML & "<img alt=""图片预览"" src=""" & SiteUrl & "/" & ChannelDir & "/" & UploadDir & "/" & rsPhoto("PhotoThumb") & """/><br/>" & vbCrLf
                    End If
                End If
            End If
            If EnableWapPl = True Then strHTML = strHTML & "<a href=""" & WapDomain & "?Source=AComment|" & rsPhoto("ChannelID") & "|3|" & rsPhoto("ClassID") & "|" & iPhotoID & """>评论</a>" & vbCrLf
            strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ChannelList|" & iChannelID & """>返回</a></p>" & vbCrLf
        End If
        rsPhoto.Close
        Set rsPhoto = Nothing
    End If
    strHTML = strHTML & "</card>" & vbCrLf
End Sub

'**************************************************
'函数名：ShowProduct
'作  用：显示商品内容
'**************************************************
Sub ShowProduct(ByVal iSize, ByVal iChannelID, ByVal iProductID)
    Dim sqlProduct, rsProduct
    strHTML = strHTML & "<card id=""main"" title=""" & SiteName & """>" & vbCrLf
    If iProductID = 0 Then
        strHTML = strHTML & "<p>无此商品！</p>" & vbCrLf
    Else
        sqlProduct = "select * from PE_Product Where ProductID=" & iProductID & " and EnableSale=" & PE_True & " and Deleted=" & PE_False & " and Stocks>0"
        Set rsProduct = Conn.Execute(sqlProduct)
        If rsProduct.BOF And rsProduct.EOF Then
            strHTML = strHTML & "<p>无此商品！</p>" & vbCrLf
        Else
            strHTML = strHTML & "<p>名称:" & GetSubStr2(xml_nohtml(rsProduct("ProductName")), iSize) & "<br/>" & vbCrLf
            strHTML = strHTML & "厂商:" & GetSubStr2(xml_nohtml(rsProduct("ProducerName")), iSize) & "<br/>" & vbCrLf
            strHTML = strHTML & "品牌:" & GetSubStr2(xml_nohtml(rsProduct("TrademarkName")), iSize) & "<br/>" & vbCrLf
            strHTML = strHTML & "原价:" & rsProduct("Price_Original") & "<br/>" & vbCrLf
            strHTML = strHTML & "现价:" & rsProduct("Price") & "<br/>" & vbCrLf
            If Not IsNull(rsProduct("ProductIntro")) Then strHTML = strHTML & "简介:" & GetSubStr2(xml_nohtml(rsProduct("ProductIntro")), 160) & "<br/>" & vbCrLf
            If ShowWapAppendix = True Then
                If rsProduct("ProductThumb") > "" Then
                    If Left(rsProduct("ProductThumb"), 4) = "http" Then
                        strHTML = strHTML & "<img alt=""产品图片"" src=""" & rsProduct("ProductThumb") & """/><br/>" & vbCrLf
                    Else
                        strHTML = strHTML & "<img alt=""产品图片"" src=""" & SiteUrl & "/" & ChannelDir & "/" & UploadDir & "/" & rsProduct("ProductThumb") & """/><br/>" & vbCrLf
                    End If
                End If
            End If
            If EnableWapPl = True Then strHTML = strHTML & "<a href=""" & WapDomain & "?Source=AComment|" & rsProduct("ChannelID") & "|5|" & rsProduct("ClassID") & "|" & iProductID & """>评论</a>" & vbCrLf
            If ShowWapShop = True Then strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=dg|" & iChannelID & "|" & iProductID & """>订购</a>" & vbCrLf
            strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ChannelList|" & iChannelID & """>返回</a></p>" & vbCrLf
        End If
        rsProduct.Close
        Set rsProduct = Nothing
    End If
    strHTML = strHTML & "</card>" & vbCrLf
End Sub

'**************************************************
'函数名：Comment
'作  用：显示或发表评论
'**************************************************
Sub Comment(ByVal iType, ByVal iID, ByVal iChannelID, ByVal iModuleType, ByVal iClassID)
    Dim sqlComment, rsComment, rsClass
    strHTML = strHTML & "<card id=""main"" title=""" & SiteName & """>" & vbCrLf
    If iID = 0 Then
        strHTML = strHTML & "<p>无此对象！</p>" & vbCrLf
    Else
        If iType = 0 Then
            sqlComment = "select top 10 * from PE_Comment where InfoID=" & iID & " and Passed=" & PE_True & " order by WriteTime desc"
            Set rsComment = Conn.Execute(sqlComment)
            If rsComment.BOF And rsComment.EOF Then
                strHTML = strHTML & "<p>没有评论！<br/>" & vbCrLf
                strHTML = strHTML & "<a href=""" & WapDomain & "?Source=AComment2|" & iChannelID & "|" & iModuleType & "|" & iClassID & "|" & iID & """>发表</a></p>" & vbCrLf
            Else
                strHTML = strHTML & "<p>"
                Do While Not rsComment.EOF
                    strHTML = strHTML & xml_nohtml(rsComment("UserName")) & "的评论：<br/>" & vbCrLf
                    strHTML = strHTML & GetSubStr2(xml_nohtml(rsComment("Content")), 40) & "<br/>" & vbCrLf
                    strHTML = strHTML & FormatDateTime(rsComment("WriteTime"), 2) & "<br/>" & vbCrLf
                    rsComment.MoveNext
                Loop
                strHTML = strHTML & "<a href=""" & WapDomain & "?Source=AComment2|" & iChannelID & "|" & iModuleType & "|" & iClassID & "|" & iID & """>发表</a></p>" & vbCrLf
            End If
            rsComment.Close
            Set rsComment = Nothing
        Else
            Set rsClass = Conn.Execute("select EnableComment,CheckComment from PE_Class Where ClassID=" & iClassID)
            If rsClass.BOF And rsClass.EOF Then
                strHTML = strHTML & "<p>本栏禁止评论！</p>" & vbCrLf
            Else
                If rsClass("EnableComment") = True Then
                    strHTML = strHTML & "<p>您的姓名:<br/>" & vbCrLf
                    strHTML = strHTML & "<input name=""namer"" emptyok=""false"" value=""" & PhoneNumber & """/><br/>" & vbCrLf
                    strHTML = strHTML & "评论:<br/>" & vbCrLf
                    strHTML = strHTML & "<input name=""Comcont"" emptyok=""false""/>" & vbCrLf
                    If rsClass("CheckComment") = True Then
                        strHTML = strHTML & "<a href=""" & WapDomain & "?Source=CommentSave|" & iChannelID & "|" & iModuleType & "|" & iID & "|$(namer)|$(Comcont)|0"">提交</a></p>" & vbCrLf
                    Else
                        strHTML = strHTML & "<a href=""" & WapDomain & "?Source=CommentSave|" & iChannelID & "|" & iModuleType & "|" & iID & "|$(namer)|$(Comcont)|1"">提交</a></p>" & vbCrLf
                    End If
                Else
                    strHTML = strHTML & "<p>本栏禁止评论</p>" & vbCrLf
                End If
            End If
            rsClass.Close
            Set rsClass = Nothing
        End If
    End If
    strHTML = strHTML & "</card>" & vbCrLf
End Sub

'**************************************************
'函数名：CommentSave
'作  用：保存评论
'**************************************************
Sub CommentSave(ByVal iChannelID, ByVal iModuleType, ByVal iID, ByVal iName, ByVal iComcont, ByVal iCheck)
    Dim sqlComment, rsComment
    strHTML = strHTML & "<card id=""main"" title=""" & SiteName & """>" & vbCrLf
    If iName = "" Or iComcont = "" Then
        strHTML = strHTML & "<p>请填写内容!" & vbCrLf
    Else
        sqlComment = "Select * from PE_Comment"
        Set rsComment = Server.CreateObject("Adodb.RecordSet")
        rsComment.Open sqlComment, Conn, 1, 3
            rsComment.addnew
            rsComment("ModuleType") = iModuleType
            rsComment("InfoID") = iID
            rsComment("UserType") = 0
            rsComment("UserName") = UTF2GB(iName)
            rsComment("Sex") = 0
            rsComment("WriteTime") = Now()
            rsComment("Score") = 3
            rsComment("Content") = UTF2GB(iComcont)
            If iCheck = 1 Then
                rsComment("Passed") = True
            Else
                rsComment("Passed") = False
            End If
            rsComment.Update
        rsComment.Close
        Set rsComment = Nothing
        strHTML = strHTML & "<p>评论发表成功!" & vbCrLf
    End If
    strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ChannelList|" & iChannelID & """>返回</a></p>" & vbCrLf
    strHTML = strHTML & "</card>" & vbCrLf
End Sub

'**************************************************
'函数名：Appendix
'作  用：显示文章图片
'**************************************************
Sub Appendix(ByVal iChannelID, ByVal iID)
    Dim rsFJ
    strHTML = strHTML & "<card id=""main"" title=""" & SiteName & """>" & vbCrLf
    If iID = 0 Then
        strHTML = strHTML & "<p>无图片！" & vbCrLf
        strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ShowArticle|" & iChannelID & "|" & iID & "|0"">返回</a></p>" & vbCrLf
    Else
        strHTML = strHTML & "<p>" & vbCrLf
        Set rsFJ = Conn.Execute("select DefaultPicUrl from PE_Article Where ArticleID=" & iID)
        If Left(LCase(rsFJ("DefaultPicUrl")), 4) = "http" Then
            strHTML = strHTML & "<img alt=""图片"" src=""" & rsFJ("DefaultPicUrl") & """/>" & vbCrLf
        Else
            strHTML = strHTML & "<img alt=""图片"" src=""" & SiteUrl & "/" & ChannelDir & "/" & UploadDir & "/" & rsFJ("DefaultPicUrl") & """/>" & vbCrLf
        End If
        rsFJ.Close
        Set rsFJ = Nothing
        strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ShowArticle|" & iChannelID & "|" & iID & "|0"">返回</a></p>" & vbCrLf
    End If
    strHTML = strHTML & "</card>" & vbCrLf
End Sub

'**************************************************
'函数名：getjynum
'作  用：取得用户交易码
'**************************************************
Sub getjynum(ByVal iChannelID, ByVal iID, ByVal iType, ByVal iUser, ByVal iPass)
    strHTML = strHTML & "<card id=""main"" title=""" & SiteName & """>" & vbCrLf
    If iType = 1 Then
        strHTML = strHTML & "<p>用户名:<br/>" & vbCrLf
        strHTML = strHTML & "<input name=""username"" emptyok=""false""/><br/>" & vbCrLf
        strHTML = strHTML & "密码:<br/>" & vbCrLf
        strHTML = strHTML & "<input name=""password"" emptyok=""false""/><br/>" & vbCrLf
        strHTML = strHTML & "<a href=""" & WapDomain & "?Source=getjynum2|" & iChannelID & "|" & iID & "|$(username)|$(password)"">提交</a><br/></p>" & vbCrLf
    Else
        Dim rsUser, sqlUser
        sqlUser = "select UserName,UserPassword,CheckNum from PE_User Where UserName='" & UTF2GB(iUser) & "' and UserPassword='" & MD5(iPass, 16) & "'"
        Set rsUser = Conn.Execute(sqlUser)
        If rsUser.BOF And rsUser.EOF Then
            strHTML = strHTML & "<p>用户名或密码错！</p>" & iUser & MD5(iPass, 16) & vbCrLf
        Else
            strHTML = strHTML & "<p>" & rsUser("UserName") & "您好,您在本站的交易码是:<br/>" & rsUser("CheckNum") & "<br/><a href=""" & WapDomain & "?Source=dgacept|" & iChannelID & "|" & iID & "|" & rsUser("CheckNum") & "|1|none|none|none"">继续</a></p>" & vbCrLf
        End If
        rsUser.Close
        Set rsUser = Nothing
    End If
    strHTML = strHTML & "</card>" & vbCrLf
End Sub

'**************************************************
'函数名：dg
'作  用：订购商品
'**************************************************
Sub dg(ByVal iChannelID, ByVal iID)
    strHTML = strHTML & "<card id=""main"" title=""" & SiteName & """>" & vbCrLf
    If iID = 0 Then
        strHTML = strHTML & "<p>参数错！<br/>" & vbCrLf
    Else
        strHTML = strHTML & "<p>订购数量:<br/>" & vbCrLf
        strHTML = strHTML & "<input name=""shuliang"" emptyok=""false"" value=""1"" maxlength=""3""/><br/>" & vbCrLf
        strHTML = strHTML & "您的交易码:<br/>" & vbCrLf
        strHTML = strHTML & "<input name=""number"" emptyok=""false"" maxlength=""8""/><br/>" & vbCrLf
        strHTML = strHTML & "<a href=""" & WapDomain & "?Source=dgacept|" & iChannelID & "|" & iID & "|$(number)|$(shuliang)|none|none|none"">提交</a><br/>" & vbCrLf
        strHTML = strHTML & "-----<br/>" & vbCrLf
        strHTML = strHTML & "<a href=""" & WapDomain & "?Source=getjynum|" & iChannelID & "|" & iID & """>忘记交易码</a></p>" & vbCrLf
    End If
    strHTML = strHTML & "</card>" & vbCrLf
End Sub

'**************************************************
'函数名：dgacept
'作  用：确认订购商品
'**************************************************
Sub dgacept(ByVal iChannelID, ByVal iID, ByVal iNumber, ByVal ishuliang, ByVal iaddress, ByVal izipcode, ByVal iemail)
    Dim trs, sqlOrder, rsOrder, rsItem, rsUser, rsProduct, OrderFormID, OrderFormNum
    '得到订单ID
    Set trs = Conn.Execute("select max(OrderFormID) from PE_OrderForm")
    If trs.BOF And trs.EOF Then
        OrderFormID = 0
    Else
        If IsNull(trs(0)) Then
            OrderFormID = 1
        Else
            OrderFormID = trs(0) + 1
        End If
    End If
    Set trs = Nothing
    '得到订单编号
    OrderFormNum = Prefix_OrderFormNum & GetNumString()
    
    strHTML = strHTML & "<card id=""main"" title=""" & SiteName & """>" & vbCrLf
    If iID = 0 Then
        strHTML = strHTML & "<p>参数错！<br/>" & vbCrLf
    Else
        If iNumber = 0 Or ishuliang = "" Then
            strHTML = strHTML & "<p>请填写内容!" & vbCrLf
        Else
            Set rsUser = Conn.Execute("select U.UserID,U.UserName,U.IsLocked,U.CheckNum,C.Address,C.Email,C.ZipCode,C.Mobile,C.OfficePhone,C.HomePhone,U.ClientID from PE_User U inner join PE_Contacter C on U.ContacterID=C.ContacterID Where U.CheckNum=" & iNumber)
            If rsUser.BOF And rsUser.EOF Then
                strHTML = strHTML & "<p>交易号码错！" & vbCrLf
            Else
                If rsUser(2) = True Then
                    strHTML = strHTML & "<p>您已被锁定无法完成交易!" & vbCrLf
                ElseIf IsNull(rsUser(4)) Or IsNull(rsUser(5)) Or IsNull(rsUser(6)) Then
                    strHTML = strHTML & "<p>收货地址:" & vbCrLf
                    strHTML = strHTML & "<br/><input name=""address"" emptyok=""false"" value=" & rsUser(4) & "/>" & vbCrLf
                    strHTML = strHTML & "<br/>邮政编码:" & vbCrLf
                    strHTML = strHTML & "<br/><input name=""zipcode"" emptyok=""false"" value=" & rsUser(6) & "/>" & vbCrLf
                    strHTML = strHTML & "<br/>电子邮件:" & vbCrLf
                    strHTML = strHTML & "<br/><input name=""email"" emptyok=""false"" value=" & rsUser(5) & "/>" & vbCrLf
                    strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=dgacept|" & iChannelID & "|" & iID & "|" & iNumber & "|" & ishuliang & "|$(address)|$(zipcode)|$(email)"">提交</a></p>" & vbCrLf
                Else
                    Set rsProduct = Conn.Execute("Select * from PE_Product Where ProductID= " & iID & " and Stocks>0")
                    If rsProduct.BOF And rsProduct.EOF Then
                        strHTML = strHTML & "<p>暂时无货！" & vbCrLf
                    Else
                        sqlOrder = "Select * from PE_OrderForm"
                        Set rsOrder = Server.CreateObject("Adodb.RecordSet")
                        rsOrder.Open sqlOrder, Conn, 1, 3
                            rsOrder.addnew
                            rsOrder("OrderFormID") = OrderFormID
                            rsOrder("OrderFormNum") = OrderFormNum
                            rsOrder("UserName") = rsUser(1)
                            rsOrder("ClientID") = rsUser(10)

                            If iaddress = "none" Then
                                rsOrder("Address") = rsUser(4)
                            Else
                                rsOrder("Address") = ConvChinese(iaddress)
                            End If

                            If izipcode = "none" Then
                                rsOrder("ZipCode") = rsUser(6)
                            Else
                                rsOrder("ZipCode") = izipcode
                            End If

                            If PhoneNumber = "" Then
                                rsOrder("Mobile") = rsUser(7)
                            Else
                                rsOrder("Mobile") = PhoneNumber
                            End If

                            If rsUser(8) = "" Then
                                rsOrder("Phone") = rsUser(9)
                            Else
                                rsOrder("Phone") = rsUser(8)
                            End If

                            If iemail = "none" Then
                                rsOrder("Email") = rsUser(5)
                            Else
                                rsOrder("Email") = ConvChinese(iemail)
                            End If

                            rsOrder("PaymentType") = 1
                            rsOrder("DeliverType") = 3
                            rsOrder("NeedInvoice") = False
                            rsOrder("InvoiceContent") = "发票抬头："
                            rsOrder("Invoiced") = False
                            rsOrder("Remark") = "本订单通过手机订购，请尽快联系客户"
                            rsOrder("MoneyTotal") = rsProduct("Price") * ishuliang
                            rsOrder("MoneyGoods") = 0
                            rsOrder("PresentMoney") = 0
                            rsOrder("PresentExp") = 0
                            rsOrder("MoneyReceipt") = 0
                            rsOrder("BeginDate") = FormatDateTime(Date, 2)
                            rsOrder("InputTime") = Now()
                            rsOrder("Status") = 1
                            rsOrder("DeliverStatus") = 1
                            rsOrder("EnableDownload") = False
                            rsOrder("Discount_Payment") = rsProduct("Discount")
                            rsOrder("Charge_Deliver") = 1
                            rsOrder.Update
                        rsOrder.Close

                        rsOrder.Open "select top 1 * from PE_OrderFormItem", Conn, 1, 3
                            '添加进订单详细表中
                            rsOrder.addnew
                            rsOrder("ItemID") = GetItemID()
                            rsOrder("OrderFormID") = OrderFormID
                            rsOrder("ProductID") = rsProduct("ProductID")
                            rsOrder("SaleType") = 1
                            rsOrder("Price_Original") = rsProduct("Price_Original")
                            rsOrder("Price") = rsProduct("Price")
                            rsOrder("TruePrice") = rsProduct("Price")
                            rsOrder("Amount") = ishuliang
                            rsOrder("Subtotal") = rsProduct("Price") * ishuliang
                            rsOrder("Remark") = "手机订购"
                            rsOrder("BeginDate") = FormatDateTime(Date, 2)
                            rsOrder("ServiceTerm") = rsProduct("ServiceTerm")
                            rsOrder("PresentExp") = rsProduct("PresentExp")
                            rsOrder.Update
                        rsOrder.Close
                        Set rsOrder = Nothing
                        strHTML = strHTML & "<p>成功!订单编号：<br/>" & vbCrLf
                        strHTML = strHTML & OrderFormNum & vbCrLf
                    End If
                    rsProduct.Close
                    Set rsProduct = Nothing
                End If
            End If
            rsUser.Close
            Set rsUser = Nothing
        End If
        strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ChannelList|" & iChannelID & """>返回</a></p>" & vbCrLf
    End If
    strHTML = strHTML & "</card>" & vbCrLf
End Sub

Function GetItemID()
    Dim trs
    Set trs = Conn.Execute("select max(ItemID) from PE_OrderFormItem")
    If IsNull(trs(0)) Then
        GetItemID = 1
    Else
        GetItemID = trs(0) + 1
    End If
    Set trs = Nothing
End Function

'**************************************************
'网站管理部分开始
'**************************************************
Sub ManageLogin(ByVal iStep, ByVal iUsername, ByVal iPassword)
    strHTML = strHTML & "<card id=""main"" title=""后台管理"">" & vbCrLf
    If iStep = 0 Or iStep = 1 Then
        strHTML = strHTML & "<p>用户名:<br/>" & vbCrLf
        strHTML = strHTML & "<input name=""username"" emptyok=""false""/><br/>" & vbCrLf
        strHTML = strHTML & "密码:<br/>" & vbCrLf
        strHTML = strHTML & "<input name=""password"" emptyok=""false""/><br/>" & vbCrLf
        strHTML = strHTML & "<a href=""" & WapDomain & "?Source=ManageLogin|2|$(username)|$(password)"">提交</a></p>" & vbCrLf
    ElseIf iStep = 2 Then
        Dim rsChannel, sqlChannel
        If CheckAdmin(iUsername, iPassword) = False Then
            strHTML = strHTML & "<p>权限不足,请登录网站管理！" & vbCrLf
        Else
            strHTML = strHTML & "<p>" & UTF2GB(iUsername) & "您好:" & vbCrLf
            sqlChannel = "select ChannelID,OrderID,ChannelName,ChannelShortName,ChannelDir,ModuleType,Disabled from PE_Channel where Disabled = " & PE_False & " order by OrderID"
            Set rsChannel = Conn.Execute(sqlChannel)
            Do While Not rsChannel.EOF
                If rsChannel("ModuleType") > 0 And rsChannel("ModuleType") < 6 Then
                    strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ChannelManage|" & rsChannel("ChannelID") & "|" & iUsername & "|" & iPassword & """>" & rsChannel("ChannelShortName") & "管理</a>" & vbCrLf
                End If
               rsChannel.MoveNext
            Loop
            rsChannel.Close
            Set rsChannel = Nothing
        End If
        strHTML = strHTML & "<br/>-----<br/><a href=""" & WapDomain & """>退出管理</a></p>" & vbCrLf
    End If
    strHTML = strHTML & "</card>" & vbCrLf
End Sub

Sub ChannelManage(ByVal iChannelID, ByVal iUsername, ByVal iPassword)
    Dim rsChannel, sqlChannel, rsArticle, sqlArticle, ModuleType
    strHTML = strHTML & "<card id=""main"" title=""后台管理"">" & vbCrLf
    If CheckAdmin(iUsername, iPassword) = False Then
        strHTML = strHTML & "<p>权限不足,请登录网站管理！" & vbCrLf
    Else
        sqlChannel = "select ChannelID,OrderID,ModuleType,Disabled from PE_Channel where ChannelID=" & iChannelID & " and Disabled = " & PE_False & " order by OrderID"
        Set rsChannel = Conn.Execute(sqlChannel)
        If rsChannel.BOF And rsChannel.EOF Then
            strHTML = strHTML & "<p>找不到频道！"
        Else
            ModuleType = rsChannel("ModuleType")
            strHTML = strHTML & "<p>待审列表："
            Select Case ModuleType
            Case 1
                sqlArticle = "select ArticleID,ChannelID,Title,Status from PE_Article Where ChannelID=" & iChannelID & " and Status=0 order by UpdateTime desc"
                Set rsArticle = Conn.Execute(sqlArticle)
                If Not (rsArticle.BOF And rsArticle.EOF) Then
                    Do While Not rsArticle.EOF
                        strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ArticlePass|" & iChannelID & "|" & rsArticle("ArticleID") & "|" & iUsername & "|" & iPassword & """>" & xml_nohtml(rsArticle("Title")) & "</a>"
                        rsArticle.MoveNext
                    Loop
                Else
                    strHTML = strHTML & "<br/>无未审核文章" & vbCrLf
                End If
            Case 2
                sqlArticle = "select SoftID,ChannelID,SoftName,Status from PE_Soft Where ChannelID=" & iChannelID & " and Status=0 order by UpdateTime desc"
                Set rsArticle = Conn.Execute(sqlArticle)
                If Not (rsArticle.BOF And rsArticle.EOF) Then
                    Do While Not rsArticle.EOF
                        strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=SoftPass|" & iChannelID & "|" & rsArticle("SoftID") & "|" & iUsername & "|" & iPassword & """>" & xml_nohtml(rsArticle("SoftName")) & "</a>"
                        rsArticle.MoveNext
                    Loop
                Else
                    strHTML = strHTML & "<br/>无未审核软件" & vbCrLf
                End If
            Case 3
                sqlArticle = "select PhotoID,ChannelID,PhotoName,Status from PE_Photo Where ChannelID=" & iChannelID & " and Status=0 order by UpdateTime desc"
                Set rsArticle = Conn.Execute(sqlArticle)
                If Not (rsArticle.BOF And rsArticle.EOF) Then
                    Do While Not rsArticle.EOF
                        strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=PhotoPass|" & iChannelID & "|" & rsArticle("PhotoID") & "|" & iUsername & "|" & iPassword & """>" & xml_nohtml(rsArticle("PhotoName")) & "</a>"
                        rsArticle.MoveNext
                    Loop
                Else
                    strHTML = strHTML & "<br/>无未审核图片" & vbCrLf
                End If
            Case 4
                sqlArticle = "select top 20 * from PE_GuestBook Where GuestIsPassed=" & PE_False & " order by GuestDatetime desc"
                Set rsArticle = Conn.Execute(sqlArticle)
                If Not (rsArticle.BOF And rsArticle.EOF) Then
                    Do While Not rsArticle.EOF
                        strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=GuestPass|" & iChannelID & "|" & rsArticle("GuestID") & "|" & iUsername & "|" & iPassword & """>" & xml_nohtml(rsArticle("GuestTitle")) & "</a>"
                        rsArticle.MoveNext
                    Loop
                Else
                    strHTML = strHTML & "<br/>无未审核留言" & vbCrLf
                End If
            Case 5
                sqlArticle = "select ProductID,ChannelID,ProductName,EnableSale from PE_Product Where ChannelID=" & iChannelID & " and EnableSale=" & PE_False & " order by UpdateTime desc"
                Set rsArticle = Conn.Execute(sqlArticle)
                If Not (rsArticle.BOF And rsArticle.EOF) Then
                    Do While Not rsArticle.EOF
                        strHTML = strHTML & "<br/><a href=""" & WapDomain & "?Source=ProductPass|" & iChannelID & "|" & rsArticle("ProductID") & "|" & iUsername & "|" & iPassword & """>" & xml_nohtml(rsArticle("ProductName")) & "</a>"
                        rsArticle.MoveNext
                    Loop
                Else
                    strHTML = strHTML & "<br/>无停止销售商品" & vbCrLf
                End If
            End Select
            rsArticle.Close
            Set rsArticle = Nothing
        End If
        rsChannel.Close
        Set rsChannel = Nothing
    End If
    strHTML = strHTML & "<br/>-----<br/><a href=""" & WapDomain & """>退出管理</a></p>" & vbCrLf
    strHTML = strHTML & "</card>" & vbCrLf
End Sub

Sub ArticlePass(ByVal iChannelID, ByVal iArticleID, ByVal iUsername, ByVal iPassword)
    If CheckAdmin(iUsername, iPassword) = False Then
        strHTML = strHTML & "<card id=""main"" title=""错误"">" & vbCrLf
        strHTML = strHTML & "<p>您不是本站超管,请登录网站管理！" & vbCrLf
        strHTML = strHTML & "<br/>-----<br/><a href=""" & WapDomain & """>退出管理</a></p>" & vbCrLf
        strHTML = strHTML & "</card>" & vbCrLf
    Else
        Conn.Execute ("update PE_Article set Status=3 where ArticleID=" & iArticleID & "")
        Call ChannelManage(iChannelID, iUsername, iPassword)
    End If
End Sub

Sub SoftPass(ByVal iChannelID, ByVal iArticleID, ByVal iUsername, ByVal iPassword)
    If CheckAdmin(iUsername, iPassword) = False Then
        strHTML = strHTML & "<card id=""main"" title=""错误"">" & vbCrLf
        strHTML = strHTML & "<p>您不是本站超管,请登录网站管理！" & vbCrLf
        strHTML = strHTML & "<br/>-----<br/><a href=""" & WapDomain & """>退出管理</a></p>" & vbCrLf
        strHTML = strHTML & "</card>" & vbCrLf
    Else
        Conn.Execute ("update PE_Soft set Status=3 where SoftID=" & iArticleID & "")
        Call ChannelManage(iChannelID, iUsername, iPassword)
    End If
End Sub

Sub PhotoPass(ByVal iChannelID, ByVal iArticleID, ByVal iUsername, ByVal iPassword)
    If CheckAdmin(iUsername, iPassword) = False Then
        strHTML = strHTML & "<card id=""main"" title=""错误"">" & vbCrLf
        strHTML = strHTML & "<p>您不是本站超管,请登录网站管理！" & vbCrLf
        strHTML = strHTML & "<br/>-----<br/><a href=""" & WapDomain & """>退出管理</a></p>" & vbCrLf
        strHTML = strHTML & "</card>" & vbCrLf
    Else
        Conn.Execute ("update PE_Photo set Status=3 where PhotoID=" & iArticleID & "")
        Call ChannelManage(iChannelID, iUsername, iPassword)
    End If
End Sub

Sub GuestPass(ByVal iChannelID, ByVal iArticleID, ByVal iUsername, ByVal iPassword)
    If CheckAdmin(iUsername, iPassword) = False Then
        strHTML = strHTML & "<card id=""main"" title=""错误"">" & vbCrLf
        strHTML = strHTML & "<p>您不是本站超管,请登录网站管理！" & vbCrLf
        strHTML = strHTML & "<br/>-----<br/><a href=""" & WapDomain & """>退出管理</a></p>" & vbCrLf
        strHTML = strHTML & "</card>" & vbCrLf
    Else
        Conn.Execute ("update PE_GuestBook set GuestIsPassed=" & PE_True & " where GuestID=" & iArticleID & "")
        Call ChannelManage(iChannelID, iUsername, iPassword)
    End If
End Sub

Sub ProductPass(ByVal iChannelID, ByVal iArticleID, ByVal iUsername, ByVal iPassword)
    If CheckAdmin(iUsername, iPassword) = False Then
        strHTML = strHTML & "<card id=""main"" title=""错误"">" & vbCrLf
        strHTML = strHTML & "<p>您不是本站超管,请登录网站管理！" & vbCrLf
        strHTML = strHTML & "<br/>-----<br/><a href=""" & WapDomain & """>退出管理</a></p>" & vbCrLf
        strHTML = strHTML & "</card>" & vbCrLf
    Else
        Conn.Execute ("update PE_Product set EnableSale=" & PE_True & " where ProductID=" & iArticleID & "")
        Call ChannelManage(iChannelID, iUsername, iPassword)
    End If
End Sub

'**************************************************
'通用函数部分开始
'**************************************************
'**************************************************
'函数名：GetSubStr2
'作  用：截字符串且不替换空格，汉字一个算两个字符，英文算一个字符
'**************************************************
Function GetSubStr2(ByVal str, ByVal strlen)
    If str = "" Then
        GetSubStr2 = ""
        Exit Function
    End If
    Dim l, t, c, i, strTemp
    str = Replace(Replace(Replace(Replace(str, "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
    l = Len(str)
    t = 0
    strTemp = str
    strlen = PE_CLng(strlen)
    For i = 1 To l
        c = Abs(Asc(Mid(str, i, 1)))
        If c > 255 Then
            t = t + 2
        Else
            t = t + 1
        End If
        If t >= strlen Then
            strTemp = Left(str, i)
            Exit For
        End If
    Next
    If strTemp <> str Then
        strTemp = strTemp & "…"
    End If
    GetSubStr2 = Replace(Replace(Replace(strTemp, Chr(34), "&quot;"), ">", "&gt;"), "<", "&lt;")
End Function

'**************************************************
'函数名：getpage
'作  用：根据字数分页函数(文章内容专用)
'**************************************************
Function getpage(ByVal iChannelID, ByVal iArticleID, ByVal str, ByVal PageNum)
    Dim StartWord, Length, PageAll, strlen, i, textmp
    StartWord = 1
    strlen = 160 '每页字数
    Length = Len(str) '要显示内容的总的长度
    PageAll = (Length + strlen - 1) \ strlen '该篇新闻总共可分的总页数
    i = PageNum '第几页的标记
    textmp = Mid(str, StartWord + i * strlen, strlen)
    
    If 0 <= i < PageAll Then
            textmp = textmp & "<br/>(" & i + 1 & "/" & PageAll & ")"
        If CInt(i) < CInt(PageAll) - 1 Then
            textmp = textmp & "<br/><a href=""" & WapDomain & "?Source=ShowArticle|" & iChannelID & "|" & iArticleID & "|" & i + 1 & """>下页</a>"
        End If
        If CInt(i) > 0 Then
            textmp = textmp & "<br/><a href=""" & WapDomain & "?Source=ShowArticle|" & iChannelID & "|" & iArticleID & "|" & i - 1 & """>上页</a>"
        End If
        textmp = textmp & "<br/><a href=""" & WapDomain & "?Source=ChannelList|" & iChannelID & """>返回列表</a>" & vbCrLf
    End If
    getpage = textmp
End Function

'**************************************************
'函数名：UTF2GB
'作  用：将UTF-8转换为GB2312码
'**************************************************
Function UTF2GB(ByVal UTFStr)
    Dim Dig, GBStr
    For Dig = 1 To Len(UTFStr)
        If Mid(UTFStr, Dig, 1) = "%" Then
            If Len(UTFStr) >= Dig + 8 Then
                GBStr = GBStr & ConvChinese(Mid(UTFStr, Dig, 9))
                Dig = Dig + 8
            Else
                GBStr = GBStr & Mid(UTFStr, Dig, 1)
            End If
        Else
            GBStr = GBStr & Mid(UTFStr, Dig, 1)
        End If
    Next
    UTF2GB = GBStr
End Function

Function ConvChinese(ByVal x)
    Dim a, i, j, DigS, unicode
    a = Split(Mid(x, 2), "%")
    i = 0
    j = 0
    For i = 0 To UBound(a)
        a(i) = c16to2(a(i))
    Next
    For i = 0 To UBound(a) - 1
    DigS = InStr(a(i), "0")
    unicode = ""
    For j = 1 To DigS - 1
        If j = 1 Then
            a(i) = Right(a(i), Len(a(i)) - DigS)
            unicode = unicode & a(i)
        Else
            i = i + 1
            a(i) = Right(a(i), Len(a(i)) - 2)
            unicode = unicode & a(i)
        End If
    Next
    If Len(c2to16(unicode)) = 4 Then
        ConvChinese = ConvChinese & ChrW(Int("&H" & c2to16(unicode)))
    Else
        ConvChinese = ConvChinese & Chr(Int("&H" & c2to16(unicode)))
    End If
    Next
End Function

Function c2to16(ByVal x)
    Dim i
    i = 1
    For i = 1 To Len(x) Step 4
        c2to16 = c2to16 & Hex(c2to10(Mid(x, i, 4)))
    Next
End Function

Function c2to10(ByVal x)
    Dim i
    c2to10 = 0
    If x = "0" Then Exit Function
    i = 0
    For i = 0 To Len(x) - 1
        If Mid(x, Len(x) - i, 1) = "1" Then c2to10 = c2to10 + 2 ^ (i)
    Next
End Function

Function c16to2(ByVal x)
    Dim i, tempstr
    i = 0
    For i = 1 To Len(Trim(x))
        tempstr = c10to2(CInt(Int("&h" & Mid(x, i, 1))))
        Do While Len(tempstr) < 4
            tempstr = "0" & tempstr
        Loop
        c16to2 = c16to2 & tempstr
    Next
End Function

Function c10to2(ByVal x)
    Dim mysign, DigS, tempnum, i
    mysign = Sgn(x)
    x = Abs(x)
    DigS = 1
    Do
        If x < 2 ^ DigS Then
            Exit Do
        Else
            DigS = DigS + 1
        End If
    Loop
    tempnum = x
    i = 0
    For i = DigS To 1 Step -1
        If tempnum >= 2 ^ (i - 1) Then
            tempnum = tempnum - 2 ^ (i - 1)
            c10to2 = c10to2 & "1"
        Else
            c10to2 = c10to2 & "0"
        End If
    Next
    If mysign = -1 Then c10to2 = "-" & c10to2
End Function

'**************************************************
'函数名：CheckAdmin
'作  用：验证管理员身份
'**************************************************
Function CheckAdmin(ByVal iName, ByVal iWord)
    Dim rsUser, sqlUser
    CheckAdmin = False
    sqlUser = "select * from PE_Admin Where UserName='" & UTF2GB(iName) & "' and Password='" & MD5(iWord, 16) & "' and Purview=1"
    Set rsUser = Conn.Execute(sqlUser)
    If rsUser.BOF And rsUser.EOF Then
        CheckAdmin = False
    Else
        CheckAdmin = True
    End If
    rsUser.Close
    Set rsUser = Nothing
End Function

'**************************************************
'函数名：ReplaceText
'作  用：过滤非法字符串
'参  数：iText-----输入字符串
'返回值：替换后字符串
'**************************************************
Function ReplaceText(iText, iType)
    Dim rText, rsKey, sqlKey, i, Keyrow, Keycol
    If PE_Cache.GetValue("Site_ReplaceText") = "" Then
        Set rsKey = Server.CreateObject("Adodb.RecordSet")
        sqlKey = "Select Source,ReplaceText,OpenType,ReplaceType,Priority from PE_KeyLink where isUse=1 and LinkType=1 order by Priority"
        rsKey.Open sqlKey, Conn, 1, 1
        If Not (rsKey.BOF And rsKey.EOF) Then
            PE_Cache.SetValue "Site_ReplaceText", rsKey.GetString(, , "|||", "@@@", "")
            rsKey.Close
            Set rsKey = Nothing
        Else
            rsKey.Close
            Set rsKey = Nothing
            ReplaceText = iText
            Exit Function
        End If
    End If
    rText = iText
    Keyrow = Split(PE_Cache.GetValue("Site_ReplaceText"), "@@@")
    For i = 0 To UBound(Keyrow) - 1
        Keycol = Split(Keyrow(i), "|||")
        If Int(Keycol(3)) = 0 Or Int(Keycol(3)) = iType Then rText = PE_Replace(rText, Keycol(0), Keycol(1))
    Next
    ReplaceText = rText
End Function


Function GetDownloadUrlList(DownloadUrls)
    Dim arrDownloadUrls, arrUrls, iTemp, strUrls
    Dim rsDownServer, sqlDownServer, ShowServerName, iShowModule
    If DownloadUrls = "" Then
        GetDownloadUrlList = ""
        Exit Function
    End If
    strUrls = ""
    If InStr(DownloadUrls, "@@@") > 0 Then
    '构建镜像服务器下载地址列表
        arrDownloadUrls = Trim(Replace(DownloadUrls, "@@@", "")) '得PE_Soft中的Url地址
        sqlDownServer = "select * from PE_DownServer"
        Set rsDownServer = Server.CreateObject("adodb.recordset")
        rsDownServer.Open sqlDownServer, Conn, 1, 3
        If rsDownServer.BOF Or rsDownServer.EOF Then
           strUrls = "对不起，未找到任何镜像服务器信息！"
        End If

        Do While Not rsDownServer.EOF
            If rsDownServer("ShowType") = 0 Then
               ShowServerName = rsDownServer("ServerName")
            Else
               ShowServerName = "<img src=""" & rsDownServer("ServerLogo") & """ border=""0"" />"
            End If
            '设置了下载点数的处理（PE_DownServer中有收费相关字段）
            If rsDownServer("InfoPoint") = 0 Then
                strUrls = strUrls & "<a href=""" & rsDownServer("ServerUrl") & arrDownloadUrls & """>" & ShowServerName & "</a><br/>"
            End If
            rsDownServer.MoveNext
        Loop
        GetDownloadUrlList = strUrls
        rsDownServer.Close
        Set rsDownServer = Nothing
    Else
        arrDownloadUrls = Split(DownloadUrls, "$$$")
        For iTemp = 0 To UBound(arrDownloadUrls)
            arrUrls = Split(arrDownloadUrls(iTemp), "|")
            If UBound(arrUrls) >= 1 Then
                If arrUrls(1) <> "" And arrUrls(1) <> "http://" Then
                    If Left(arrUrls(1), 1) <> "/" And InStr(arrUrls(1), "://") <= 0 Then
                        strUrls = strUrls & "<a href=""" & SiteUrl & "/" & ChannelDir & "/" & UploadDir & "/" & arrUrls(1) & """>" & arrUrls(0) & "</a><br/>"
                    Else
                        strUrls = strUrls & "<a href=""" & GetFirstSeparatorToEnd(arrDownloadUrls(iTemp), "|") & """>" & arrUrls(0) & "</a><br/>"
                    End If
                End If
            End If
        Next
        GetDownloadUrlList = strUrls
    End If
End Function

%>
