<!--#include file="Start.asp"-->
<!--#include file="Include/PowerEasy.Cache.asp"-->
<!--#include file="Include/PowerEasy.Common.Front.asp"-->
<!--#include file="Include/PowerEasy.Common.Content.asp"-->
<!--#include file="Include/PowerEasy.Channel.asp"-->
<!--#include file="Include/PowerEasy.ArticleList.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

PageTitle = "作者信息"

Dim strList, aValue,PriClassID, ClassField(5)
Dim AuthorName
Dim rsAuthor, sqlAuthor
Dim TimeData
ChannelID = PE_CLng(Trim(Request("ChannelID")))
AuthorName = Trim(Request("AuthorName"))
If AuthorName = "" Then
    Call WriteErrMsg("<li>请指定要查看的作者姓名！</li>", ComeUrl)
    Response.End
End If
AuthorName = ReplaceBadChar(AuthorName)
strFileName = "ShowAuthor.asp?AuthorName=" & AuthorName
strFileName = strFileName & "&ChannelID=" & ChannelID
TimeData = Trim(Request("Data"))
If TimeData = "" Or Not (IsDate(TimeData)) Then
    TimeData = 0
Else
    TimeData = FormatDateTime(TimeData, 2)
    strFileName = strFileName & "&Data=" & TimeData
End If
MaxPerPage = 20
strNavPath = strNavPath & strNavLink & "&nbsp;" & PageTitle

'取频道参数
Call GetChannel(ChannelID)

sqlAuthor = "select * from PE_Author where AuthorName='" & AuthorName & "' and (ChannelID=0 or ChannelID=" & ChannelID & ") and Passed=" & PE_True
Set rsAuthor = Server.CreateObject("ADODB.Recordset")
rsAuthor.Open sqlAuthor, Conn, 1, 1
If rsAuthor.BOF And rsAuthor.EOF Then
    rsAuthor.Close
    Set rsAuthor = Nothing
    Call WriteErrMsg("<li>找不到指定的作者！</li>", ComeUrl)
    Response.End
End If

Dim iArrTemp

If rsAuthor("TemplateID") < 1 Then
    strHtml = GetTemplate(0, 10, 0)
Else
    strHtml = GetTemplate(0, 10, rsAuthor("TemplateID"))
End If
strHtml = Replace(strHtml, "{$AuthorID}", rsAuthor("ID"))
strHtml = Replace(strHtml, "{$AuthorName}", AuthorName)
strHtml = Replace(strHtml, "{$ChannelID}", ChannelID)
Call ReplaceCommonLabel

strHtml = Replace(strHtml, "{$PageTitle}", SiteTitle & " >> " & PageTitle)
strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)
strHtml = Replace(strHtml, "{$MenuJS}", GetMenuJS("", False))
strHtml = Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))
If ChannelID > 0 Then
    strHtml = Replace(strHtml, "{$Rss}", "<a href='" & strInstallDir & "rss.asp?ChannelID=" & ChannelID & "&AuthorName=" & AuthorName & "' Target='_blank'><img src='" & strInstallDir & "images/rss.gif' border=0></a>")
    strHtml = Replace(strHtml, "{$ShowList}", "AuthorList.asp?ChannelID=" & ChannelID)
Else
    strHtml = Replace(strHtml, "{$Rss}", "")
    strHtml = Replace(strHtml, "{$ShowList}", "AuthorList.asp")
End If

regEx.Pattern = "\{\$AuthorPhoto\((.*?)\)\}"
Set Matches = regEx.Execute(strHtml)
For Each Match In Matches
    aValue = Replace(Match.SubMatches(0), Chr(34), " ")
    iArrTemp = Split(aValue, ",")
    If UBound(iArrTemp) < 1 Then
        strHtml = "函数式标签：{$AuthorPhoto()}的参数个数不对。请检查模板中的此标签。"
    Else
        If rsAuthor("Photo") = "" Or IsNull(rsAuthor("Photo")) Then
            strHtml = Replace(strHtml, Match.value, "<img src='AuthorPic/default.gif' width='" & PE_CLng(iArrTemp(0)) & "' height='" & PE_CLng(iArrTemp(1)) & "'>")
        Else
            strHtml = Replace(strHtml, Match.value, "<img src='" & rsAuthor("Photo") & "' width='" & PE_CLng(iArrTemp(0)) & "' height='" & PE_CLng(iArrTemp(1)) & "'>")
        End If
    End If
Next

If rsAuthor("Sex") = 1 Then
    strHtml = Replace(strHtml, "{$AuthorSex}", strMan)
Else
    strHtml = Replace(strHtml, "{$AuthorSex}", strGirl)
End If
strHtml = Replace(strHtml, "{$AuthorAddTime}", Year(rsAuthor("LastUseTime")) & strYear & Month(rsAuthor("LastUseTime")) & strMonth & Day(rsAuthor("LastUseTime")) & strDay)
strHtml = Replace(strHtml, "{$AuthorBirthDay}", Year(rsAuthor("BirthDay")) & strYear & Month(rsAuthor("BirthDay")) & strMonth & Day(rsAuthor("BirthDay")) & strDay)
strHtml = Replace(strHtml, "{$AuthorCompany}", ReplaceSpace(rsAuthor("Company")))
strHtml = Replace(strHtml, "{$AuthorDepartment}", ReplaceSpace(rsAuthor("Department")))
strHtml = Replace(strHtml, "{$AuthorAddress}", ReplaceSpace(rsAuthor("Address")))
strHtml = Replace(strHtml, "{$AuthorTel}", ReplaceSpace(rsAuthor("Tel")))
strHtml = Replace(strHtml, "{$AuthorFax}", ReplaceSpace(rsAuthor("Fax")))
strHtml = Replace(strHtml, "{$AuthorZipCode}", ReplaceSpace(rsAuthor("ZipCode")))
strHtml = Replace(strHtml, "{$AuthorHomePage}", ReplaceSpace(rsAuthor("HomePage")))
strHtml = Replace(strHtml, "{$AuthorEmail}", ReplaceSpace(rsAuthor("Email")))
strHtml = Replace(strHtml, "{$AuthorQQ}", ReplaceSpace(rsAuthor("QQ")))
If rsAuthor("AuthorType") = 1 Then
    strHtml = Replace(strHtml, "{$AuthorType}", XmlText("ShowSource", "ShowAuthor/AuthorType1", "大陆作者"))
ElseIf rsAuthor("AuthorType") = 2 Then
    strHtml = Replace(strHtml, "{$AuthorType}", XmlText("ShowSource", "ShowAuthor/AuthorType2", "港台作者"))
ElseIf rsAuthor("AuthorType") = 3 Then
    strHtml = Replace(strHtml, "{$AuthorType}", XmlText("ShowSource", "ShowAuthor/AuthorType3", "海外作者"))
ElseIf rsAuthor("AuthorType") = 4 Then
    strHtml = Replace(strHtml, "{$AuthorType}", XmlText("ShowSource", "ShowAuthor/AuthorType4", "本站特约"))
Else
    strHtml = Replace(strHtml, "{$AuthorType}", XmlText("ShowSource", "ShowAuthor/AuthorType5", "其他作者"))
End If
strHtml = PE_Replace(strHtml, "{$AuthorIntro}", rsAuthor("Intro"))

regEx.Pattern = "\{\$AuthorArticleList\((.*?)\)\}"
Set Matches = regEx.Execute(strHtml)
For Each Match In Matches
    aValue = Replace(Replace(Replace(Match.SubMatches(0), Chr(34), ""), "{$AuthorArticleList(", ""), ")}", "")
    strList = ShowArticleList(AuthorName & "," & ChannelID & "," & ModuleName & "," & ChannelDir & "," & UploadDir & "," & aValue, 1)
    strHtml = Replace(strHtml, Match.value, strList)
Next
If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
rsAuthor.Close
Set rsAuthor = Nothing
Response.Write strHtml
Call CloseConn
%>
