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

PageTitle = "来源信息"
Dim strtmp, strList, aValue, PriClassID, ClassField(5)
Dim SourceName, i
Dim rsCopyFrom, sqlCopyFrom, rsArticle, sqlArticle
Dim TimeData
ChannelID = PE_CLng(Trim(Request("ChannelID")))

strFileName = "ShowCopyfrom.asp?ChannelID=" & ChannelID
MaxPerPage = 20
strNavPath = strNavPath & strNavLink & "&nbsp;" & PageTitle

SourceName = Trim(Request("SourceName"))
If SourceName = "" Then
    Call WriteErrMsg("<li>请指定要查看的来源名称！</li>", ComeUrl)
    Response.End
Else
    SourceName = ReplaceBadChar(SourceName)
    strFileName = strFileName & "&SourceName=" & SourceName
End If
TimeData = Trim(Request("Data"))
If TimeData = "" Or Not (IsDate(TimeData)) Then
    TimeData = "0"
Else
    TimeData = FormatDateTime(TimeData, 2)
    strFileName = strFileName & "&Data=" & TimeData
End If

'取频道参数
Call GetChannel(ChannelID)

sqlCopyFrom = "select * from PE_CopyFrom where SourceName='" & SourceName & "' and (ChannelID=0 or ChannelID=" & ChannelID & ") and Passed=" & PE_True
Set rsCopyFrom = Server.CreateObject("ADODB.Recordset")
rsCopyFrom.Open sqlCopyFrom, Conn, 1, 1
If rsCopyFrom.BOF And rsCopyFrom.EOF Then
    rsCopyFrom.Close
    Set rsCopyFrom = Nothing
    Call WriteErrMsg("<li>找不到指定的来源！</li>", ComeUrl)
    Response.End
End If
strHtml = GetTemplate(0, 12, 0)
strHtml = Replace(strHtml, "{$ChannelID}", ChannelID)
strHtml = Replace(strHtml, "{$ShowName}", SourceName)
Call ReplaceCommonLabel

strHtml = Replace(strHtml, "{$PageTitle}", SiteTitle & " >> " & PageTitle)
strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)
strHtml = Replace(strHtml, "{$MenuJS}", GetMenuJS("", False))
strHtml = Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))
If ChannelID > 0 Then
        strHtml = Replace(strHtml, "{$ShowList}", "ShowCopyForm.asp?Action=List&ChannelID=" & ChannelID)
Else
        strHtml = Replace(strHtml, "{$ShowList}", "ShowCopyForm.asp?Action=List")
End If

If rsCopyFrom("Photo") = "" Or IsNull(rsCopyFrom("Photo")) Then
    strHtml = Replace(strHtml, "{$ShowPhoto}", "<img src='CopyFromPic/default.gif' width='150' height='175'>")
Else
    strHtml = Replace(strHtml, "{$ShowPhoto}", "<img src='" & rsCopyFrom("Photo") & "' width='150' height='175'>")
End If
strHtml = Replace(strHtml, "{$ShowContacterName}", ReplaceSpace(rsCopyFrom("ContacterName")))
strHtml = Replace(strHtml, "{$ShowAddress}", ReplaceSpace(rsCopyFrom("Address")))
strHtml = Replace(strHtml, "{$ShowTel}", ReplaceSpace(rsCopyFrom("Tel")))
strHtml = Replace(strHtml, "{$ShowFax}", ReplaceSpace(rsCopyFrom("Fax")))
strHtml = Replace(strHtml, "{$ShowZipCode}", ReplaceSpace(rsCopyFrom("ZipCode")))
strHtml = Replace(strHtml, "{$ShowMail}", ReplaceSpace(rsCopyFrom("Mail")))
strHtml = Replace(strHtml, "{$ShowHomePage}", ReplaceSpace(rsCopyFrom("HomePage")))
strHtml = Replace(strHtml, "{$ShowEmail}", ReplaceSpace(rsCopyFrom("Email")))
strHtml = Replace(strHtml, "{$ShowQQ}", ReplaceSpace(rsCopyFrom("QQ")))
Select Case rsCopyFrom("SourceType")
Case 1
    strHtml = Replace(strHtml, "{$ShowType}", XmlText("ShowSource", "ShowCopyFrom/CopyFromType1", "友情站点"))
Case 2
    strHtml = Replace(strHtml, "{$ShowType}", XmlText("ShowSource", "ShowCopyFrom/CopyFromType2", "中文站点"))
Case 3
    strHtml = Replace(strHtml, "{$ShowType}", XmlText("ShowSource", "ShowCopyFrom/CopyFromType3", "外文站点"))
Case Else
    strHtml = Replace(strHtml, "{$ShowType}", XmlText("ShowSource", "ShowCopyFrom/CopyFromType4", "其他站点"))
End Select
strHtml = PE_Replace(strHtml, "{$ShowIntro}", rsCopyFrom("Intro"))

regEx.Pattern = "\{\$AuthorArticleList\((.*?)\)\}"
Set Matches = regEx.Execute(strHtml)
For Each Match In Matches
    aValue = Replace(Replace(Replace(Match.SubMatches(0), Chr(34), ""), "{$AuthorArticleList(", ""), ")}", "")
    strList = ShowArticleList(SourceName & "," & ChannelID & "," & ModuleName & "," & ChannelDir & "," & UploadDir & "," & aValue, 2)
    strHtml = Replace(strHtml, Match.value, strList)
Next
If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, ChannelItemUnit & ChannelShortName, False))
rsCopyFrom.Close
Set rsCopyFrom = Nothing
Response.Write strHtml
Call CloseConn

%>
