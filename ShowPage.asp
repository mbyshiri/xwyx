<!--#include file="Start.asp"-->
<!--#include file="Include/PowerEasy.Cache.asp"-->
<!--#include file="Include/PowerEasy.Channel.asp"-->
<!--#include file="Include/PowerEasy.Class.asp"-->
<!--#include file="Include/PowerEasy.Special.asp"-->
<!--#include file="Include/PowerEasy.Article.asp"-->
<!--#include file="Include/PowerEasy.Soft.asp"-->
<!--#include file="Include/PowerEasy.Photo.asp"-->
<!--#include file="Include/PowerEasy.Product.asp"-->
<!--#include file="Include/PowerEasy.Common.Front.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

ChannelID = 0
Dim PageID, sqlPage, rsPage, tPageContent, PE_Content

PageID = PE_CLng(Trim(Request("id")))
If PageID = 0 Then
    Response.Write "参数丢失！"
    Call CloseConn
    Response.End
End If

sqlPage = "select PageName,PageIntro,PageUrl,PageFileName,PageContent,arrGroupID from PE_Page where ID=" & PageID
Set rsPage = Server.CreateObject("ADODB.Recordset")
rsPage.Open sqlPage, Conn, 1, 1
If rsPage.BOF And rsPage.EOF Then
    Response.Write "找不到指定的页面！"
    rsPage.Close
    Set rsPage = Nothing
    Call CloseConn
    Response.End
End If

strHtml = rsPage("PageContent")

'加入对权限判断的处理
If Trim(rsPage("arrGroupID") & "") <> "" Then
    UserLogined = CheckUserLogined()
    If UserLogined = True Then
        UserLogined = FoundInArr(rsPage("arrGroupID"), GroupID, ",")
    End If
    If UserLogined <> True Then
        rsPage.Close
        Set rsPage = Nothing
        Call CloseConn
        Response.Write "您尚未被授权访问此页面！"
        Response.End
    End If
End If

'加入对输入参数的处理
Dim inputarr, inputarr2, i, inputerr
inputerr = False
inputarr = Split(rsPage("PageIntro"), vbCrLf)
For i = 0 To UBound(inputarr)
    If inputarr(i) <> "" Then
       inputarr2 = Split(inputarr(i), "|")
       If UBound(inputarr2) = 3 Then
          If LCase(inputarr2(2)) = "false" Then '判断是否为必须参数
              If inputarr2(1) = "0" Then
                  If Trim(Request(inputarr2(0))) = "" And inputarr2(3) <> "" Then
                      strHtml = Replace(strHtml, "{$pageinput(" & i & ")}", PE_CLng(inputarr2(3)))
                  Else
                      strHtml = Replace(strHtml, "{$pageinput(" & i & ")}", PE_CLng(Trim(Request(inputarr2(0)))))
                  End If
              Else
                  If Trim(Request(inputarr2(0))) = "" And inputarr2(3) <> "" Then
                      strHtml = Replace(strHtml, "{$pageinput(" & i & ")}", ReplaceBadChar(inputarr2(3)))
                  Else
                      strHtml = Replace(strHtml, "{$pageinput(" & i & ")}", ReplaceBadChar(Trim(Request(inputarr2(0)))))
                  End If
              End If
          Else
              If Trim(Request(inputarr2(0))) = "" Then
                  strHtml = Replace(XmlText("Site", "PrivatePage/FieldErr", "参数{$FieldName}不能为空!"), "{$FieldName}", inputarr2(0))
                  inputerr = True
              Else
                  If inputarr2(1) = "0" Then
                      strHtml = Replace(strHtml, "{$pageinput(" & i & ")}", PE_CLng(Trim(Request(inputarr2(0)))))
                  Else
                      strHtml = Replace(strHtml, "{$pageinput(" & i & ")}", ReplaceBadChar(Trim(Request(inputarr2(0)))))
                  End If
              End If
          End If
       End If
    End If
Next

If inputerr = True Then
    rsPage.Close
    Set rsPage = Nothing
    Call CloseConn
    Response.Write strHtml
    Response.End
End If

Call ReplaceCommonLabel
strHtml = Replace(strHtml, "{$ShowPath}", XmlText("BaseText", "Nav", "您现在的位置：") & "&nbsp;<a class='LinkPath' href='" & SiteUrl & "'>" & SiteName & "</a>"& strNavLink & "&nbsp;" &rsPage("PageName"))
strHtml = Replace(strHtml, "{$PageTitle}", rsPage("PageName"))
strHtml = Replace(strHtml, "{$MenuJS}", GetMenuJS("", False))
strHtml = Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))

Set PE_Content = New Article
PE_Content.Init
strHtml = PE_Content.GetCustomFromTemplate(strHtml)
strHtml = PE_Content.GetPicFromTemplate(strHtml)
strHtml = PE_Content.GetListFromTemplate(strHtml)
strHtml = PE_Content.GetSlidePicFromTemplate(strHtml)
Set PE_Content = Nothing

Set PE_Content = New Soft
PE_Content.Init
strHtml = PE_Content.GetCustomFromTemplate(strHtml)
strHtml = PE_Content.GetPicFromTemplate(strHtml)
strHtml = PE_Content.GetListFromTemplate(strHtml)
strHtml = PE_Content.GetSlidePicFromTemplate(strHtml)
Set PE_Content = Nothing

Set PE_Content = New Photo
PE_Content.Init
strHtml = PE_Content.GetCustomFromTemplate(strHtml)
strHtml = PE_Content.GetPicFromTemplate(strHtml)
strHtml = PE_Content.GetListFromTemplate(strHtml)
strHtml = PE_Content.GetSlidePicFromTemplate(strHtml)
Set PE_Content = Nothing

Set PE_Content = New Product
PE_Content.Init
strHtml = PE_Content.GetCustomFromTemplate(strHtml)
strHtml = PE_Content.GetPicFromTemplate(strHtml)
strHtml = PE_Content.GetListFromTemplate(strHtml)
strHtml = PE_Content.GetSlidePicFromTemplate(strHtml)
Set PE_Content = Nothing
rsPage.Close
Set rsPage = Nothing
Response.Write strHtml
Call CloseConn
%>
