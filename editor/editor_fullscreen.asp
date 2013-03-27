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


Dim ChannelID, TrueSiteUrl
ChannelID = PE_CLng(Trim(Request("ChannelID")))

TrueSiteUrl = Trim(Request.ServerVariables("HTTP_HOST"))


Response.Write "<HTML>" & vbCrLf
Response.Write "<HEAD>" & vbCrLf
Response.Write "<TITLE>HtmlEdit - 全屏编辑</TITLE>" & vbCrLf
Response.Write "<META http-equiv=Content-Type content=""text/html; charset=gb2312"">" & vbCrLf
Response.Write "</HEAD>" & vbCrLf
Response.Write "<body leftmargin=0 topmargin=0 onunload=""Minimize()"">" & vbCrLf
Response.Write "<input type=""hidden"" id=""ContentFullScreen"" name=""ContentFullScreen"" value="""">" & vbCrLf
Response.Write "<script language=VBScript>" & vbCrLf
Response.Write "   Dim Matches, Match, arrContent, Content1, Content2,Content3,Content5" & vbCrLf
Response.Write "   Dim strTemp, strTemp2, StrBody,TemplateContent" & vbCrLf
Response.Write "   Set regEx = New RegExp" & vbCrLf

If Request.QueryString("num") = 1 Then
	Response.Write "ContentFullScreen.value=opener.editor.HtmlEdit.document.body.innerHTML" & vbCrLf
	Response.Write "TemplateContent= opener.document.form1.Content.value" & vbCrLf
Else
	Response.Write "ContentFullScreen.value =opener.editor2.HtmlEdit.document.body.innerHTML" & vbCrLf
	Response.Write "TemplateContent= opener.document.form1.Content2.value" & vbCrLf
End If

Response.Write "   ContentFullScreen.value =""<html><head><META http-equiv=Content-Type content=text/html; charset=gb2312><link href='" & InstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'></head><body leftmargin=0 topmargin=0 >"" & ContentFullScreen.value" & vbCrLf
Response.Write "   document.Write ""<iframe ID='EditorFullScreen' src='../editor.asp?ChannelID=" & ChannelID & "&ShowType=1&TemplateType=3&tContentid=ContentFullScreen' frameborder='0' scrolling=no width='100%' HEIGHT='100%'></iframe>""" & vbCrLf

Response.Write "Function  Resumeblank(byval Content)" & vbCrLf
Response.Write " Dim strHtml,strHtml2,Num,Numtemp,Strtemp" & vbCrLf
Response.Write "   strHtml=Replace(Content, ""<DIV"", ""<div"")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""</DIV>"", ""</div>"")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""<TABLE"", ""<table"")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""</TABLE>"", vbCrLf & ""</table>""& vbCrLf)" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""<TBODY>"", """")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""</TBODY>"","""" & vbCrLf)" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""<TR"", ""<tr"")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""</TR>"", vbCrLf & ""</tr>""& vbCrLf)" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""<TD"", ""<td"")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""</TD>"", ""</td>"")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""<!--"", vbCrLf & ""<!--"")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""<SELECT"",vbCrLf & ""<Select"")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""</SELECT>"",vbCrLf & ""</Select>"")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""<OPTION"",vbCrLf & ""  <Option"")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""</OPTION>"",""</Option>"")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""<INPUT"",vbCrLf & ""  <Input"")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""<script"",vbCrLf & ""<script"")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""&amp;"",""&"")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""{$--"",vbCrLf & ""<!--$"")" & vbCrLf
Response.Write "   strHtml=Replace(strHtml, ""--}"",""$-->"")" & vbCrLf
Response.Write "   arrContent = Split(strHtml,vbCrLf)" & vbCrLf
Response.Write "    For i = 0 To UBound(arrContent)" & vbCrLf
Response.Write "        Numtemp=false" & vbCrLf
Response.Write "        if Instr(arrContent(i),""<table"")>0 then" & vbCrLf
Response.Write "            Numtemp=True" & vbCrLf
Response.Write "            if Strtemp<>""<table"" and Strtemp <>""</table>"" then" & vbCrLf
Response.Write "              Num=Num+2" & vbCrLf
Response.Write "            End if " & vbCrLf
Response.Write "            Strtemp=""<table""" & vbCrLf
Response.Write "        elseif Instr(arrContent(i),""<tr"")>0 then" & vbCrLf
Response.Write "            Numtemp=True" & vbCrLf
Response.Write "            if Strtemp<>""<tr"" and Strtemp<>""</tr>"" then" & vbCrLf
Response.Write "              Num=Num+2" & vbCrLf
Response.Write "            End if " & vbCrLf
Response.Write "            Strtemp=""<tr""" & vbCrLf
Response.Write "        elseif Instr(arrContent(i),""<td"")>0 then" & vbCrLf
Response.Write "            Numtemp=True" & vbCrLf
Response.Write "            if Strtemp<>""<td"" and Strtemp<>""</td>"" then" & vbCrLf
Response.Write "              Num=Num+2" & vbCrLf
Response.Write "            End if " & vbCrLf
Response.Write "            Strtemp=""<td""" & vbCrLf
Response.Write "        elseif Instr(arrContent(i),""</table>"")>0 then" & vbCrLf
Response.Write "            Numtemp=True" & vbCrLf
Response.Write "            if Strtemp<>""</table>"" and Strtemp<>""<table"" then" & vbCrLf
Response.Write "              Num=Num-2" & vbCrLf
Response.Write "            End if " & vbCrLf
Response.Write "            Strtemp=""</table>""" & vbCrLf
Response.Write "        elseif Instr(arrContent(i),""</tr>"")>0 then" & vbCrLf
Response.Write "            Numtemp=True" & vbCrLf
Response.Write "            if Strtemp<>""</tr>"" and Strtemp<>""<tr"" then" & vbCrLf
Response.Write "              Num=Num-2" & vbCrLf
Response.Write "            End if " & vbCrLf
Response.Write "            Strtemp=""</tr>""" & vbCrLf
Response.Write "        elseif Instr(arrContent(i),""</td>"")>0 then" & vbCrLf
Response.Write "            Numtemp=True" & vbCrLf
Response.Write "            if Strtemp<>""</td>"" and Strtemp<>""<td"" then" & vbCrLf
Response.Write "              Num=Num-2" & vbCrLf
Response.Write "            End if " & vbCrLf
Response.Write "            Strtemp=""</td>""" & vbCrLf
Response.Write "        elseif Instr(arrContent(i),""<!--"")>0 then" & vbCrLf
Response.Write "            Numtemp=True" & vbCrLf
Response.Write "        End if" & vbCrLf
Response.Write "        if Num< 0 then Num = 0" & vbCrLf
Response.Write "        if trim(arrContent(i))<>"""" then" & vbCrLf
Response.Write "            if i=0 then" & vbCrLf
Response.Write "                strHtml2= string(Num,"" "") & arrContent(i) " & vbCrLf
Response.Write "            elseif Numtemp=True then" & vbCrLf
Response.Write "                strHtml2= strHtml2 & vbCrLf & string(Num,"" "") & arrContent(i) " & vbCrLf
Response.Write "            else" & vbCrLf
Response.Write "                strHtml2= strHtml2 & vbCrLf & arrContent(i) " & vbCrLf
Response.Write "            end if" & vbCrLf
Response.Write "        end if" & vbCrLf
Response.Write "    Next" & vbCrLf
Response.Write " Resumeblank=strHtml2" & vbCrLf
Response.Write "End function" & vbCrLf

Response.Write "Function Minimize()" & vbCrLf
Response.Write "       regEx.IgnoreCase = True" & vbCrLf
Response.Write "       regEx.Global = True" & vbCrLf
Response.Write "       regEx.Pattern = ""(\<body)(.[^\<]*)(\>)""" & vbCrLf
Response.Write "       Set Matches = regEx.Execute(TemplateContent)" & vbCrLf
Response.Write "         For Each Match In Matches" & vbCrLf
Response.Write "            StrBody = Match.Value" & vbCrLf
Response.Write "         Next" & vbCrLf
Response.Write "         arrContent = Split(TemplateContent, StrBody)" & vbCrLf
Response.Write "         Content1 = arrContent(0) & StrBody" & vbCrLf
Response.Write "         Content2 = arrContent(1)" & vbCrLf
Response.Write "         Content5 = EditorFullScreen.HtmlEdit.document.Body.innerHTML" & vbCrLf
Response.Write "         regEx.Pattern = ""\<IMG(.[^\<]*?)\}['|""""]\>""" & vbCrLf
Response.Write "         Set Matches = regEx.Execute(Content5)" & vbCrLf
Response.Write "         For Each Match In Matches" & vbCrLf
Response.Write "             regEx.Pattern = ""\{\$(.*)\}""" & vbCrLf
Response.Write "             Set strTemp = regEx.Execute(Match.Value)" & vbCrLf
Response.Write "             For Each Match2 In strTemp" & vbCrLf
Response.Write "                strTemp2 = Replace(Match2.Value, ""?"", """""""")" & vbCrLf
Response.Write "                Content5 = Replace(Content5, Match.Value, ""<!--""&strTemp2&""-->"")" & vbCrLf
Response.Write "             Next" & vbCrLf
Response.Write "         Next" & vbCrLf
Response.Write "         regEx.Pattern = ""\<IMG(.[^\<]*)\$\>""" & vbCrLf
Response.Write "         Set Matches = regEx.Execute(Content5)" & vbCrLf
Response.Write "         For Each Match In Matches" & vbCrLf
Response.Write "         regEx.Pattern = ""\#(.*)\#""" & vbCrLf
Response.Write "         Set strTemp = regEx.Execute(Match.Value)" & vbCrLf
Response.Write "            For Each Match2 In strTemp" & vbCrLf
Response.Write "               strTemp2 = Replace(strTemp2, ""#"", """")" & vbCrLf
Response.Write "               strTemp2 = Replace(strTemp2, ""[!"", ""<"")" & vbCrLf
Response.Write "               strTemp2 = Replace(strTemp2, ""!]"", "">"")" & vbCrLf
Response.Write "               Content5 = Replace(Content5, Match.Value, strTemp2)" & vbCrLf
Response.Write "            Next" & vbCrLf
Response.Write "         Next" & vbCrLf
Response.Write "        Content5=Replace(Content5, ""http://" & TrueSiteUrl & InstallDir & """,""{$InstallDir}"")" & vbCrLf
Response.Write "        Content5=Replace(Content5, ""http://" & LCase(TrueSiteUrl) & LCase(InstallDir) & """,""{$InstallDir}"")" & vbCrLf

If Request.QueryString("num") = 1 Then
	Response.Write "opener.editor.HtmlEdit.document.body.innerHTML=Resumeblank(EditorFullScreen.getHTML())" & vbCrLf
	Response.Write "opener.document.form1.Content.value=Content1& vbCrLf & Resumeblank(Content5) & vbCrLf &""</body>""& vbCrLf &""</html>""" & vbCrLf
	Response.Write "opener.editor.showBorders()" & vbCrLf
	Response.Write "opener.editor.showBorders()" & vbCrLf
Else
	Response.Write "opener.editor2.HtmlEdit.document.body.innerHTML=Resumeblank(EditorFullScreen.getHTML())" & vbCrLf
	Response.Write "opener.document.form1.Content2.value=Content1& vbCrLf & Resumeblank(Content5) & vbCrLf &""</body>""& vbCrLf &""</html>""" & vbCrLf
	Response.Write "opener.editor2.showBorders()" & vbCrLf
	Response.Write "opener.editor2.showBorders()" & vbCrLf
End If

Response.Write "    Set regEx = Nothing" & vbCrLf
Response.Write "End function" & vbCrLf
Response.Write "function setstatus()" & vbCrLf '这两个为了兼容editor.asp多用途临时作用
Response.Write "End function" & vbCrLf
Response.Write "function setContent(zhi,TemplateType)" & vbCrLf
Response.Write "End function" & vbCrLf
Response.Write "</script>" & vbCrLf
Response.Write "<script language = 'JavaScript'>" & vbCrLf
Response.Write "   setTimeout(""EditorFullScreen.showBorders()"",2000);" & vbCrLf
Response.Write "</script>" & vbCrLf
Response.Write "</BODY>" & vbCrLf
Response.Write "</HTML>" & vbCrLf

%>