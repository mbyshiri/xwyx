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
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改代码属性</title>
</head>
<script language="javascript" event="onerror(msg, url, line)" for="window">return true;</script>
<body scroll=no topmargin="0" leftmargin="0">
  <table cellSpacing=0 cellPadding=0 width="100%" border=0>
    <tr>
      <td vAlign=center width=400>
       <TEXTAREA id=EditTagCode style="WIDTH: 470px; HEIGHT: 240px" name=EditTagCode></TEXTAREA></td>
      <td vAlign=top width=100>
        <table cellSpacing=0 cellPadding=0 width="100%" border=0>
          <tr>
            <td height=70><Input onclick=SetAttribute(); type=button value=" 设 置 " name=BtnSet></td>
          </tr>
          <tr>
            <td height=40><Input onclick=window.close(); type=button value=" 取 消 " name=Submit>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
<script language="VBScript">
    Function  Resumeblank(byval Content)
        Dim strHtml,strHtml2,Num,Numtemp,StrTemp
        strHtml=Replace(Content, "<DIV", "<div")
        strHtml=Replace(strHtml, "</DIV>", "</div>")
        strHtml=Replace(strHtml, "<TABLE", "<table")
        strHtml=Replace(strHtml, "</TABLE>", vbCrLf & "</table>"& vbCrLf)
        strHtml=Replace(strHtml, "<TBODY>", "")
        strHtml=Replace(strHtml, "</TBODY>","")
        strHtml=Replace(strHtml, "<TR", "<tr")
        strHtml=Replace(strHtml, "</TR>", vbCrLf & "</tr>"& vbCrLf)
        strHtml=Replace(strHtml, "<TD", "<td")
        strHtml=Replace(strHtml, "</TD>", "</td>")
        strHtml=Replace(strHtml, "<!--", vbCrLf & "<!--")
        strHtml=Replace(strHtml, "<SELECT",vbCrLf & "<Select")
        strHtml=Replace(strHtml, "</SELECT>",vbCrLf & "</Select>")
        strHtml=Replace(strHtml, "<OPTION",vbCrLf & "  <Option")
        strHtml=Replace(strHtml, "</OPTION>","</Option>")
        strHtml=Replace(strHtml, "<INPUT",vbCrLf & "  <Input")
        strHtml=Replace(strHtml, "<"& "script",vbCrLf & "<" &"script")
        strHtml=Replace(strHtml, "&amp;","&")
        strHtml=Replace(strHtml, "{$--",vbCrLf & "<!--$")
        strHtml=Replace(strHtml, "--}","$-->")
        arrContent = Split(Trim(strHtml),vbCrLf)
        For i = 0 To UBound(arrContent)
            Numtemp=false
            if Instr(arrContent(i),"<table")>0 then
                Numtemp=True
                if Strtemp<>"<table" and Strtemp <>"</table>" then
                  Num=Num+2
                End if 
                Strtemp="<table"
            elseif Instr(arrContent(i),"<tr")>0 then
                Numtemp=True
                if Strtemp<>"<tr" and Strtemp<>"</tr>" then
                  Num=Num+2
                End if 
                Strtemp="<tr"
            elseif Instr(arrContent(i),"<td")>0 then
                Numtemp=True
                if Strtemp<>"<td" and Strtemp<>"</td>" then
                  Num=Num+2
                End if 
                Strtemp="<td"
            elseif Instr(arrContent(i),"</table>")>0 then
                Numtemp=True
                if Strtemp<>"</table>" and Strtemp<>"<table" then
                  Num=Num-2
                End if 
                Strtemp="</table>"
            elseif Instr(arrContent(i),"</tr>")>0 then
                Numtemp=True
                if Strtemp<>"</tr>" and Strtemp<>"<tr" then
                  Num=Num-2
                End if 
                Strtemp="</tr>"
            elseif Instr(arrContent(i),"</td>")>0 then
                Numtemp=True
                if Strtemp<>"</td>" and Strtemp<>"<td" then
                  Num=Num-2
                End if 
                Strtemp="</td>"
            elseif Instr(arrContent(i),"<"&"!--")>0 then
                Numtemp=True
            End if
            if Num< 0 then Num = 0
            if trim(arrContent(i))<>"" then
                if i=0 then
                    strHtml2= string(Num," ") & arrContent(i) 
                elseif Numtemp=True then
                    strHtml2= strHtml2 & vbCrLf & string(Num," ") & arrContent(i) 
                else
                    strHtml2= strHtml2 & vbCrLf & arrContent(i) 
                end if
            end if
        Next
        Resumeblank=strHtml2
    End function

    Function LableFilter(byval Content)
        Dim regEx,Match,Match2,strTemp2,Matches
        Set regEx = New RegExp
        regEx.IgnoreCase = True
        regEx.Global = True
        
        regEx.Pattern = "\<IMG(.[^\<]*)\}['|""]\>"
        Set Matches = regEx.Execute(Content)
        For Each Match In Matches
            regEx.Pattern = "\{\$(.*)\}"
            Set strTemp = regEx.Execute(replace(Match.Value," ",""))
            For Each Match2 In strTemp
                strTemp2 = Replace(Match2.Value, "?", """")
                Content = Replace(Content, Match.Value, "<!--" & strTemp2 & "-->")
            Next
        Next

        regEx.Pattern = "\<IMG(.[^\<]*)\$\>"
        Set Matches = regEx.Execute(Content)
        For Each Match In Matches
            regEx.Pattern = "\#(.*)\#"
            Set strTemp = regEx.Execute(Match.Value)
            For Each Match2 In strTemp
                strTemp2 = Replace(Match2.Value, "?", "?")
                strTemp2 = Replace(Match2.Value, "&amp;", "&")
                strTemp2 = Replace(strTemp2, "#", "")
                strTemp2 = Replace(strTemp2,"&13;&10;",vbCrLf)
                strTemp2 = Replace(strTemp2,"&9;",vbTab)
                strTemp2 = Replace(strTemp2, "[!", "<")
                strTemp2 = Replace(strTemp2, "!]", ">")
                Content = Replace(Content, Match.Value, strTemp2)
            Next
        Next
        Set regEx = Nothing
        LableFilter=Content
    
    End Function

    Function ShiftCharacter(ByVal Content)

        Dim regEx, Match, Matches, strTemp,arrMatch, strMatch, i
                Set regEx = New RegExp
        regEx.IgnoreCase = True
        regEx.Global = True
        '替换文件的注解函数符，解决不显示问题
        Content = Replace(Content, "<"&"!--{$", "{$")
        Content = Replace(Content, "}"&"-->", "}")
        Content = Replace(Content, "<NOSCRIPT><IFRAME src='*' Width='0' Height='0'></IFRAME></NOSCRIPT>", "")
          
        '图片替换JS
        regEx.Pattern = "(\<"&"Script)(.[^\<]*)(\<\/Script\>)"
        Set Matches = regEx.Execute(Content)
        For Each Match In Matches
            strTemp = Replace(Match.Value, "<", "[!")
            strTemp = Replace(strTemp, ">", "!]")
            strTemp = Replace(strTemp, "'", """")
            strTemp = "<IMG alt='#" & strTemp & "#' src=""editor/images/jscript.gif"" border=0 $>"
            Content = Replace(Content, Match.Value, strTemp)
        Next    
        '图片替换超级标签
        regEx.Pattern = "\{\$GetPicArticle\((.*?)\)\}"
        Content = regEx.Replace(Content, "<IMG src=""editor/images/label.gif"" border=0 zzz='{$GetPicArticle($1)}'>")
        regEx.Pattern = "\{\$GetArticleList\((.*?)\)\}"
        Content = regEx.Replace(Content, "<IMG src=""editor/images/label.gif"" border=0 zzz='{$GetArticleList($1)}'>")
        regEx.Pattern = "\{\$GetSlidePicArticle\((.*?)\)\}"
        Content = regEx.Replace(Content, "<IMG src=""editor/images/label.gif"" border=0 zzz='{$GetSlidePicArticle($1)}'>")
        regEx.Pattern = "\{\$GetPicSoft\((.*?)\)\}"
        Content = regEx.Replace(Content, "<IMG src=""editor/images/label.gif"" border=0 zzz='{$GetPicSoft($1)}'>")
        regEx.Pattern = "\{\$GetSoftList\((.*?)\)\}"
        Content = regEx.Replace(Content, "<IMG src=""editor/images/label.gif"" border=0 zzz='{$GetSoftList($1)}'>")
        regEx.Pattern = "\{\$GetSlidePicSoft\((.*?)\)\}"
        Content = regEx.Replace(Content, "<IMG src=""editor/images/label.gif"" border=0 zzz='{$GetSlidePicSoft($1)}'>")
        regEx.Pattern = "\{\$GetPicPhoto\((.*?)\)\}"
        Content = regEx.Replace(Content, "<IMG src=""editor/images/label.gif"" border=0 zzz='{$GetPicPhoto($1)}'>")
        regEx.Pattern = "\{\$GetPhotoList\((.*?)\)\}"
        Content = regEx.Replace(Content, "<IMG src=""editor/images/label.gif"" border=0 zzz='{$GetPhotoList($1)}'>")
        regEx.Pattern = "\{\$GetSlidePicPhoto\((.*?)\)\}"
        Content = regEx.Replace(Content, "<IMG src=""editor/images/label.gif"" border=0 zzz='{$GetSlidePicPhoto($1)}'>")
        regEx.Pattern = "\{\$GetPicProduct\((.*?)\)\}"
        Content = regEx.Replace(Content, "<IMG src=""editor/images/label.gif"" border=0 zzz='{$GetPicProduct($1)}'>")
        regEx.Pattern = "\{\$GetProductList\((.*?)\)\}"
        Content = regEx.Replace(Content, "<IMG src=""editor/images/label.gif"" border=0 zzz='{$GetProductList($1)}'>")
        regEx.Pattern = "\{\$GetSlidePicProduct\((.*?)\)\}"
        Content = regEx.Replace(Content, "<IMG src=""editor/images/label.gif"" border=0 zzz='{$GetSlidePicProduct($1)}'>")
        regEx.Pattern = "\{\$GetPositionList\((.*?)\)\}"
        Content = regEx.Replace(Content, "<IMG src=""editor/images/label.gif"" border=0 zzz='{$GetPositionList($1)}'>")
        regEx.Pattern = "\{\$GetSearchResult\((.*?)\)\}"
        Content = regEx.Replace(Content, "<IMG src=""editor/images/label.gif"" border=0 zzz='{$GetSearchResult($1)}'>")
        Set regEx = Nothing    
        ShiftCharacter = Content
        
    End Function
</script>
<script language="JavaScript">
    var oControl;
    var oSeletion;
    var sRangeType;
    oSelection = dialogArguments.HtmlEdit.document.selection.createRange();
    sRangeType = dialogArguments.HtmlEdit.document.selection.type;
    if (sRangeType == "Control") {
            oControl = oSelection.item(0);
    }else {
        if (dialogArguments.HtmlEdit!=null) oControl=dialogArguments.HtmlEdit;
    }
    if (oControl==null){
        document.all.BtnSet.disabled=true;
    }else{
        document.all.EditTagCode.value=LableFilter(Resumeblank(oControl.outerHTML));
    }
    function SetAttribute()
    {            
        oControl.outerHTML=ShiftCharacter(document.all.EditTagCode.value);
        window.close();
    }
</script>

</body>
</html>