<!--#include file="../Include/PowerEasy.Common.Content.asp"-->
<!--#include file="../Include/PowerEasy.XmlHttp.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim KindType, LinkType, ShowType, KindID

Dim iMod
Dim HtmlDir, strHtml, PageTitle, strNavPath, strPageTitle, strTemplate, arrTemplate
Dim PrevChannelID, strTempContent, strContentPageTitleArr

Dim strNavLink
Dim strYear, strMonth, strDay, strListStr_Font
Dim strGirl, strMan, Secrit, NoEnter

'从语言包中读取的相应变量
Dim strTop, strElite, strCommon, strNew, strHot
Dim strTop2, strElite2, strHot2
Dim Character_Author, Character_Date, Character_Hits, Character_Class
Dim SearchResult_Content_NoPurview, SearchResult_ContentLenth
Dim strList_Content_Div
Dim strList_Title, strComment

XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))

strYear = XmlText("BaseText", "Year", "年")
strMonth = XmlText("BaseText", "Month", "月")
strDay = XmlText("BaseText", "Day", "日")
strNavLink = XmlText("BaseText", "NavLink", "&gt;&gt;")
strNavPath = XmlText("BaseText", "Nav", "您现在的位置：") & "&nbsp;<a class='LinkPath' href='" & SiteUrl & "'>" & SiteName & "</a>"
strGirl = XmlText("BaseText", "Girl", "女")
strMan = XmlText("BaseText", "Man", "男")
Secrit = XmlText("BaseText", "Secrit", "保密")
NoEnter = XmlText("BaseText", "NoEnter", "未填")
FileExt_SiteSpecial = arrFileExt(FileExt_SiteSpecial)


'=================================================
'函数名：UrlPrefix
'作  用：如果使用绝对路径，并且频道地址不包括域名，则链接地址前缀为域名
'参  数：iUrlType ---- 链接地址类型，0为相对路径，1为绝对路径
'        strChannelUrl ---- 频道地址
'=================================================
Function UrlPrefix(iUrlType, strChannelUrl)
    Dim strUrlPrefix
    strUrlPrefix = ""
    If iUrlType = 1 And Left(strChannelUrl, 1) = "/" Then
        strUrlPrefix = "http://" & Trim(Request.ServerVariables("HTTP_HOST"))
    End If
    UrlPrefix = strUrlPrefix
End Function


Function GetInfoList_StrNoItem(arrClassID, iSpecialID, IsHot, IsElite, strHot, strElite)
    Dim strNoItem, strThis
    strThis = ""
    If arrClassID <> "0" Then
        strThis = "此栏目下"
    End If
    If iSpecialID > 0 Then
        strThis = "此专题下"
    End If
    If IsHot = False And IsElite = False Then
        strNoItem = "<li>" & strThis & "没有" & ChannelShortName & "</li>"
    ElseIf IsHot = True And IsElite = False Then
        strNoItem = "<li>" & strThis & "没有" & strHot & ChannelShortName & "</li>"
    ElseIf IsHot = False And IsElite = True Then
        strNoItem = "<li>" & strThis & "没有" & strElite & ChannelShortName & "</li>"
    Else
        strNoItem = "<li>" & strThis & "没有" & strHot & strElite & ChannelShortName & "</li>"
    End If
    GetInfoList_StrNoItem = strNoItem
End Function

Function GetInfoList_GetStrTitle(Title, TitleLen, TitleFontType, TitleFontColor)
    Dim strTitle
    If TitleLen > 0 Then
        strTitle = ReplaceText(GetSubStr(Title, TitleLen, ShowSuspensionPoints), 2)
    Else
        strTitle = ReplaceText(Title, 2)
    End If
    Select Case TitleFontType
    Case 1
        strTitle = "<b>" & strTitle & "</b>"
    Case 2
        strTitle = "<em>" & strTitle & "</em>"
    Case 3
        strTitle = "<b><em>" & strTitle & "</em></b>"
    End Select
    If TitleFontColor <> "" Then
        strTitle = "<font color=""" & TitleFontColor & """>" & strTitle & "</font>"
    End If
    GetInfoList_GetStrTitle = strTitle
End Function

Function GetInfoList_GetStrProperty(ShowPropertyType, OnTop, Elite, iNumber, strCommon, strTop, strElite)
    Dim strProperty
    Select Case ShowPropertyType
    Case 0
        strProperty = ""
    Case 1
        If OnTop = True Then
            strProperty = "<img src=""" & ChannelUrl & "/images/" & ModuleName & "_ontop.gif"" alt=""" & strTop & ChannelShortName & """>"
        ElseIf Elite = True Then
            strProperty = "<img src=""" & ChannelUrl & "/images/" & ModuleName & "_elite.gif"" alt=""" & strElite & ChannelShortName & """>"
        Else
            strProperty = "<img src=""" & ChannelUrl & "/images/" & ModuleName & "_common.gif"" alt=""" & strCommon & ChannelShortName & """>"
        End If
    Case 2
        strProperty = "·"
    Case 11
        strProperty = iNumber
    Case Else
        If OnTop = True Then
            strProperty = "<img src=""" & ChannelUrl & "/images/" & ModuleName & "_ontop" & ShowPropertyType - 1 & ".gif"" alt=""" & strTop & ChannelShortName & """>"
        ElseIf Elite = True Then
            strProperty = "<img src=""" & ChannelUrl & "/images/" & ModuleName & "_elite" & ShowPropertyType - 1 & ".gif"" alt=""" & strElite & ChannelShortName & """>"
        Else
            strProperty = "<img src=""" & ChannelUrl & "/images/" & ModuleName & "_common" & ShowPropertyType - 1 & ".gif"" alt=""" & strCommon & ChannelShortName & """>"
        End If
    End Select
    GetInfoList_GetStrProperty = strProperty
End Function

Function GetInfoList_GetStrClassLink(Character_Class, Css_ListItem, ClassID_ListItem, ClassName_ListItem, ClassUrl_ListItem)
    Dim strClassName
    If ClassID_ListItem <> -1 Then
        strClassName = Replace(Character_Class, "{$Text}", "<a class=""" & Css_ListItem & """ href=""" & ClassUrl_ListItem & """>" & ClassName_ListItem & "</a>")
    Else
        strClassName = ""
    End If
    GetInfoList_GetStrClassLink = strClassName
End Function

Function GetInfoList_GetStrInfoLink(strList_Title, ShowTips, OpenType, Css_ListItem, Title_ListItem, InfoUrl, LinkTips_ListItem, Author_ListItem, UpdateTime_ListItem)
    Dim strInfoLink, strTemp
    strInfoLink = "<a class=""" & Css_ListItem & """ href=""" & InfoUrl & """"
    If ShowTips = True Then
        strTemp = Replace(strList_Title, "{$Title}", LinkTips_ListItem)
        strTemp = Replace(strTemp, "{$PhotoName}", LinkTips_ListItem)
        strTemp = Replace(strTemp, "{$SoftName}", LinkTips_ListItem)
        strTemp = Replace(strTemp, "{$Author}", Author_ListItem)
        strTemp = Replace(strTemp, "{$UpdateTime}", UpdateTime_ListItem)
        strTemp = Replace(strTemp, "{$br}", vbCrLf)
        strInfoLink = strInfoLink & " title=""" & strTemp & """"
    Else
        strInfoLink = strInfoLink & " title=""" & LinkTips_ListItem & """"
    End If
    If OpenType = 0 Then
        strInfoLink = strInfoLink & " target=""_self"">"
    Else
        strInfoLink = strInfoLink & " target=""_blank"">"
    End If
    strInfoLink = strInfoLink & Title_ListItem & "</a>"
    GetInfoList_GetStrInfoLink = strInfoLink
End Function

Function GetInfoList_GetStrUpdateTime(UpdateTime, ShowDateType)
    Dim strUpdateTime
    If Not IsDate(UpdateTime) Then
        GetInfoList_GetStrUpdateTime = ""
        Exit Function
    End If
    Select Case ShowDateType
    Case 1
        strUpdateTime = Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2)
    Case 2
        strUpdateTime = Month(UpdateTime) & strMonth & Day(UpdateTime) & strDay
    Case 3
        strUpdateTime = Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2)
    Case 4
        strUpdateTime = Year(UpdateTime) & strYear & Month(UpdateTime) & strMonth & Day(UpdateTime) & strDay
    Case 5
        strUpdateTime = UpdateTime
    Case 6
        strUpdateTime = UpdateTime
    End Select
    If DateDiff("D", UpdateTime, Now()) < DaysOfNew Then
        strUpdateTime = "<font " & strListStr_Font & ">" & strUpdateTime & "</font>"
    End If
    GetInfoList_GetStrUpdateTime = strUpdateTime
End Function

Function GetInfoList_GetStrAuthorDateHits(ShowAuthor, ShowDateType, ShowHits, Author_ListItem, UpdateTime_ListItem, Hits_ListItem, iChannelID)
    Dim strAuthorDateHits
    strAuthorDateHits = ""
    If ShowAuthor = True Or ShowDateType > 0 Or ShowHits = True Then
        strAuthorDateHits = strAuthorDateHits & "&nbsp;("
        If ShowAuthor = True Then
            strAuthorDateHits = strAuthorDateHits & Author_ListItem
        End If
        If ShowDateType > 0 Then
            If ShowAuthor = True Then
                strAuthorDateHits = strAuthorDateHits & "，"
            End If
            strAuthorDateHits = strAuthorDateHits & UpdateTime_ListItem
        End If
        If ShowHits = True Then
            If ShowAuthor = True Or ShowDateType > 0 Then
                strAuthorDateHits = strAuthorDateHits & "，"
            End If
            strAuthorDateHits = strAuthorDateHits & Hits_ListItem
        End If
        strAuthorDateHits = strAuthorDateHits & ")"
    End If
    GetInfoList_GetStrAuthorDateHits = strAuthorDateHits
End Function

Function GetInfoList_GetStrHotSign(ShowHotSign, Hits_ListItem, strHot)
    If ShowHotSign = True And Hits_ListItem >= HitsOfHot Then
        GetInfoList_GetStrHotSign = "<img src=""" & strInstallDir & "images/hot.gif"" alt=""" & strHot & ChannelShortName & """>"
    Else
        GetInfoList_GetStrHotSign = ""
    End If
End Function

Function GetInfoList_GetStrNewSign(ShowNewSign, UpdateTime_ListItem, strNew)
    If ShowNewSign = True And DateDiff("D", UpdateTime_ListItem, Now()) < DaysOfNew Then
        GetInfoList_GetStrNewSign = "<img src=""" & strInstallDir & "images/new.gif"" alt=""" & strNew & ChannelShortName & """>"
    Else
        GetInfoList_GetStrNewSign = ""
    End If
End Function


Function GetInfoList_GetStrContent(ContentLen, Content_ListItem, Intro_ListItem)
    If Trim(Intro_ListItem & "") = "" Then
        GetInfoList_GetStrContent = Left(Replace(Replace(Replace(nohtml(Content_ListItem), "[NextPage]", ""), ">", "&gt;"), "<", "&lt;"), ContentLen) & "……"
    Else
        GetInfoList_GetStrContent = Left(nohtml(Intro_ListItem), ContentLen)
    End If
End Function

Function GetInfoList_GetStrAuthor_Xml(ShowAuthor, strAuthor)
    If ShowAuthor = True Then
        GetInfoList_GetStrAuthor_Xml = Replace(Character_Author, "{$Text}", strAuthor)
    Else
        GetInfoList_GetStrAuthor_Xml = ""
    End If
End Function

Function GetInfoList_GetStrUpdateTime_Xml(ShowDateType, strUpdateTime)
    If ShowDateType > 0 Then
        GetInfoList_GetStrUpdateTime_Xml = Replace(Character_Date, "{$Text}", strUpdateTime)
    Else
        GetInfoList_GetStrUpdateTime_Xml = ""
    End If
End Function

Function GetInfoList_GetStrHits_Xml(ShowHits, strHits)
    If ShowHits = True Then
        GetInfoList_GetStrHits_Xml = Replace(Character_Hits, "{$Text}", strHits)
    Else
        GetInfoList_GetStrHits_Xml = ""
    End If
End Function

Function GetInfoList_GetStrAuthor_RSS(Author)
    If Trim(Author & "") = "" Then
        GetInfoList_GetStrAuthor_RSS = "本站原创"
    Else
        GetInfoList_GetStrAuthor_RSS = xml_nohtml(Author)
    End If
End Function

Function GetInfoList_GetStrRSS(strTitle, strLink, strContent, strAuthor, strClassName, strUpdateTime)
    XMLDOM.appendChild (XMLDOM.createElement("item"))
    Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("title"))
    Node.Text = xml_nohtml(strTitle)
    Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("link"))
    Node.Text = strLink
    Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("description"))
    Node.Text = strContent
    Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("author"))
    Node.Text = strAuthor
    Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("category"))
    Node.Text = strClassName
    Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("pubDate"))
    Node.Text = strUpdateTime
    GetInfoList_GetStrRSS = XMLDOM.documentElement.xml
End Function
'=================================================
'过程名：ShowVoteJS_Comment()
'作  用：评论输入判断
'参  数：无
'=================================================
Function ShowVoteJS_Comment()
    Dim strJS
    strJS = "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf	
    strJS = strJS & "function CreateAjax() {" & vbCrLf
    strJS = strJS & "    var XMLHttp;" & vbCrLf
    strJS = strJS & "    if(window.XMLHttpRequest) {" & vbCrLf
    strJS = strJS & "        XMLHttp = new XMLHttpRequest(); //firefox下执行此语句" & vbCrLf
    strJS = strJS & "    }" & vbCrLf
    strJS = strJS & "    else if(window.ActiveXObject){" & vbCrLf
    strJS = strJS & "        try{" & vbCrLf
    strJS = strJS & "            XMLHttp = new ActiveXObject(""Msxm12.XMLHTTP"");" & vbCrLf
    strJS = strJS & "        }catch(e){" & vbCrLf
    strJS = strJS & "            try{" & vbCrLf
    strJS = strJS & "                XMLHttp = new ActiveXObject(""Microsoft.XMLHTTP"");" & vbCrLf
    strJS = strJS & "            }" & vbCrLf
    strJS = strJS & "			catch(e)" & vbCrLf
    strJS = strJS & "			{" & vbCrLf
    strJS = strJS & "    XMLHttp = false;" & vbCrLf    			    
    strJS = strJS & "			}" & vbCrLf
    strJS = strJS & "        }" & vbCrLf
    strJS = strJS & "    }" & vbCrLf
    strJS = strJS & "    return XMLHttp;" & vbCrLf
    strJS = strJS & "}" & vbCrLf
    strJS = strJS & "function Support(id)" & vbCrLf
    strJS = strJS & "{" & vbCrLf
    strJS = strJS & "	_xmlhttp = CreateAjax();" & vbCrLf
    strJS = strJS & "	var url = '"&ChannelUrl&"/Comment.asp?Action=UpdateVote&votetype=1&id='+id+'&n='+Math.random()+'';	" & vbCrLf	
    strJS = strJS & "	if(_xmlhttp)" & vbCrLf 
    strJS = strJS & "    {" & vbCrLf 
    strJS = strJS & "        var content = document.getElementById(""Support""+id);" & vbCrLf      
    strJS = strJS & "		var Support = document.getElementById(""SupportCount""+id);" & vbCrLf					
    strJS = strJS & "        _xmlhttp.open('GET',url,true);" & vbCrLf
    strJS = strJS & "        _xmlhttp.onreadystatechange=function()" & vbCrLf
    strJS = strJS & "        {" & vbCrLf
    strJS = strJS & "            if(_xmlhttp.readyState == 4)" & vbCrLf
    strJS = strJS & "            {" & vbCrLf
    strJS = strJS & "                if(_xmlhttp.status == 200)" & vbCrLf      
    strJS = strJS & "               {" & vbCrLf
    strJS = strJS & "                    var ResponseText = unescape(_xmlhttp.responseText);	" & vbCrLf		
    strJS = strJS & "                    Support.innerHTML=ResponseText;	" & vbCrLf					
    strJS = strJS & "                    content.innerHTML='已支持';" & vbCrLf
    strJS = strJS & "                }" & vbCrLf
    strJS = strJS & "            }" & vbCrLf
    strJS = strJS & "        }" & vbCrLf
    strJS = strJS & "        _xmlhttp.send(null); " & vbCrLf 
    strJS = strJS & "    }" & vbCrLf
    strJS = strJS & "    else" & vbCrLf    
    strJS = strJS & "   {" & vbCrLf
    strJS = strJS & "        alert(""您的浏览器不支持或未启用 XMLHttp!"");" & vbCrLf
    strJS = strJS & "    }" & vbCrLf
    strJS = strJS & "}" & vbCrLf

    strJS = strJS & "function Opposed(id)" & vbCrLf
    strJS = strJS & "{" & vbCrLf
    strJS = strJS & "	_xmlhttp = CreateAjax();" & vbCrLf
    strJS = strJS & "	var url = '"&ChannelUrl&"/Comment.asp?Action=UpdateVote&votetype=2&id='+id+'&n='+Math.random()+'';	" & vbCrLf	
    strJS = strJS & "	if(_xmlhttp)    " & vbCrLf
     strJS = strJS & "   {" & vbCrLf
    strJS = strJS & "        var content = document.getElementById(""Opposed""+id);  " & vbCrLf    
    strJS = strJS & "        var Opposed = document.getElementById(""OpposedCount""+id);" & vbCrLf					
    strJS = strJS & "        _xmlhttp.open('GET',url,true);" & vbCrLf
    strJS = strJS & "        _xmlhttp.onreadystatechange=function()" & vbCrLf
    strJS = strJS & "        {" & vbCrLf
    strJS = strJS & "            if(_xmlhttp.readyState == 4)" & vbCrLf  
    strJS = strJS & "            {" & vbCrLf
    strJS = strJS & "                if(_xmlhttp.status == 200)     " & vbCrLf 
    strJS = strJS & "                {" & vbCrLf
    strJS = strJS & "                    var ResponseText = unescape(_xmlhttp.responseText);	" & vbCrLf
    strJS = strJS & "                    Opposed.innerHTML=ResponseText;	" & vbCrLf		
    strJS = strJS & "                    content.innerHTML='已反对';" & vbCrLf
    strJS = strJS & "                }" & vbCrLf
    strJS = strJS & "            }" & vbCrLf
    strJS = strJS & "        }" & vbCrLf
    strJS = strJS & "        _xmlhttp.send(null);" & vbCrLf  
    strJS = strJS & "    }" & vbCrLf
    strJS = strJS & "    else " & vbCrLf   
    strJS = strJS & "    {" & vbCrLf
    strJS = strJS & "        alert(""您的浏览器不支持或未启用 XMLHttp!"");" & vbCrLf
    strJS = strJS & "    }" & vbCrLf
    strJS = strJS & "}" & vbCrLf
    strJS = strJS & "</script>" & vbCrLf		
    ShowVoteJS_Comment = strJS
End Function

'==================================================
'函数名：SlidePicJs
'功能：         全站通用幻灯片标签
'PicModuleType: 频道类型,文章频道--1,软件频道--2,图片频道--3,商城频道--5
'Elite          是否推荐,1--取出推荐文章,0--取出所有信息
'Ontop          是否置顶,1--取出置顶文章,0--取出所有信息
'Hits            热门文章,具体数字--调用点击数大于次数字的信息,0--取出所有信息 
'PicNum:        调用图片数量,可选范围是2到9个
'PicChannelID : 频道id,如果填0则调用所有同类频道，不同频道用|隔开
'PicClassID : 栏目id,如果填0则调用所有栏目，不同栏目用|隔开
'TitleLength :  标题字数
'PicWid :       幻灯片宽度
'PicHei :       幻灯片高度
'TextHei:       文字高度
'==================================================
Function SlidePicJs(PicModuleType, Elite, Ontop, Hits, PicNum, PicChannelID, PicClassID,TitleLength, PicWid, PicHei, TextHei)
    Dim rsPic, strSlideTemp, sqlSlide, strTitle,arrClassID
    Dim AdID, AdUrl, AdTitle, AdPic, k, i, m
    Dim PrevChannelID,ChannelUrl
    PicModuleType = PE_Clng(PicModuleType)	
    TextHei = PE_Clng(Trim(TextHei))
    PicNum = PE_Clng(Trim(PicNum))
    PicWid = PE_Clng(PicWid)
    PicHei = PE_Clng(PicHei)
    PicChannelID = ReplaceLabelBadChar(Replace(Trim(PicChannelID), "|", ","))
    PicClassID = ReplaceLabelBadChar(Replace(Trim(PicClassID), "|", ","))
    If LCase(Elite) = "true" or PE_Clng(Elite) = 1 then
        Elite = 1
    Else 
        Elite = 0
    End If
    If LCase(Ontop) = "true" or PE_Clng(Ontop) = 1 then
        Ontop = 1
    Else 
        Ontop = 0
    End If
    If LCase(Hits) = "true"  then
        Hits = 500
    Elseif PE_Clng(Hits)>0 then
        Hits = PE_Clng(Hits)
    End If
	
    If PicNum > 9 Or PicNum = 0 Then PicNum = 9
    If TextHei = 0 Then TextHei = 10
    If PicWid = 0 Then PicWid = 300
    If PicHei = 0 Then PicHei = 200
    If TitleLength < 0 Or TitleLength > 200 Then TitleLength = 50
    k = PicNum
    sqlSlide = "select "
    Select Case PicModuleType
    Case 1
        sqlSlide = sqlSlide & "top " & PicNum & " M.Title as tTitle, M.ChannelID as tChannelID, M.DefaultPicUrl as tPicUrl, C.LinkUrl, M.ArticleID as tID, C.UploadDir,  C.ModuleType, C.ChannelDir from PE_Article M"
    Case 2
        sqlSlide = sqlSlide & "top " & PicNum & " M.SoftName as tTitle, M.ChannelID as tChannelID, M.SoftPicUrl as tPicUrl, C.LinkUrl, M.SoftID as tID, C.UploadDir, C.ModuleType, C.ChannelDir from PE_Soft M"
    Case 3
        sqlSlide = sqlSlide & "top " & PicNum & " M.PhotoName as tTitle, M.ChannelID as tChannelID, M.PhotoThumb as tPicUrl, C.LinkUrl, M.PhotoID as tID, C.UploadDir,  C.ModuleType, C.ChannelDir from PE_Photo M"
    Case 5
        sqlSlide = sqlSlide & "top " & PicNum & " M.ProductName as tTitle, M.ChannelID as tChannelID, M.ProductThumb as tPicUrl, C.LinkUrl, M.ProductID as tID, C.UploadDir,  C.ModuleType, C.ChannelDir from PE_Product M"
    Case Else
        sqlSlide = sqlSlide & "top " & PicNum & " M.Title as tTitle, M.ChannelID as tChannelID, M.DefaultPicUrl as tPicUrl, C.LinkUrl, M.ArticleID as tID, C.UploadDir,  C.ModuleType, C.ChannelDir from PE_Article M"
    End Select
    sqlSlide = sqlSlide & " Inner Join PE_Channel C On M.ChannelID = C.ChannelID Where Deleted=" & PE_False
    If  Elite = 1 then 
        Select Case PicModuleType
        Case 1,2,3
            sqlSlide = sqlSlide & " and Elite = " & PE_True 
        Case 5
            sqlSlide = sqlSlide & " and IsElite = " & PE_True 
        Case Else
            sqlSlide = sqlSlide & " and Elite = " & PE_True 
        End Select	
    End If
    If  Ontop = 1 then 
        sqlSlide = sqlSlide & " and Ontop = " & PE_True 
    End If
    If  Hits > 0 then 
        Select Case PicModuleType
        Case 1,2,3
            sqlSlide = sqlSlide & " and Hits > " & Hits
        Case 5
            sqlSlide = sqlSlide & " and IsHot = " & PE_True 
        Case Else
            sqlSlide = sqlSlide & " and Hits > " & Hits		
        End Select	
    End If
    Select Case PicModuleType
    Case 1
        sqlSlide = sqlSlide & " and status=3 and M.DefaultPicUrl<>''"
    Case 2
        sqlSlide = sqlSlide & " and status=3 and M.SoftPicUrl<>''"
    Case 3
        sqlSlide = sqlSlide & ""	
    Case 5
        sqlSlide = sqlSlide & " and M.ProductThumb<>''"
    Case Else
        sqlSlide = sqlSlide & " and status=3 and M.DefaultPicUrl<>''"
    End Select
    If PicClassID <> "0" Then
        If InStr(PicClassID, ",") = 0 Then
            Dim trs
            Set trs = Conn.Execute("select arrChildID from PE_Class where ClassID=" & PE_CLng(PicClassID) & "")
            If trs.BOF And trs.EOF Then
                PicClassID = "0"
            Else
                If IsNull(trs(0)) Or Trim(trs(0)) = "" Then
                    PicClassID = "0"
                Else
                    PicClassID = trs(0)
                End If
            End If
            Set trs = Nothing
        End If	
			
        If InStr(PicClassID, ",") > 0 Then
            sqlSlide = sqlSlide & " and M.ClassID in (" & FilterArrNull(PicClassID, ",") & ")"
        Else
            If PE_CLng(PicClassID) > 0 Then sqlSlide = sqlSlide & " and M.ClassID=" & PE_CLng(PicClassID)
        End If	
    End If	 		
    If InStr(PicChannelID, ",") > 0 Then
        sqlSlide = sqlSlide & " and M.ChannelID in (" & FilterArrNull(PicChannelID, ",") & ")"
    Else
        If PE_CLng(PicChannelID) > 0 Then sqlSlide = sqlSlide & " and M.ChannelID=" & PE_CLng(PicChannelID)
    End If	
		
    Select Case PicModuleType
    Case 1
        sqlSlide = sqlSlide & " order by ArticleID desc"
    Case 2
        sqlSlide = sqlSlide & " order by SoftID desc"
    Case 3
        sqlSlide = sqlSlide & " order by PhotoID desc"
    Case 5
        sqlSlide = sqlSlide & " order by ProductID desc"
    Case Else
        sqlSlide = sqlSlide & " order by ArticleID desc"
    End Select
    Set rsPic = Server.CreateObject("Adodb.RecordSet")
    rsPic.Open sqlSlide, Conn, 1, 1
    i = 1
    k = rsPic.RecordCount
    PrevChannelID = 0
    If rsPic.RecordCount < 2 Then
        SlidePicJs = strSlideTemp & "图片数量小于2,无法用幻灯片显示" & vbCrLf
        rsPic.Close
        Set rsPic = Nothing		
        Exit Function		
    End IF
    strSlideTemp = strSlideTemp & "<script type='text/javascript'>" & vbCrLf
    Do While Not rsPic.EOF
        If rsPic("tChannelID") <> PrevChannelID Then
            'If rsPic("ModuleType") = 1 Then
                If IsNull(rsPic("LinkUrl")) Or Trim(rsPic("LinkUrl")) = "" Or Left(strInstallDir, 7) <> "http://" Then
                    ChannelUrl = strInstallDir & rsPic("ChannelDir")
                Else
                    ChannelUrl = rsPic("LinkUrl")
                End If
                If Right(ChannelUrl, 1) = "/" Then
                    ChannelUrl = Left(ChannelUrl, Len(ChannelUrl) - 1)
                End If	
							
                PrevChannelID = rsPic("tChannelID")
            'Else
                'ChannelUrl = strInstallDir & rsPic("ChannelDir")
           ' End If
        End If
        AdTitle = rsPic("tTitle")	
        AdPic = rsPic("tPicUrl")			
        Select Case PicModuleType
        Case 1
            If TitleLength > 0 Then
                AdTitle = ReplaceText(GetSubStr(AdTitle, TitleLength, False), 2)
            End If
            AdUrl = GetInfoUrl(rsPic("tID"), "Article", 1)
            If Left(AdPic, 7) <> "http://" Then
                strSlideTemp = strSlideTemp & "imgUrl" & i & "='" & ChannelUrl & "/" & rsPic("UploadDir") & "/" & AdPic & "';" & vbCrLf
            Else 
                strSlideTemp = strSlideTemp & "imgUrl" & i & "='" & AdPic & "';" & vbCrLf
            End If				
        Case 2
            If TitleLength > 0 Then
                AdTitle = ReplaceText(GetSubStr(AdTitle, TitleLength, False), 2)
            End If
            AdUrl =  GetInfoUrl(rsPic("tID"), "Soft", 1)
            strSlideTemp = strSlideTemp & "imgUrl" & i & "='" & ChannelUrl & "/" & AdPic & "';" & vbCrLf
        Case 3
            If TitleLength > 0 Then
                AdTitle = ReplaceText(GetSubStr(AdTitle, TitleLength, False), 2)
            End If
            AdUrl =  GetInfoUrl(rsPic("tID"), "Photo", 1)
            strSlideTemp = strSlideTemp & "imgUrl" & i & "='"  & ChannelUrl & "/" & rsPic("UploadDir") & "/" & AdPic & "';" & vbCrLf
        Case 5
            If TitleLength > 0 Then
                AdTitle = ReplaceText(GetSubStr(AdTitle, TitleLength, False), 2)
            End If
            AdUrl =  GetInfoUrl(rsPic("tID"), "Product", 1)
            strSlideTemp = strSlideTemp & "imgUrl" & i & "='" &  ChannelUrl & "/" & rsPic("UploadDir") & "/" & AdPic & "';" & vbCrLf
        Case Else
            If TitleLength > 0 Then
                AdTitle = ReplaceText(GetSubStr(AdTitle, TitleLength, False), 2)
            End If
            AdUrl =  GetInfoUrl(rsPic("tID"), "Article", 1)
            If Left(AdPic, 7) <> "http://" Then
                strSlideTemp = strSlideTemp & "imgUrl" & i & "='" & ChannelUrl & "/" & rsPic("UploadDir") & "/" & AdPic & "';" & vbCrLf
            Else 
                strSlideTemp = strSlideTemp & "imgUrl" & i & "='" & AdPic & "';" & vbCrLf
            End If	
        End Select
        strSlideTemp = strSlideTemp & "imgtext" & i & "='" & AdTitle & "'" & vbCrLf
        strSlideTemp = strSlideTemp & "imgLink" & i & "='" & AdUrl & "'" & vbCrLf
        If i >= k Then Exit Do
        i = i + 1
        rsPic.movenext
    Loop
    strSlideTemp = strSlideTemp & "   var focus_width=" & PicWid & vbCrLf
    strSlideTemp = strSlideTemp & "   var focus_height=" & PicHei & vbCrLf
    strSlideTemp = strSlideTemp & "   var text_height=" & TextHei & vbCrLf
    strSlideTemp = strSlideTemp & "   var swf_height = focus_height+text_height" & vbCrLf
    strSlideTemp = strSlideTemp & "   var pics="
    For m = 1 To k
        If m < k Then
            strSlideTemp = strSlideTemp & "imgUrl" & m & "+'|'+"
        Else
            strSlideTemp = strSlideTemp & "imgUrl" & m & "" & vbCrLf
        End If
    Next
    strSlideTemp = strSlideTemp & "   var links="
    For m = 1 To k
        If m < k Then
            strSlideTemp = strSlideTemp & "imgLink" & m & "+'|'+"
        Else
            strSlideTemp = strSlideTemp & "imgLink" & m & "" & vbCrLf
        End If
    Next
    strSlideTemp = strSlideTemp & "   var texts="
    For m = 1 To k
        If m < k Then
            strSlideTemp = strSlideTemp & "imgtext" & m & "+'|'+"
        Else
            strSlideTemp = strSlideTemp & "imgtext" & m & "" & vbCrLf
        End If
    Next
    strSlideTemp = strSlideTemp & "   document.write('<object classid=clsid:d27cdb6e-ae6d-11cf-96b8-444553540000 codebase=http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0 width='+ focus_width +' height='+ swf_height +'>');" & vbCrLf
    strSlideTemp = strSlideTemp & "   document.write('<param name=allowScriptAccess value=sameDomain><param name=movie value=" & strInstallDir & "images/xman.swf><param name=quality value=high><param name=bgcolor value=#F0F0F0>');" & vbCrLf
    strSlideTemp = strSlideTemp & "   document.write('<param name=menu value=false><param name=wmode value=opaque>');" & vbCrLf
    strSlideTemp = strSlideTemp & "   document.write('<param name=FlashVars value=pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'>');" & vbCrLf
    strSlideTemp = strSlideTemp & "   document.write('<embed  height='+ swf_height +' src=" & strInstallDir & "images/xman.swf  wmode=opaque FlashVars=pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+' menu=false bgcolor=#F0F0F0 quality=high width='+ focus_width +' height='+ focus_height +' allowScriptAccess=sameDomain type=application/x-shockwave-flash pluginspage=http://www.macromedia.com/go/getflashplayer />');" & vbCrLf
    strSlideTemp = strSlideTemp & "   document.write('</object>');" & vbCrLf
    strSlideTemp = strSlideTemp & "</script>"
    rsPic.Close
    Set rsPic = Nothing
    SlidePicJs = strSlideTemp
End Function

'==================================================
'函数名：IsLogin
'功能： 判断当前用户是否登录,是的话返回第一个参数,否则返回第二个参数
'==================================================

Function IsLogin(str,Tips)
    If CheckUserLogined() = True Then
        IsLogin = str
    Else
        IsLogin = Tips
    End If
End Function

'==================================================
'函数名：GetUserName
'功能： 取得当前登录的会员的用户名,如果是游客,则用户名为空
'==================================================
Function GetUserName()
    If CheckUserLogined() = True Then
        GetUserName = UserName
	Else
	    GetUserName = ""	
    End If
End Function

'==================================================
'函数名：YN

'功能：     条件判断函数,可以根据条件运算参数的运算来输出相应的结果
'condition: 条件运算参数,根据运行结果,如果是真则输出Fir,否则输出Sec
'Fir:       条件成立的时候输出Fir的内容
'Sec :      条件不成立的时候输出Sec的内容
'==================================================

Function YN(Condition, Fir, Sec)
    If Condition = "" Or IsNull(Condition) Then '条件判断参数为空,则返回Sec的内容
        YN = Sec
	Elseif LCase(Condition)="true" Then
	    YN=Fir 
	Elseif LCase(Condition)="false" Then
	    YN=Sec
    Else
        regEx.Pattern = "^[0-9\<\>\=\%\+\-\*\/\""]+$"    '匹配只是数字还有运算符
        Dim Temp, result
        Temp = regEx.Test(Condition)  '判断是否只有数字和运算符
        If Temp = True Then           '如果只有数字和运算符
		    Condition = Replace(Condition,"%"," mod ")
            result = Eval(Condition)  '执行算术运算
            If (result) Then
                YN = Fir           '计算结果为真,返回条件1
            Else
                YN = Sec             '计算结果为假,返回条件2
            End If
        ElseIf InStr(Condition, "=") Then   '字符串允许等于判断

            Dim Tempequal
            Tempequal = Split(Condition, "=")
            If Tempequal(0) = Tempequal(1) Then
                YN = Fir
            Else
                YN = Sec
            End If
        ElseIf InStr(Condition, "<>") Then   '字符串允许不等于判断
            Dim Tempuneuqal
            Tempuneuqal = Split(Condition, "<>")
            If Tempuneuqal(0) <> Tempuneuqal(1) Then
                YN = Fir
            Else
                YN = Sec
            End If

        Else                            '其他情况都设置成非法参数
            YN = "参数类型不正确"
        End If
    End If
End Function

'==================================================
'函数名：ShowPath
'作  用：显示“你现在所有位置”导航信息
'参  数：无
'==================================================
Function ShowPath()
    If PageTitle <> "" Then
        strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;" & PageTitle
    End If
    ShowPath = strNavPath
End Function


'==================================================
'函数名：GetLogo
'作  用：得到显示网站LOGO的HTML代码
'参  数：无
'==================================================
Function GetLogo(LogoWidth, LogoHeight)
    Dim strLogo, strLogoUrl
    If LogoUrl <> "" Then
        If LCase(Left(LogoUrl, 7)) = "http://" Or Left(LogoUrl, 1) = "/" Then
            strLogoUrl = LogoUrl
        Else
            strLogoUrl = strInstallDir & LogoUrl
        End If
        If LCase(Right(strLogoUrl, 3)) = "swf" Then
            strLogo = "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0'"
            If LogoWidth > 0 Then strLogo = strLogo & " width='" & LogoWidth & "'"
            If LogoHeight > 0 Then strLogo = strLogo & " height='" & LogoHeight & "'"
            strLogo = strLogo & "><param name='movie' value='" & strLogoUrl & "'>"
            strLogo = strLogo & "<param name='wmode' value='transparent'>"
            strLogo = strLogo & "<param name='quality' value='autohigh'>"
            strLogo = strLogo & "<embed"
            If LogoWidth > 0 Then strLogo = strLogo & " width='" & LogoWidth & "'"
            If LogoHeight > 0 Then strLogo = strLogo & " height='" & LogoHeight & "'"
            strLogo = strLogo & " src='" & strLogoUrl & "'"
            strLogo = strLogo & " wmode='transparent'"
            strLogo = strLogo & " quality='autohigh'"
            strLogo = strLogo & "pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>"
            strLogo = strLogo & "</object>"
        Else
            strLogo = "<a href='" & SiteUrl & "' title='" & SiteName & "' target='_blank'>"
            strLogo = strLogo & "<img src='" & strLogoUrl & "'"
            If LogoWidth > 0 Then strLogo = strLogo & " width='" & LogoWidth & "'"
            If LogoHeight > 0 Then strLogo = strLogo & " height='" & LogoHeight & "'"
            strLogo = strLogo & " border='0'>"
            strLogo = strLogo & "</a>"
        End If
    End If
    GetLogo = strLogo
End Function

'==================================================
'过程名：GetBanner
'作  用：得到网站Banner的HTML代码
'参  数：无
'==================================================
Function GetBanner(BannerWidth, BannerHeight)
    Dim strBanner
    If BannerUrl <> "" Then
        If LCase(Right(BannerUrl, 3)) = "swf" Then
            If LCase(Left(BannerUrl, 7)) = "http://" Then
                strBanner = "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='" & BannerWidth & "' height='" & BannerHeight & "'><param name='movie' value='" & BannerUrl & "'><param name='wmode' value='transparent'><param name='quality' value='high'><embed src='" & BannerUrl & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='" & BannerWidth & "' height='" & BannerHeight & "'></embed></object>"
            Else
                strBanner = "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='" & BannerWidth & "' height='" & BannerHeight & "'><param name='movie' value='" & strInstallDir & BannerUrl & "'><param name='wmode' value='transparent'><param name='quality' value='high'><embed src='" & strInstallDir & BannerUrl & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='" & BannerWidth & "' height='" & BannerHeight & "'></embed></object>"
            End If
        Else
            If LCase(Left(BannerUrl, 7)) = "http://" Then
                strBanner = "<a href='" & SiteUrl & "' title='" & SiteName & "'><img src='" & BannerUrl & "' width='" & BannerWidth & "' height='" & BannerHeight & "' border='0'></a>"
            Else
                strBanner = "<a href='" & SiteUrl & "' title='" & SiteName & "'><img src='" & strInstallDir & BannerUrl & "' width='" & BannerWidth & "' height='" & BannerHeight & "' border='0'></a>"
            End If
        End If
    End If
    GetBanner = strBanner
End Function

Function GetChannelList(NumNewLine)
    If ShowSiteChannel = False Then
        GetChannelList = ""
        Exit Function
    End If
    Dim tmpCacheName
    tmpCacheName = "ChannelListHtml_" & ChannelID & "_" & NumNewLine
    If PE_Cache.CacheIsEmpty(tmpCacheName) Then
        Dim rsChannel, strChannel, ChannelLink, ChannelUrl, LineNum
        LineNum = 1
        ChannelLink = XmlText("BaseText", "ChannelLink", "&nbsp;|&nbsp;")
        If ChannelID = 0 Then
            strChannel = ChannelLink & "<a class='Channel2' href='" & strInstallDir & FileName_SiteIndex & "'>" & XmlText("BaseText", "FirstPage", "网站首页") & "</a>" & ChannelLink
        Else
            strChannel = ChannelLink & "<a class='Channel' href='" & strInstallDir & FileName_SiteIndex & "'>" & XmlText("BaseText", "FirstPage", "网站首页") & "</a>" & ChannelLink
        End If
        Set rsChannel = Conn.Execute("select * from PE_Channel order by OrderID")
        Do While Not rsChannel.EOF
            If rsChannel("Disabled") <> True And (rsChannel("ShowName") <> False Or rsChannel("ChannelType") = 2) Then
                If NumNewLine > 0 And LineNum = NumNewLine Then
                    LineNum = 0
                    strChannel = strChannel & "<br>" & ChannelLink
                End If
                '只使用绝对地址时，才使用频道子域名
                If IsNull(rsChannel("LinkUrl")) Or Trim(rsChannel("LinkUrl")) = "" Or Left(strInstallDir, 7) <> "http://" Then
                    ChannelUrl = strInstallDir & rsChannel("ChannelDir")
                Else
                    ChannelUrl = rsChannel("LinkUrl")
                End If
                If rsChannel("ChannelID") = ChannelID Then
                    strChannel = strChannel & "<a class='Channel2' "
                Else
                    strChannel = strChannel & "<a class='Channel' "
                End If
                If rsChannel("ChannelType") <= 1 Then
                    If rsChannel("UseCreateHTML") > 0 Then
                        strChannel = strChannel & " href='" & ChannelUrl & "/Index" & arrFileExt(rsChannel("FileExt_Index")) & "'"
                    Else
                        strChannel = strChannel & " href='" & ChannelUrl & "/Index.asp'"
                    End If
                Else
                    strChannel = strChannel & " href='" & rsChannel("LinkUrl") & "'"
                End If
                If rsChannel("OpenType") = 0 Then
                    strChannel = strChannel & " target='_self'"
                Else
                    strChannel = strChannel & " target='_blank'"
                End If
                
                strChannel = strChannel & " title='" & Trim(nohtml(rsChannel("ReadMe"))) & "'"

                If rsChannel("ChannelPicUrl") = "" Or IsNull(rsChannel("ChannelPicUrl")) = True Then
                    strChannel = strChannel & ">" & rsChannel("ChannelName") & "</a>" & ChannelLink
                Else
                    strChannel = strChannel & "><img src='" & rsChannel("ChannelPicUrl") & "' border=0 alt='" & rsChannel("ChannelName") & "'></a>" & ChannelLink
                End If
                If NumNewLine > 0 Then
                    LineNum = LineNum + 1
                End If
            End If
            rsChannel.MoveNext
        Loop
        rsChannel.Close
        Set rsChannel = Nothing
        PE_Cache.SetValue tmpCacheName, strChannel
    Else
        strChannel = PE_Cache.GetValue(tmpCacheName)
    End If

    GetChannelList = strChannel
End Function

'=================================================
'函数名：GetBrotherClass
'作  用：显示当前栏目的同级栏目
'参  数：
'1       theClassID ---- 栏目ID,0为本栏目
'2       ClassNum ---- 栏目数，若大于0，则只查询前几个栏目
'3       ShowPropertyType ---- 显示栏目前的小图标，0为不显示，1为符号，其他为小图片：/images/article_common*.gif
'4       OpenType ---- 栏目打开方式，0为在原窗口打开，1为在新窗口打开，3为根据栏目设置
'5       Cols ---- 每行的列数。超过此列数就换行。
'=================================================
Function GetBrotherClass(theClassID, ClassNum, ShowPropertyType, OpenType, Cols)
    Dim sqlBro, rsBro, i, strBro, tOpenType

    If Cols = 0 Then Cols = 1
    
    sqlBro = "select"
    If ClassNum > 0 Then
        sqlBro = sqlBro & " top " & ClassNum
    End If
    sqlBro = sqlBro & " ClassID,ClassName,Depth,ParentPath,NextID,ClassType,Child,ParentDir,ClassDir,OpenType,LinkUrl,ClassPurview from PE_Class where ChannelID=" & ChannelID & " "
    
    If theClassID <> 0 Then
        sqlBro = sqlBro & " and ParentID=(select ParentID from PE_Class where ClassID= " & theClassID & ")"
    Else
        sqlBro = sqlBro & " and ParentID=(select ParentID from PE_Class where ClassID= " & ClassID & ")"
    End If

    sqlBro = sqlBro & " and IsElite=" & PE_True & " order by OrderID,RootID"

    Set rsBro = Conn.Execute(sqlBro)
    If rsBro.BOF And rsBro.EOF Then
        strBro = "没有任何同级栏目"
    Else
        i = 0
        Do While Not rsBro.EOF
            If i > 0 Then
                If i Mod Cols = 0 Then
                    strBro = strBro & "<br>"
                Else
                    strBro = strBro & "&nbsp;&nbsp;"
                End If
            
            End If
                
            If ShowPropertyType = 0 Then
                strBro = strBro & ""
            ElseIf ShowPropertyType = 1 Then
                strBro = strBro & "·"
            Else
                strBro = strBro & "<img src='" & ChannelUrl & "/images/" & ModuleName & "_common" & ShowPropertyType & ".gif' border='0'>"
            End If

            If rsBro("ClassType") = 1 Then
                strBro = strBro & "&nbsp;<a class='childclass' href='" & GetClassUrl(rsBro("ParentDir"), rsBro("ClassDir"), rsBro("ClassID"), rsBro("ClassPurview")) & "'"
            Else
                strBro = strBro & "&nbsp;<a class='childclass' href='" & rsBro("LinkUrl") & "'"
            End If
            If OpenType = 3 Then
                tOpenType = rsBro("OpenType")
            Else
                tOpenType = OpenType
            End If
            If tOpenType = 0 Then
                strBro = strBro & " target=""_self"">"
            Else
                strBro = strBro & " target=""_blank"">"
            End If
            strBro = strBro & rsBro("ClassName") & "</a>"

            rsBro.MoveNext
            i = i + 1
        Loop

    End If
    rsBro.Close
    Set rsBro = Nothing
    GetBrotherClass = strBro
End Function

'=================================================
'函数名：GetChildClass
'作  用：显示当前栏目的下一级子栏目
'参  数：
'1       theClassID ---- 栏目ID,0为本栏目
'2       ClassNum ---- 栏目数，若大于0，则只查询前几个栏目
'3       ShowPropertyType ---- 显示栏目前的小图标，0为不显示，1为符号，其他为小图片：/images/article_common*.gif
'4       OpenType ---- 栏目打开方式，0为在原窗口打开，1为在新窗口打开，3为根据栏目设置
'5       Cols ---- 每行的列数。超过此列数就换行。
'6       ShowChildNum ---- 是否显示子栏目个数，有子栏目时才显示，
'=================================================
Function GetChildClass(theClassID, ClassNum, ShowPropertyType, OpenType, Cols, ShowChildNum)
    Dim sqlChild, rsChild, i, strChild, tOpenType

    If Cols = 0 Then Cols = 1
    
    sqlChild = "select"
    If ClassNum > 0 Then
        sqlChild = sqlChild & " top " & ClassNum
    End If
    sqlChild = sqlChild & " ClassID,ClassName,Depth,ParentPath,NextID,ClassType,Child,ParentDir,ClassDir,OpenType,LinkUrl,ClassPurview from PE_Class where ChannelID=" & ChannelID & " "
    
    If theClassID <> 0 Then
        sqlChild = sqlChild & " and ParentID=" & theClassID
    Else
        sqlChild = sqlChild & " and ParentID=" & ClassID
    End If

    sqlChild = sqlChild & " and IsElite=" & PE_True & " order by OrderID"
    Set rsChild = Conn.Execute(sqlChild)
    If rsChild.BOF And rsChild.EOF Then
        strChild = "没有任何子栏目"
    Else
        i = 0
        Do While Not rsChild.EOF
            If i > 0 Then
                If i Mod Cols = 0 Then
                    strChild = strChild & "<br>"
                Else
                    strChild = strChild & "&nbsp;&nbsp;"
                End If
            
            End If
                
            If ShowPropertyType = 0 Then
                strChild = strChild & ""
            ElseIf ShowPropertyType = 1 Then
                strChild = strChild & "·"
            Else
                strChild = strChild & "<img src='" & ChannelUrl & "/images/" & ModuleName & "_common" & ShowPropertyType & ".gif' border='0'>"
            End If

            If rsChild("ClassType") = 1 Then
                strChild = strChild & "&nbsp;<a class='childclass' href='" & GetClassUrl(rsChild("ParentDir"), rsChild("ClassDir"), rsChild("ClassID"), rsChild("ClassPurview")) & "'"
            Else
                strChild = strChild & "&nbsp;<a class='childclass' href='" & rsChild("LinkUrl") & "'"
            End If
            If OpenType = 3 Then
                tOpenType = rsChild("OpenType")
            Else
                tOpenType = OpenType
            End If
            If tOpenType = 0 Then
                strChild = strChild & " target=""_self"">"
            Else
                strChild = strChild & " target=""_blank"">"
            End If
            strChild = strChild & rsChild("ClassName") & "</a>"

            If rsChild("Child") > 0 And ShowChildNum = True Then
                strChild = strChild & "(" & rsChild("Child") & ")"
            End If
            rsChild.MoveNext
            i = i + 1
        Loop

    End If
    rsChild.Close
    Set rsChild = Nothing
    GetChildClass = strChild
End Function

Function GetClassUrl(sParentDir, sClassDir, iClassID, iClassPurview)
    Dim strClassUrl
    If (UseCreateHTML = 1 Or UseCreateHTML = 3) And iClassPurview < 2 Then
        strClassUrl = ChannelUrl & GetListPath(StructureType, ListFileType, sParentDir, sClassDir) & GetListFileName(ListFileType, iClassID, 1, 1) & FileExt_List
    Else
        strClassUrl = ChannelUrl_ASPFile & "/ShowClass.asp?ClassID=" & iClassID
    End If
    GetClassUrl = strClassUrl
End Function

Function GetClass_1Url(sParentDir, sClassDir, iClassID, iClassPurview)
    Dim strClassUrl
    If (UseCreateHTML = 1 Or UseCreateHTML = 3) And iClassPurview < 2 Then
        strClassUrl = ChannelUrl & GetListPath(StructureType, ListFileType, sParentDir, sClassDir) & GetList_1FileName(ListFileType, iClassID) & FileExt_List
    Else
        strClassUrl = ChannelUrl_ASPFile & "/ShowClass.asp?ShowType=2&ClassID=" & iClassID
    End If
    GetClass_1Url = strClassUrl
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

'==================================================
'函数名：GetVote
'作  用：显示网站调查
'参  数：无
'==================================================
Function GetVote()
    Dim strVoteBody
    If PE_Cache.CacheIsEmpty(ChannelID & "_Site_Vote") Then
        Dim sqlVote, rsVote, i, strVote
        sqlVote = "select * from PE_Vote where IsSelected=" & PE_True & " and (ChannelID=-1 or ChannelID=" & ChannelID & ") and IsItem=" & PE_False & " and  EndTime> "& PE_Now &" order by ID Desc"
        Set rsVote = Conn.Execute(sqlVote)
        If rsVote.BOF And rsVote.EOF Then
            strVote = XmlText("Site", "ShowVote/VoteErr", "&nbsp;没有任何调查")
        Else
            Dim j: j = 1
            Dim strVoteContent
            strVoteContent = XmlText("Site", "ShowVote/VoteBody", "<form id='VoteForm{$lid}' name='VoteForm{$lid}' method='post' action='{$strInstallDir}vote.asp' target='_blank'>&nbsp;&nbsp;&nbsp;&nbsp;{$Title}<br>{$VoteBody}<br><input name='VoteType' type='hidden'value='{$VoteType}'><input name='Action' type='hidden' value='Vote'><input name='ID' type='hidden' value='{$ID}'><div align='center'><a href='javascript:VoteForm{$lid}.submit();'><img src='{$strInstallDir}images/voteSubmit.gif' width='52' height='18' border='0'></a>&nbsp;&nbsp;<a href='{$strInstallDir}Vote.asp?ID={$ID}&Action=Show' target='_blank'><img src='{$strInstallDir}images/voteView.gif' width='52' height='18' border='0'></a></div></form>")

			
            Do While Not rsVote.EOF
                If rsVote("VoteType") = "Single" Then
                    strVoteBody = ""
                    For i = 1 To 20
                        If Trim(rsVote("Select" & i) & "") = "" Then Exit For
                        strVoteBody = strVoteBody & "<input type='radio' name='VoteOption' value='" & i & "' style='border:0'>" & rsVote("Select" & i) & "<br>"
                    Next
                Else
                    strVoteBody = ""
                    For i = 1 To 20
                        If Trim(rsVote("Select" & i) & "") = "" Then Exit For
                        strVoteBody = strVoteBody & "<input type='checkbox' id='VoteForm"& j &"vote"& i &"' onClick=""CheckNum('"&rsVote("VoteNum")&"','VoteForm"&j&"','VoteForm"& j &"vote"& i &"')""  name='VoteOption' value='" & i & "' style='border:0'>" & rsVote("Select" & i) & "<br>"
                    Next
                End If
                strVote = strVote & Replace(Replace(Replace(Replace(Replace(Replace(strVoteContent, "{$lid}", j), "{$strInstallDir}", strInstallDir), "{$Title}", rsVote("Title")), "{$VoteBody}", strVoteBody), "{$VoteType}", rsVote("VoteType")), "{$ID}", rsVote("ID"))
                rsVote.MoveNext
                j = j + 1
            Loop
        End If
        rsVote.Close
        Set rsVote = Nothing
        PE_Cache.SetValue ChannelID & "_Site_Vote", strVote
    Else
        strVote = PE_Cache.GetValue(ChannelID & "_Site_Vote")
    End If	
    strVote = strVote & "<script language='javascript'>" & vbCrLf
    strVote = strVote & "function CheckNum(num,obj,item){" & vbCrLf
    strVote = strVote & " var count;" & vbCrLf
    strVote = strVote & " count=0" & vbCrLf
    strVote = strVote & " for(var i=0 ;i<document.getElementById(obj).elements.length;i++){" & vbCrLf
    strVote = strVote & " if(document.getElementById(obj).elements[i].checked==true){" & vbCrLf
    strVote = strVote & " count++;" & vbCrLf
    strVote = strVote & " if (count>num&&num!=0){" & vbCrLf
    strVote = strVote & "  alert('最多只能选择'+num+'个');" & vbCrLf
    strVote = strVote & "  document.getElementById(item).checked=false;" & vbCrLf
    strVote = strVote & "  }" & vbCrLf
    strVote = strVote & " }" & vbCrLf
    strVote = strVote & " }" & vbCrLf
    strVote = strVote & " }" & vbCrLf
    strVote = strVote & "</script>" & vbCrLf
    GetVote = strVote
End Function

'==================================================
'函数名：ShowAnnounce
'作  用：显示本站公告信息
'参  数：ShowType     ----显示方式，1为纵向，2为横向，3为输出DIV格式，4为输出RSS格式
'        AnnounceNum  ----最多显示多少条公告
'        ShowAuthor   ----是否显示公告发布人
'        ShowDate     ----是否显示公告发布日期
'        ContentLen   ----公告内容最多字符数，一个汉字=两个英文字符，为0时全部显示。
'==================================================
Function ShowAnnounce(ShowType, AnnounceNum, ShowAuthor, ShowDate, ContentLen)
    Dim sqlAnnounce, rsAnnounce, i, strAnnounce, AnnounceInfo, strContent
    If AnnounceNum > 0 And AnnounceNum <= 10 Then
        sqlAnnounce = "select top " & AnnounceNum
    Else
        sqlAnnounce = "select top 10"
    End If
    If ContentLen < 0 Then
        ContentLen = 100
    End If
    sqlAnnounce = sqlAnnounce & " * from PE_Announce where IsSelected=" & PE_True & " and (ChannelID=-1 or ChannelID=" & ChannelID & ") and (ShowType=0 or ShowType=1) and (OutTime=0 or OutTime>DateDiff(" & PE_DatePart_D & ",DateAndTime, " & PE_Now & ")) order by ID Desc"
    Set rsAnnounce = Conn.Execute(sqlAnnounce)
    If rsAnnounce.BOF And rsAnnounce.EOF Then
        If ShowType < 4 Then strAnnounce = XmlText("Site", "ShowAnnounce/AnnounceErr", "<p>&nbsp;&nbsp;没有公告</p>")
    Else
        If ShowType < 4 Then Set XMLDOM = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
        Dim strAnnounceBody1, strAnnounceBody2
        strAnnounceBody1 = XmlText("Site", "ShowAnnounce/AnnounceBody1", "&nbsp;&nbsp;&nbsp;&nbsp;<a class='AnnounceBody1' href='#' onclick=""javascript:window.open('{$strInstallDir}Announce.asp?ChannelID={$ChannelID}&ID={$ID}', 'newwindow', 'height=440, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='{$Content}'>{$title}{$AnnounceInfo}</a>")
        strAnnounceBody2 = XmlText("Site", "ShowAnnounce/AnnounceBody2", "&nbsp;&nbsp;&nbsp;&nbsp;<a class='AnnounceBody2' href='#' onclick=""javascript:window.open('{$strInstallDir}Announce.asp?ChannelID={$ChannelID}&ID={$ID}', 'newwindow', 'height=440, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='{$Content}'>{$title}{$AnnounceInfo}</a>")
        Do While Not rsAnnounce.EOF
            AnnounceInfo = ""
            Select Case ShowType
            Case 1
                If ShowAuthor = True Then
                    AnnounceInfo = AnnounceInfo & "<br><div align='right'>" & rsAnnounce("Author") & "&nbsp;&nbsp;"
                End If
                If ShowDate = True Then
                    If ShowAuthor = True Then
                        AnnounceInfo = AnnounceInfo & "<br>" & FormatDateTime(rsAnnounce("DateAndTime"), 1)
                    Else
                        AnnounceInfo = AnnounceInfo & "<br><div align='right'>" & FormatDateTime(rsAnnounce("DateAndTime"), 1)
                    End If
                End If
                If ShowAuthor = True Or ShowDate = True Then
                    AnnounceInfo = AnnounceInfo & "</div>"
                End If
                If ContentLen > 0 Then
                    strContent = GetSubStr(nohtml(PE_HtmlDecode(rsAnnounce("Content"))), ContentLen, False)
                Else
                    strContent = nohtml(PE_HtmlDecode(rsAnnounce("Content")))
                End If
                strAnnounce = strAnnounce & Replace(Replace(Replace(Replace(Replace(Replace(strAnnounceBody1, "{$strInstallDir}", strInstallDir), "{$ChannelID}", ChannelID), "{$ID}", rsAnnounce("id")), "{$Content}", strContent), "{$title}", rsAnnounce("title")), "{$AnnounceInfo}", AnnounceInfo)
                rsAnnounce.MoveNext
                i = i + 1
                If i < AnnounceNum Then strAnnounce = strAnnounce & "<hr>"
            Case 2
                If ShowAuthor = True Then
                    AnnounceInfo = AnnounceInfo & "&nbsp;&nbsp;[" & rsAnnounce("Author")
                End If
                If ShowDate = True Then
                    If ShowAuthor = True Then
                        AnnounceInfo = AnnounceInfo & "&nbsp;&nbsp;" & FormatDateTime(rsAnnounce("DateAndTime"), 1)
                    Else
                        AnnounceInfo = AnnounceInfo & "&nbsp;&nbsp;[" & FormatDateTime(rsAnnounce("DateAndTime"), 1)
                    End If
                End If
                If ShowAuthor = True Or ShowDate = True Then
                    AnnounceInfo = AnnounceInfo & "]"
                End If
                If ContentLen > 0 Then
                    strContent = GetSubStr(nohtml(PE_HtmlDecode(rsAnnounce("Content"))), ContentLen, False)
                Else
                    strContent = nohtml(PE_HtmlDecode(rsAnnounce("Content")))
                End If
                strAnnounce = strAnnounce & Replace(Replace(Replace(Replace(Replace(Replace(strAnnounceBody2, "{$strInstallDir}", strInstallDir), "{$ChannelID}", ChannelID), "{$ID}", rsAnnounce("id")), "{$Content}", strContent), "{$title}", rsAnnounce("title")), "{$AnnounceInfo}", AnnounceInfo)
                strAnnounce = strAnnounce & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                rsAnnounce.MoveNext
            Case 3
                strAnnounce = strAnnounce & "<div class=""announce"">"
                strAnnounce = strAnnounce & "<div class=""announce_title""><a class=""announcetitle"" href=""#"" onclick=""javascript:window.open('" & strInstallDir & "Announce.asp?ChannelID=" & ChannelID & "&ID=" & rsAnnounce("id") & "', 'newwindow', 'height=440, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"">" & rsAnnounce("title") & "</a></div>"
                If ShowAuthor = True Then
                    strAnnounce = strAnnounce & "<div class=""announce_author"">" & rsAnnounce("Author") & "</div>"
                End If
                If ShowDate = True Then
                    strAnnounce = strAnnounce & "<div class=""announce_time"">" & FormatDateTime(rsAnnounce("DateAndTime"), 1) & "</div>"
                End If
                If ContentLen > 0 Then
                    strAnnounce = strAnnounce & ("<div class=""announce_content"">" & GetSubStr(nohtml(PE_HtmlDecode(rsAnnounce("Content"))), ContentLen, False) & "</div>")
                Else
                    strAnnounce = strAnnounce & ("<div class=""announce_content"">" & nohtml(PE_HtmlDecode(rsAnnounce("Content"))) & "</div>")
                End If
                strAnnounce = strAnnounce & "</div>"
                rsAnnounce.MoveNext
            Case 4
                XMLDOM.appendChild (XMLDOM.createElement("item"))
                Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("title"))
                Node.Text = xml_nohtml(rsAnnounce("title"))
                Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("link"))
                Node.Text = "http://" & Trim(Request.ServerVariables("HTTP_HOST")) & "/Announce.asp?ChannelID=" & ChannelID & "&ID=" & rsAnnounce("id")
                Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("description"))
                If ContentLen > 0 Then
                    Node.Text = GetSubStr(xml_nohtml(rsAnnounce("Content")), ContentLen, False)
                Else
                    Node.Text = xml_nohtml(rsAnnounce("Content"))
                End If
                Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("author"))
                Node.Text = xml_nohtml(rsAnnounce("Author"))
                Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("category"))
                Node.Text = "网站公告"
                Set Node = XMLDOM.documentElement.appendChild(XMLDOM.createElement("pubDate"))
                Node.Text = rsAnnounce("DateAndTime")
                strAnnounce = strAnnounce & XMLDOM.documentElement.xml
                rsAnnounce.MoveNext
            End Select
        Loop
    End If
    rsAnnounce.Close
    Set rsAnnounce = Nothing
    ShowAnnounce = strAnnounce
End Function

'==================================================
'函数名：ShowFriendSite
'作  用：显示友情链接站点
'参  数：LinkType  ----链接方式，1为LOGO链接，2为文字链接
'       SiteNum   ----最多显示多少个站点
'       Cols      ----分几列显示
'       ShowType  ----显示方式。1为向上滚动，2为横向列表，3为下拉列表框，4为输出DIV格式
'       KindID    ----所属类别
'       SpecialID ----所属专题
'       TDWidth   ----所用表格宽度
'       IsOnlyElite ----是否只显示推荐
'       IsElitFirst ----是否推荐优先
'       OrderType ----排序方式：1---友情链接ID升序；
'                               2---友情链接ID降序；
'                               3---排序ID升序；
'                               4---排序ID降序；
'                               5---网站评分等级升序；
'                               6---网站评分等级降序；
'==================================================
Function ShowFriendSite(LinkType, SiteNum, Cols, ShowType, KindID, SpecialID, TDWidth, IsOnlyElite, IsElitFirst, OrderType)
    Dim sqlLink, rsLink, SiteCount, i, j, strLink, strLogo
    Dim LinkSiteUrl
    LinkType = PE_CLng(LinkType)
    If LinkType <> 1 And LinkType <> 2 Then
        LinkType = 1
    End If
    SiteNum = PE_Clng(SiteNum)
    If SiteNum <= 0 Or SiteNum > 100 Then
        SiteNum = 10
    End If
    If Cols <= 0 Or Cols > 20 Then
        Cols = 10
    End If
    If ShowType = 1 Then
        strLink = strLink & "<div id=rolllink style=overflow:hidden;height:100;width:100><div id=rolllink1>"
    ElseIf ShowType = 3 Then
        strLink = strLink & "<select name='FriendSite' onchange=""if(this.options[this.selectedIndex].value!=''){window.open(this.options[this.selectedIndex].value,'_blank');}""><option value=''>" & XmlText("Site", "ShowFriendSite/option", "友情文字链接站点") & "</option>"
    End If
    If ShowType = 1 Or ShowType = 2 Then
        strLink = strLink & XmlText("Site", "ShowFriendSite/Showtable", "<table width='100%' cellSpacing='5'><tr align='center' class='tdbg'>")
    End If
    If IsValidID(KindID) = True Then
        KindID = Replace(Replace(KindID, "|", ","), " ", "")
    Else
        KindID = 0
    End If
    If IsValidID(SpecialID) = True Then
        SpecialID = Replace(Replace(SpecialID, "|", ","), " ", "")
    Else
        SpecialID = 0
    End If
    If PE_CLng(TDWidth) <= 0 Then
        TDWidth = 88
    End If
    sqlLink = "select top " & SiteNum & " * from PE_FriendSite where Passed=" & PE_True & " and LinkType=" & LinkType
    If KindID <> 0 Then
        sqlLink = sqlLink & " and KindID in (" & KindID & ")"
    End If
    If SpecialID <> 0 Then
        sqlLink = sqlLink & " and SpecialID in (" & SpecialID & ")"
    End If
    If IsOnlyElite = True Then
        sqlLink = sqlLink & " and Elite=" & PE_True
    End If
    'sqlLink = sqlLink & " order by Elite " & PE_OrderType & ",ID desc"
    sqlLink = sqlLink & " order by "
    If IsElitFirst = True Then
        sqlLink = sqlLink & "Elite " & PE_OrderType & ","
    End If
    Select Case OrderType
    Case 1
        sqlLink = sqlLink & "ID asc"
    Case 2
        sqlLink = sqlLink & "ID desc"
    Case 3
        sqlLink = sqlLink & "OrderID asc,ID desc"
    Case 4
        sqlLink = sqlLink & "OrderID desc,ID desc"
    Case 5
        sqlLink = sqlLink & "Stars asc,ID desc"
    Case 6
        sqlLink = sqlLink & "Stars desc,ID desc"
    Case Else
        sqlLink = sqlLink & "OrderID asc,ID desc"
    End Select

    Dim strGetFriendSite
    strGetFriendSite = XmlText("Site", "ShowFriendSite/GetFriendSite", "点击申请")
    Set rsLink = Conn.Execute(sqlLink)
    If rsLink.BOF And rsLink.EOF Then
        If ShowType = 1 Or ShowType = 2 Then
            For i = 1 To SiteNum
                strLink = strLink & "<td><a class='LinkFriendSite' href='" & strInstallDir & "FriendSite/FriendSiteReg.asp' target='_blank'>"
                If LinkType = 1 Then
                    strLink = strLink & "<img src='" & strInstallDir & "images/nologo.gif' width='88' height='31' border='0' alt='" & strGetFriendSite & "'>"
                Else
                    strLink = strLink & strGetFriendSite
                End If
                strLink = strLink & "</a></td>"
                If i Mod Cols = 0 And i < SiteNum Then
                    strLink = strLink & "</tr><tr align='center' class='tdbg'>"
                End If
            Next
        ElseIf ShowType = 4 Then
            For i = 1 To SiteNum
                strLink = strLink & "<div class=""linkfriendsite""><a href=""" & strInstallDir & "FriendSite/FriendSiteReg.asp"" target=""_blank"">"
                If LinkType = 1 Then
                    strLink = strLink & "<img src=""" & strInstallDir & "images/nologo.gif"" width=""88"" height=""31"" border=""0"" alt=""" & strGetFriendSite & """>"
                Else
                    strLink = strLink & strGetFriendSite
                End If
                strLink = strLink & "</a></div>"
            Next
        End If
    Else
        i = 1
        Dim strFriendSiteTitle
        strFriendSiteTitle = XmlText("Site", "ShowFriendSite/FriendSiteTitle", "<a class='LinkFriendSite' href='{$LinkSiteUrl}' target='_blank' title='网站名称：{$SiteName}{$br}网站地址：{$SiteUrl}{$br}网站简介：{$SiteIntro}'>{$SiteShow}</a>")
        Do While Not rsLink.EOF
            If EnableCountFriendSiteHits = True Then
                LinkSiteUrl = strInstallDir & "FriendSite/FriendSiteUrl.asp?ID=" & rsLink("ID")
            Else
                LinkSiteUrl = rsLink("SiteUrl")
            End If
            Select Case ShowType
            Case 1, 2
                If LinkType = 1 Then
                    If rsLink("LogoUrl") = "" Or rsLink("LogoUrl") = "http://" Then
                        strLogo = "<img src='" & strInstallDir & "images/nologo.gif' width='88' height='31' border='0'>"
                    Else
                        If LCase(Right(rsLink("LogoUrl"), 3)) = "swf" Then
                            strLogo = "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='88' height='31'><param name='movie' value='" & rsLink("LogoUrl") & "'><param name='quality' value='high'><embed src='" & rsLink("LogoUrl") & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='88' height='31'></embed></object>"
                        Else
                            strLogo = "<img src='" & rsLink("LogoUrl") & "' width='88' height='31' border='0'>"
                        End If
                    End If
                    strLink = strLink & "<td width='" & TDWidth & "'>"
                    strLink = strLink & Replace(Replace(Replace(Replace(Replace(Replace(strFriendSiteTitle, "{$LinkSiteUrl}", LinkSiteUrl), "{$SiteName}", rsLink("SiteName")), "{$SiteUrl}", rsLink("SiteUrl")), "{$SiteIntro}", rsLink("SiteIntro")), "{$SiteShow}", strLogo), "{$br}", vbCrLf)
                    strLink = strLink & "</td>"
                Else
                    strLink = strLink & "<td width='" & TDWidth & "'>"
                    strLink = strLink & Replace(Replace(Replace(Replace(Replace(Replace(strFriendSiteTitle, "{$LinkSiteUrl}", LinkSiteUrl), "{$SiteName}", rsLink("SiteName")), "{$SiteUrl}", rsLink("SiteUrl")), "{$SiteIntro}", rsLink("SiteIntro")), "{$SiteShow}", rsLink("SiteName")), "{$br}", vbCrLf)
                    strLink = strLink & "</td>"
                End If
                If i Mod Cols = 0 And i < SiteNum Then
                    strLink = strLink & "</tr><tr align='center' class='tdbg'>"
                End If
            Case 3
                strLink = strLink & "<option value='" & LinkSiteUrl & "'>" & rsLink("SiteName") & "</option>"
            Case 4
                If LinkType = 1 Then
                    If rsLink("LogoUrl") = "" Or rsLink("LogoUrl") = "http://" Then
                        strLogo = "<img src=""" & strInstallDir & "images/nologo.gif"" width=""88"" height=""31"" border=""0"">"
                    Else
                        If LCase(Right(rsLink("LogoUrl"), 3)) = "swf" Then
                            strLogo = "<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0"" width=""88"" height=""31""><param name=""movie"" value=""" & rsLink("LogoUrl") & """><param name=""quality"" value=""high""><embed src=""" & rsLink("LogoUrl") & """ pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" width=""88"" height=""31""></embed></object>"
                        Else
                            strLogo = "<img src=""" & rsLink("LogoUrl") & """ width=""88"" height=""31"" border=""0"">"
                        End If
                    End If
                    strLink = strLink & "<div class=""linkfriendsite"">"
                    strLink = strLink & Replace(Replace(Replace(Replace(Replace(Replace(strFriendSiteTitle, "{$LinkSiteUrl}", LinkSiteUrl), "{$SiteName}", rsLink("SiteName")), "{$SiteUrl}", rsLink("SiteUrl")), "{$SiteIntro}", rsLink("SiteIntro")), "{$SiteShow}", strLogo), "{$br}", vbCrLf)
                    strLink = strLink & "</div>"
                Else
                    strLink = strLink & "<div class=""linkfriendsite"">"
                    strLink = strLink & Replace(Replace(Replace(Replace(Replace(Replace(strFriendSiteTitle, "{$LinkSiteUrl}", LinkSiteUrl), "{$SiteName}", rsLink("SiteName")), "{$SiteUrl}", rsLink("SiteUrl")), "{$SiteIntro}", rsLink("SiteIntro")), "{$SiteShow}", rsLink("SiteName")), "{$br}", vbCrLf)
                    strLink = strLink & "</div>"
                End If
            End Select
            rsLink.MoveNext
            i = i + 1
        Loop
        If i < SiteNum And (ShowType = 1 Or ShowType = 2 Or ShowType = 4) Then
            For j = i To SiteNum
                If ShowType = 4 Then
                    If LinkType = 1 Then
                        strLink = strLink & "<div class=""linkfriendsite""><a href=""" & strInstallDir & "FriendSite/FriendSiteReg.asp"" target=""_blank""><img src=""" & strInstallDir & "images/nologo.gif"" width=""88"" height=""31"" border=""0"" alt=""" & strGetFriendSite & """></a></div>"
                    Else
                        strLink = strLink & "<div class=""linkfriendsite""><a href=""" & strInstallDir & "FriendSite/FriendSiteReg.asp"" target=""_blank"">" & strGetFriendSite & "</a></div>"
                    End If
                Else
                    If LinkType = 1 Then
                        strLink = strLink & "<td width='" & TDWidth & "'><a class='LinkFriendSite' href='" & strInstallDir & "FriendSite/FriendSiteReg.asp' target='_blank'><img src='" & strInstallDir & "images/nologo.gif' width='88' height='31' border='0' alt='" & strGetFriendSite & "'></a></td>"
                    Else
                        strLink = strLink & "<td width='" & TDWidth & "'><a class='LinkFriendSite' href='" & strInstallDir & "FriendSite/FriendSiteReg.asp' target='_blank'>" & strGetFriendSite & "</a></td>"
                    End If
                    If j Mod Cols = 0 And j < SiteNum Then
                        strLink = strLink & "</tr><tr align='center' class='tdbg'>"
                    End If
                End If
            Next
        End If
    End If
    Select Case ShowType
    Case 1
        strLink = strLink & "</tr></table>"
        strLink = strLink & "</div><div id=rolllink2></div></div>"
        strLink = strLink & RollFriendSite()
    Case 2
        strLink = strLink & "</tr></table>"
    Case 3
        strLink = strLink & "</select>"
    End Select
    rsLink.Close
    Set rsLink = Nothing
    ShowFriendSite = strLink
End Function

'==================================================
'函数名：RollFriendSite
'作  用：滚动显示友情链接站点
'参  数：无
'==================================================
Function RollFriendSite()
    Dim strTemp
    strTemp = "<script>" & vbCrLf
    strTemp = strTemp & "var rollspeed=30" & vbCrLf
    strTemp = strTemp & "rolllink2.innerHTML=rolllink1.innerHTML" & vbCrLf
    strTemp = strTemp & "function Marquee(){" & vbCrLf
    strTemp = strTemp & "if(rolllink2.offsetTop-rolllink.scrollTop<=0)" & vbCrLf
    strTemp = strTemp & "rolllink.scrollTop-=rolllink1.offsetHeight" & vbCrLf
    strTemp = strTemp & "else{" & vbCrLf
    strTemp = strTemp & "rolllink.scrollTop++" & vbCrLf
    strTemp = strTemp & "}" & vbCrLf
    strTemp = strTemp & "}" & vbCrLf
    strTemp = strTemp & "var MyMar=setInterval(Marquee,rollspeed)" & vbCrLf
    strTemp = strTemp & "rolllink.onmouseover=function() {clearInterval(MyMar)}" & vbCrLf
    strTemp = strTemp & "rolllink.onmouseout=function() {MyMar=setInterval(Marquee,rollspeed)}" & vbCrLf
    strTemp = strTemp & "</script>" & vbCrLf
    RollFriendSite = strTemp
End Function

Function ShowGoodSite(SiteNum)
    Dim sqlLink, rsLink, SiteCount, i, j, strLink, strLogo
    Dim LinkSiteUrl
    
    If SiteNum <= 0 Or SiteNum > 100 Then
        SiteNum = 10
    End If
    strLink = strLink & "<table width='100%' cellSpacing='5'>"
    Dim strFriendSiteTitle, strGetFriendSite
    strFriendSiteTitle = strFriendSiteTitle
    strGetFriendSite = XmlText("Site", "ShowFriendSite/GetFriendSite", "点击申请")
    sqlLink = "select top " & SiteNum & " * from PE_FriendSite where Passed=" & PE_True & " and Elite=" & PE_True & " order by OrderID asc"
    Set rsLink = Conn.Execute(sqlLink)
    If rsLink.BOF And rsLink.EOF Then
        For i = 1 To SiteNum
            strLink = strLink & "<tr align='center'><td><a class='LinkFriendSite' href='" & strInstallDir & "FriendSite/FriendSiteReg.asp' target='_blank'><img src='" & strInstallDir & "images/nologo.gif' width='88' height='31' border='0' alt='" & strGetFriendSite & "'></a></td></tr>"
        Next
    Else
        Do While Not rsLink.EOF
            If EnableCountFriendSiteHits = True Then
                LinkSiteUrl = strInstallDir & "FriendSite/FriendSiteUrl.asp?ID=" & rsLink("ID")
            Else
                LinkSiteUrl = rsLink("SiteUrl")
            End If
            strLink = strLink & "<tr align='center'><td>"
            If rsLink("LogoUrl") = "" Or rsLink("LogoUrl") = "http://" Then
                strLogo = "<img src='" & strInstallDir & "images/nologo.gif' width='88' height='31' border='0'>"
            Else
                strLogo = "<img src='" & rsLink("LogoUrl") & "' width='88' height='31' border='0'>"
            End If
            strLink = strLink & Replace(Replace(Replace(Replace(Replace(Replace(strFriendSiteTitle, "{$LinkSiteUrl}", LinkSiteUrl), "{$SiteName}", rsLink("SiteName")), "{$SiteUrl}", rsLink("SiteUrl")), "{$SiteIntro}", rsLink("SiteIntro")), "{$SiteShow}", strLogo), "{$br}", vbCrLf)
            strLink = strLink & "</td></tr>"
            rsLink.MoveNext
        Loop
    End If
    strLink = strLink & "</table>"
    rsLink.Close
    Set rsLink = Nothing
    ShowGoodSite = strLink
End Function

'==================================================
'函数名：GetTopUser
'作  用：显示用户排行，按已发表的文章数排序，若相等，再按注册先后顺序排序
'参  数：UserNum-------显示的用户个数
'        OrderType ----排序方式，1为按发表信息数降序，2为按发表信息数升序，3按用户ID降序，4为按用户ID升序，5为按点数降序，6为按点数升序，7为资金降序，8为按资金升序
'        ShowNum ---- 是否显示名次，True为显示，False为不显示
'        ShowPassedItems ---- 是否显示发表信息数，True为显示，False为不显示
'        ShowPoints ----- 是否显示点数，True为显示，False为不显示
'        ShowMoney ---- 是否显示资金数，True为显示，False为不显示
'        strMore ---- “更多”的字符，如果为空，则不显示“更多”字样
'        ShowType ---- 输出模式，1为输出表格 2为输出DIV格式
'==================================================
Function GetTopUser(UserNum, OrderType, ShowNum, ShowPassedItems, ShowPoint, ShowMoney, strMore, ShowType)
    Dim sqlTopUser, rsTopUser, i, strTopUser
    If UserNum <= 0 Or UserNum > 100 Or IsNull(UserNum) Or UserNum = "" Then UserNum = 10
    If IsNull(ShowType) Then ShowType = 1
    sqlTopUser = "select top " & PE_CLng(UserNum) & " UserID,UserName,PassedItems,UserPoint,Balance from PE_User order by"
    Select Case OrderType
    Case 1
        sqlTopUser = sqlTopUser & " PassedItems desc,UserID asc"
    Case 2
        sqlTopUser = sqlTopUser & " PassedItems asc,UserID asc"
    Case 3
        sqlTopUser = sqlTopUser & " UserID desc"
    Case 4
        sqlTopUser = sqlTopUser & " UserID asc"
    Case 5
        sqlTopUser = sqlTopUser & " UserPoint desc,UserID asc"
    Case 6
        sqlTopUser = sqlTopUser & " UserPoint asc,UserID asc"
    Case 7
        sqlTopUser = sqlTopUser & " Balance desc,UserID asc"
    Case 8
        sqlTopUser = sqlTopUser & " Balance asc,UserID asc"
    Case Else
        sqlTopUser = sqlTopUser & " UserID desc"
    End Select
    Set rsTopUser = Server.CreateObject("adodb.recordset")
    rsTopUser.Open sqlTopUser, Conn, 1, 1
    If rsTopUser.BOF And rsTopUser.EOF Then
        strTopUser = XmlText("Site", "Errmsg/UserErr", "没有任何会员")
    Else
        If ShowType = 1 Then
            strTopUser = "<table width='98%' border='0' cellspacing='0' cellpadding='0'><tr>" & XmlText("Site", "GetTopUser/Table1", "<td width='30' align='center'>名次</td><td align='center'>用户名</td><td align='center'>文章数</td>") & "</tr>"
            For i = 1 To rsTopUser.RecordCount
                strTopUser = strTopUser & "<tr>"
                If ShowNum = True Then
                    strTopUser = strTopUser & "<td  width='30' align='center'>" & CStr(i) & "</td>"
                End If
                strTopUser = strTopUser & "<td align='center'><a class='LinkTopUser' href='" & strInstallDir & "ShowUser.asp?ChannelID=" & ChannelID & "&UserID=" & rsTopUser("UserID") & "'>" & rsTopUser("UserName") & "</a></td>"
                If ShowPassedItems = True Then
                    strTopUser = strTopUser & "<td align='center'>" & rsTopUser("PassedItems") & "</td>"
                End If
                If ShowPoint = True Then
                    strTopUser = strTopUser & "<td align='center'>" & rsTopUser("UserPoint") & "</td>"
                End If
                If ShowMoney = True Then
                    strTopUser = strTopUser & "<td align='center'>" & rsTopUser("Balance") & "</td>"
                End If
                strTopUser = strTopUser & "</tr>"
                rsTopUser.MoveNext
            Next
            strTopUser = strTopUser & "</table>"
            If Trim(strMore) <> "" Then
                strTopUser = strTopUser & "<div align='right'><a class='LinkTopUser' href='" & strInstallDir & "UserList.asp?ChannelID=" & ChannelID & "'>" & strMore & "</a>&nbsp;&nbsp;</div>"
            End If
        Else
            For i = 1 To rsTopUser.RecordCount
                strTopUser = "<div class=""userlist"">"
                If ShowNum = True Then
                    strTopUser = strTopUser & "<div class=""userlist_num"">" & CStr(i) & "</div>"
                End If
                strTopUser = strTopUser & "<div calss=""userlist_name""><a class=""LinkTopUser"" href=""" & strInstallDir & "ShowUser.asp?ChannelID=" & ChannelID & "&UserID=" & rsTopUser("UserID") & """>" & rsTopUser("UserName") & "</a></div>"
                If ShowPassedItems = True Then
                    strTopUser = strTopUser & "<div calss=""userlist_passitem"">" & rsTopUser("PassedItems") & "</div>"
                End If
                If ShowPoint = True Then
                    strTopUser = strTopUser & "<div calss=""userlist_point"">" & rsTopUser("UserPoint") & "</div>"
                End If
                If ShowMoney = True Then
                    strTopUser = strTopUser & "<div calss=""userlist_money"">" & rsTopUser("Balance") & "</div>"
                End If
                strTopUser = strTopUser & "</div>"
                rsTopUser.MoveNext
            Next
            If Trim(strMore) <> "" Then
                strTopUser = strTopUser & "<div class=""userlist_more""><a class=""LinkTopUser"" href=""" & strInstallDir & "UserList.asp?ChannelID=" & ChannelID & """>" & strMore & "</a></div>"
            End If
        End If
    End If
    Set rsTopUser = Nothing
    GetTopUser = strTopUser
End Function

'==================================================
'函数名：GetBlogList
'作  用：显示Blog排行

'        ListNum ----显示列表数
'        ClassID ----栏目ID
'        ShowType ---- 显示类型
'        Elite ----- 是否推荐
'        Hot ---- 是否热点
'        imgH ---- 图片宽度
'        imgW ---- 图片高度
'        IntroNum ---- 简介文字数
'        Imgstat ---- 　图片于文字的位置
'        ShowMore ---- 显示更多
'        ColNum ---- 显示列数
'==================================================
Function GetBlogList(ListNum, ClassID, ShowType, Elite, Hot, imgH, imgW, IntroNum, Imgstat, ShowMore, ColNum)
    Dim sqlBlog, rsBlog, strBlog, imgtmp, iCol, BlogLinkUrl, spacename
    If PE_CLng(ListNum) = 0 Then ListNum = 5
    sqlBlog = "select top " & ListNum & " A.UserID,A.ClassID,A.Name,A.Intro,A.Photo,A.Passed,A.Type,A.isElite,A.Hits,A.onTop,A.LastUseTime,C.UserName from PE_Space A left join PE_User C on A.UserID=C.UserID Where A.Passed=" & PE_True & " and A.Type=1"
    If PE_CLng(ClassID) > 0 Then sqlBlog = sqlBlog & " and A.ClassID=" & PE_CLng(ClassID)
    If Elite = True Then sqlBlog = sqlBlog & " and A.isElite=" & PE_True
    If Hot = True Then sqlBlog = sqlBlog & " and A.Hits>" & XmlText("ShowSource", "Space/HitsOfHot", "100")
    sqlBlog = sqlBlog & " order by A.onTop " & PE_OrderType & ",A.LastUseTime Desc"
    Set rsBlog = Server.CreateObject("ADODB.Recordset")
    rsBlog.Open sqlBlog, Conn, 1, 1
    If rsBlog.BOF And rsBlog.EOF Then
        strBlog = XmlText("ShowSource", "Space/NotFound", "尚未添加聚合空间!")
    Else
        iCol = 1

        Select Case PE_CLng(ShowType)
        Case 0
            strBlog = strBlog & "<table Class='spaceList'><tr>"
            Do While Not rsBlog.EOF
               spacename = Replace(LCase(rsBlog("UserName") & rsBlog(0)), ".", "")
                BlogLinkUrl = strInstallDir & "Space/" & spacename & "/"
                strBlog = strBlog & ("<td Class='spaceList_intro'><a class='LinkspaceList' href='" & BlogLinkUrl & "'")
                If Trim(rsBlog("Intro") & "") <> "" Then
                    strBlog = strBlog & (" alt='" & GetSubStr(nohtml(PE_HtmlDecode(rsBlog("Intro"))), IntroNum, False) & "'")
                End If
                strBlog = strBlog & (" target='_blank'>" & rsBlog("Name") & "</a></td>")
                rsBlog.MoveNext
                If Not rsBlog.EOF Then
                    If iCol >= Int(ColNum) Then
                        strBlog = strBlog & "</tr><tr>"
                        iCol = 1
                    Else
                        iCol = iCol + 1
                    End If
                End If
            Loop
            strBlog = strBlog & "</tr></table>"
        Case 1
            strBlog = strBlog & "<table Class='spaceList'><tr>"
            Do While Not rsBlog.EOF
                spacename = Replace(LCase(rsBlog("UserName") & rsBlog(0)), ".", "")
                BlogLinkUrl = strInstallDir & "Space/" & spacename & "/"
                If Trim(rsBlog("Photo") & "") = "" Then
                    imgtmp = strInstallDir & "Space/default.gif"
                Else
                    imgtmp = rsBlog("Photo")
                End If
                Select Case Imgstat
                Case 0, 1
                    strBlog = strBlog & ("<td Class='spaceList_image'><a class='LinkspaceList' href='" & BlogLinkUrl & "' target='_blank'><img src='" & imgtmp & "' height='" & imgH & "' width='" & imgW & "' border=0 alt='" & nohtml(rsBlog("Name")) & "'></a></td><td Class='spaceList_intro'>")
                    If Trim(rsBlog("Intro") & "") = "" Or Imgstat = 0 Then
                        strBlog = strBlog & GetSubStr(rsBlog("Name"), IntroNum, False) & "</td>"
                    Else
                        strBlog = strBlog & GetSubStr(nohtml(PE_HtmlDecode(rsBlog("Intro"))), IntroNum, False) & "</td>"
                    End If
                Case 2, 3
                    strBlog = strBlog & ("<td Class='spaceList_image'><a class='LinkspaceList' href='" & BlogLinkUrl & "' target='_blank'><img src='" & imgtmp & "' height='" & imgH & "' width='" & imgW & "' border=0 alt='" & nohtml(rsBlog("Name")) & "'><br>")
                    If Trim(rsBlog("Intro") & "") = "" Or Imgstat = 2 Then
                        strBlog = strBlog & GetSubStr(rsBlog("Name"), IntroNum, False) & "</td>"
                    Else
                        strBlog = strBlog & GetSubStr(nohtml(PE_HtmlDecode(rsBlog("Intro"))), IntroNum, False) & "</a></td>"
                    End If
                Case 4, 5
                    strBlog = strBlog & "<td Class='spaceList_intro'>"
                    If Trim(rsBlog("Intro") & "") = "" Or Imgstat = 4 Then
                        strBlog = strBlog & GetSubStr(rsBlog("Name"), IntroNum, False)
                    Else
                        strBlog = strBlog & GetSubStr(nohtml(PE_HtmlDecode(rsBlog("Intro"))), IntroNum, False)
                    End If
                    strBlog = strBlog & ("</td><td Class='spaceList_image'><a class='LinkspaceList' href='" & BlogLinkUrl & "' target='_blank'><img src='" & imgtmp & "' height='" & imgH & "' width='" & imgW & "' border=0 alt='" & nohtml(rsBlog("Name")) & "'></a></td>")
                Case 6, 7
                    strBlog = strBlog & ("<td Class='spaceList_image'><a class='LinkspaceList' href='" & BlogLinkUrl & "' target='_blank'>")
                    If Trim(rsBlog("Intro") & "") = "" Or Imgstat = 6 Then
                        strBlog = strBlog & GetSubStr(rsBlog("Name"), IntroNum, False)
                    Else
                        strBlog = strBlog & GetSubStr(nohtml(PE_HtmlDecode(rsBlog("Intro"))), IntroNum, False)
                    End If
                    strBlog = strBlog & ("<br><img src='" & imgtmp & "' height='" & imgH & "' width='" & imgW & "' border=0 alt='" & nohtml(rsBlog("Name")) & "'></a></td>")
                End Select
                rsBlog.MoveNext
                If Not rsBlog.EOF Then
                    If iCol >= Int(ColNum) Then
                        strBlog = strBlog & "</tr><tr>"
                        iCol = 1
                    Else
                        iCol = iCol + 1
                    End If
                End If
            Loop
            strBlog = strBlog & "</tr></table>"
        Case 2
            strBlog = strBlog & "<table Class='spaceList'><tr>"
            Do While Not rsBlog.EOF
                spacename = Replace(LCase(rsBlog("UserName") & rsBlog(0)), ".", "")
                BlogLinkUrl = strInstallDir & "Space/" & spacename & "/"
                If Trim(rsBlog("Photo") & "") = "" Then
                    strBlog = strBlog & ("<td Class='spaceList_image'><a class='LinkspaceList' href='" & BlogLinkUrl & "' target='_blank'><img src='" & strInstallDir & "Space/default.gif' height='" & imgH & "' width='" & imgW & "' border=0 alt='" & nohtml(rsBlog("Name")) & "'></a></td>")
                Else
                    strBlog = strBlog & ("<td Class='spaceList_image'><a class='LinkspaceList' href='" & BlogLinkUrl & "' target='_blank'><img src='" & rsBlog("Photo") & "' height='" & imgH & "' width='" & imgW & "' border=0 alt='" & nohtml(rsBlog("Name")) & "'></a></td>")
                End If
                rsBlog.MoveNext
                If Not rsBlog.EOF Then
                    If iCol >= Int(ColNum) Then
                        strBlog = strBlog & "</tr><tr>"
                        iCol = 1
                    Else
                        iCol = iCol + 1
                    End If
                End If
            Loop
            strBlog = strBlog & "</tr></table>"
        Case 3
            Dim iChannelLink: iChannelLink = XmlText("BaseText", "ChannelLink", "&nbsp;|&nbsp;")
            strBlog = strBlog & iChannelLink
            Do While Not rsBlog.EOF
                spacename = Replace(LCase(rsBlog("UserName") & rsBlog(0)), ".", "")
                BlogLinkUrl = strInstallDir & "Space/" & spacename & "/"
                strBlog = strBlog & ("<a class='LinkspaceList' href='" & BlogLinkUrl & "'")
                If Trim(rsBlog("Intro") & "") <> "" Then
                    strBlog = strBlog & (" alt='" & GetSubStr(nohtml(PE_HtmlDecode(rsBlog("Intro"))), IntroNum, False) & "'")
                End If
                strBlog = strBlog & ("' target='_blank'>" & rsBlog("Name") & "</a>" & iChannelLink)
                rsBlog.MoveNext
                If Not rsBlog.EOF Then
                    If iCol >= Int(ColNum) Then
                        strBlog = strBlog & ("<br>" & iChannelLink)
                        iCol = 1
                    Else
                        iCol = iCol + 1
                    End If
                End If
            Loop
        Case 4
            Do While Not rsBlog.EOF
                spacename = Replace(LCase(rsBlog("UserName") & rsBlog(0)), ".", "")
                BlogLinkUrl = strInstallDir & "Space/" & spacename & "/"
                strBlog = strBlog & ("<li><a class='LinkspaceList' href='" & BlogLinkUrl & "'")
                If Trim(rsBlog("Intro") & "") <> "" Then
                    strBlog = strBlog & (" alt='" & GetSubStr(nohtml(PE_HtmlDecode(rsBlog("Intro"))), IntroNum, False) & "'")
                End If
                strBlog = strBlog & (" target='_blank'>" & rsBlog("Name") & "</a></li>")
                rsBlog.MoveNext
            Loop
        Case 5
            Do While Not rsBlog.EOF
                spacename = Replace(LCase(rsBlog("UserName") & rsBlog(0)), ".", "")
                BlogLinkUrl = strInstallDir & "Space/" & spacename & "/"
                strBlog = strBlog & ("<div class=""showspacelist""><a href=""" & BlogLinkUrl & """")
                If Trim(rsBlog("Intro") & "") <> "" Then
                    strBlog = strBlog & (" alt=""" & GetSubStr(nohtml(PE_HtmlDecode(rsBlog("Intro"))), IntroNum, False) & """")
                End If
                strBlog = strBlog & (" target=""_blank"">" & rsBlog("Name") & "</a></div>")
                rsBlog.MoveNext
            Loop
        End Select
        If ShowMore <> "none" Then
            strBlog = strBlog & ("<div class=""showspacelist_more""><a class=""LinkspaceList"" href=""" & strInstallDir & "Space/"" target='_blank'>" & ShowMore & "</a></div>")
        End If
    End If
    rsBlog.Close
    Set rsBlog = Nothing
    GetBlogList = strBlog
End Function

'==================================================
'函数名：GetAuthorList
'作  用：显示作者排行
'==================================================
Function GetAuthorList(iChannelID, AuthorNum, AuthorType, ShowType, imgH, imgW, IntroNum, Imgstat, ShowMore, ColNum)
    Dim sqlAuthor, rsAuthor, strAuthor, imgtmp, iCol
    Dim tempChannelID
    If PE_CLng(AuthorNum) = 0 Then AuthorNum = 5
    sqlAuthor = "select top " & PE_CLng(AuthorNum) & " * from PE_Author Where Passed=" & PE_True
    If PE_CLng(iChannelID) > 0 Then sqlAuthor = sqlAuthor & " and ChannelID=" & PE_CLng(iChannelID)
    If PE_CLng(AuthorType) > 0 Then sqlAuthor = sqlAuthor & " and AuthorType=" & PE_CLng(AuthorType)
    If PE_CLng(ShowType) > 10 Then sqlAuthor = sqlAuthor & " and isElite=" & PE_True
    sqlAuthor = sqlAuthor & " order by onTop " & PE_OrderType & ",LastUseTime Desc"
    Set rsAuthor = Server.CreateObject("ADODB.Recordset")
    rsAuthor.Open sqlAuthor, Conn, 1, 1
    If rsAuthor.BOF And rsAuthor.EOF Then
        strAuthor = XmlText("Site", "Errmsg/AuthorErr", "没有任何作者")
    Else
        iCol = 1
        If ShowType >= 10 Then ShowType = ShowType - 10
        Select Case ShowType
        Case 0
            strAuthor = strAuthor & "<table Class='AuthorList'><tr>"
            Do While Not rsAuthor.EOF
                If PE_CLng(rsAuthor("ChannelID")) = 0 Then
                    tempChannelID = ChannelID
                Else
                    tempChannelID = PE_CLng(rsAuthor("ChannelID"))                
                End If													    				
                strAuthor = strAuthor & ("<td Class='AuthorList_intro'><a class='LinkAuthorList' href='" & strInstallDir & "ShowAuthor.asp?ChannelID=" & tempChannelID & "&AuthorName=" & rsAuthor("AuthorName") & "' alt='" & GetSubStr(nohtml(PE_HtmlDecode(rsAuthor("Intro"))), IntroNum, False) & "'>" & rsAuthor("AuthorName") & "</a></td>")
                rsAuthor.MoveNext
                If Not rsAuthor.EOF Then
                    If iCol >= Int(ColNum) Then
                        strAuthor = strAuthor & "</tr><tr>"
                        iCol = 1
                    Else
                        iCol = iCol + 1
                    End If
                End If
            Loop
            strAuthor = strAuthor & "</tr></table>"
        Case 1		
            strAuthor = strAuthor & "<table Class='AuthorList'><tr>"
            Do While Not rsAuthor.EOF
                If PE_CLng(rsAuthor("ChannelID")) = 0 Then
                    tempChannelID = ChannelID
                Else
                    tempChannelID = PE_CLng(rsAuthor("ChannelID"))                
                End If				
                If Trim(rsAuthor("Photo") & "") = "" Then
                    imgtmp = strInstallDir & "authorpic/default.gif"
                Else
                    imgtmp = rsAuthor("Photo")
                End If
				
                Select Case Imgstat
                Case 0
                    strAuthor = strAuthor & ("<td Class='AuthorList_image'><a class='LinkAuthorList' href='" & strInstallDir & "ShowAuthor.asp?ChannelID=" & tempChannelID & "&AuthorName=" & rsAuthor("AuthorName") & "'><img src='" & imgtmp & "' height='" & imgH & "' width='" & imgW & "' border=0></a></td><td Class='AuthorList_intro'>")
                    If Trim(rsAuthor("Intro") & "") = "" Then
                        strAuthor = strAuthor & GetSubStr(rsAuthor("AuthorName"), IntroNum, False) & "</td>"
                    Else
                        strAuthor = strAuthor & GetSubStr(nohtml(PE_HtmlDecode(rsAuthor("Intro"))), IntroNum, False) & "</td>"
                    End If
                Case 1
                    strAuthor = strAuthor & ("<td Class='AuthorList_image'><a class='LinkAuthorList' href='" & strInstallDir & "ShowAuthor.asp?ChannelID=" & tempChannelID & "&AuthorName=" & rsAuthor("AuthorName") & "'><img src='" & imgtmp & "' height='" & imgH & "' width='" & imgW & "' border=0><br>")
                    If Trim(rsAuthor("Intro") & "") = "" Then
                        strAuthor = strAuthor & GetSubStr(rsAuthor("AuthorName"), IntroNum, False) & "</td>"
                    Else
                        strAuthor = strAuthor & GetSubStr(nohtml(PE_HtmlDecode(rsAuthor("Intro"))), IntroNum, False) & "</a></td>"
                    End If
                Case 2
                    strAuthor = strAuthor & "<td Class='AuthorList_intro'>"
                    If Trim(rsAuthor("Intro") & "") = "" Then
                        strAuthor = strAuthor & GetSubStr(rsAuthor("AuthorName"), IntroNum, False)
                    Else
                        strAuthor = strAuthor & GetSubStr(nohtml(PE_HtmlDecode(rsAuthor("Intro"))), IntroNum, False)
                    End If
                    strAuthor = strAuthor & ("</td><td Class='AuthorList_image'><a class='LinkAuthorList' href='" & strInstallDir & "ShowAuthor.asp?ChannelID=" & tempChannelID & "&AuthorName=" & rsAuthor("AuthorName") & "'><img src='" & imgtmp & "' height='" & imgH & "' width='" & imgW & "' border=0></a></td>")
                Case Else
                    strAuthor = strAuthor & ("<td Class='AuthorList_image'><a class='LinkAuthorList' href='" & strInstallDir & "ShowAuthor.asp?ChannelID=" & tempChannelID & "&AuthorName=" & rsAuthor("AuthorName") & "'>")
                    If Trim(rsAuthor("Intro") & "") = "" Then
                        strAuthor = strAuthor & GetSubStr(rsAuthor("AuthorName"), IntroNum, False)
                    Else
                        strAuthor = strAuthor & GetSubStr(nohtml(PE_HtmlDecode(rsAuthor("Intro"))), IntroNum, False)
                    End If
                    strAuthor = strAuthor & ("<br><img src='" & imgtmp & "' height='" & imgH & "' width='" & imgW & "' border=0></a></td>")
                End Select
                rsAuthor.MoveNext
                If Not rsAuthor.EOF Then
                    If iCol >= Int(ColNum) Then
                        strAuthor = strAuthor & "</tr><tr>"
                        iCol = 1
                    Else
                        iCol = iCol + 1
                    End If
                End If
            Loop
            strAuthor = strAuthor & "</tr></table>"
        Case 2
            strAuthor = strAuthor & "<table Class='AuthorList'><tr>"
            Do While Not rsAuthor.EOF
                If PE_CLng(rsAuthor("ChannelID")) = 0 Then
                    tempChannelID = ChannelID
                Else
                    tempChannelID = PE_CLng(rsAuthor("ChannelID"))                
                End If			
                If Trim(rsAuthor("Photo") & "") = "" Then
                    strAuthor = strAuthor & ("<td Class='AuthorList_image'><a class='LinkAuthorList' href='" & strInstallDir & "ShowAuthor.asp?ChannelID=" & tempChannelID & "&AuthorName=" & rsAuthor("AuthorName") & "'><img src='" & strInstallDir & "AuthorPic/default.gif' height='" & imgH & "' width='" & imgW & "' border=0></a></td>")
                Else
                    strAuthor = strAuthor & ("<td Class='AuthorList_image'><a class='LinkAuthorList' href='" & strInstallDir & "ShowAuthor.asp?ChannelID=" & tempChannelID & "&AuthorName=" & rsAuthor("AuthorName") & "'><img src='" & rsAuthor("Photo") & "' height='" & imgH & "' width='" & imgW & "' border=0></a></td>")
                End If
                rsAuthor.MoveNext
                If Not rsAuthor.EOF Then
                    If iCol >= Int(ColNum) Then
                        strAuthor = strAuthor & "</tr><tr>"
                        iCol = 1
                    Else
                        iCol = iCol + 1
                    End If
                End If
            Loop
            strAuthor = strAuthor & "</tr></table>"
        Case 3
            Dim iChannelLink: iChannelLink = XmlText("BaseText", "ChannelLink", "&nbsp;|&nbsp;")
            strAuthor = strAuthor & iChannelLink
            Do While Not rsAuthor.EOF
                If PE_CLng(rsAuthor("ChannelID")) = 0 Then
                    tempChannelID = ChannelID
                Else
                    tempChannelID = PE_CLng(rsAuthor("ChannelID"))                
                End If				
                strAuthor = strAuthor & ("<a class='LinkAuthorList' href='" & strInstallDir & "ShowAuthor.asp?ChannelID=" & tempChannelID & "&AuthorName=" & rsAuthor("AuthorName") & "' alt='" & GetSubStr(nohtml(PE_HtmlDecode(rsAuthor("Intro"))), IntroNum, False) & "' target='_self'>" & rsAuthor("AuthorName") & "</a>" & iChannelLink)
                rsAuthor.MoveNext
                If Not rsAuthor.EOF Then
                    If iCol >= Int(ColNum) Then
                        strAuthor = strAuthor & ("<br>" & iChannelLink)
                        iCol = 1
                    Else
                        iCol = iCol + 1
                    End If
                End If
            Loop
        Case 4
            Do While Not rsAuthor.EOF
                If PE_CLng(rsAuthor("ChannelID")) = 0 Then
                    tempChannelID = ChannelID
                Else
                    tempChannelID = PE_CLng(rsAuthor("ChannelID"))                
                End If				
                strAuthor = strAuthor & ("<li><a class='LinkAuthorList' href='" & strInstallDir & "ShowAuthor.asp?ChannelID=" & tempChannelID & "&AuthorName=" & rsAuthor("AuthorName") & "' alt='" & GetSubStr(nohtml(PE_HtmlDecode(rsAuthor("Intro"))), IntroNum, False) & "'>" & rsAuthor("AuthorName") & "</a></li>")
                rsAuthor.MoveNext
            Loop
        Case 5
            Do While Not rsAuthor.EOF
                If PE_CLng(rsAuthor("ChannelID")) = 0 Then
                    tempChannelID = ChannelID
                Else
                    tempChannelID = PE_CLng(rsAuthor("ChannelID"))                
                End If			
                strAuthor = strAuthor & ("<div class=""showauthorlist""><a href=""" & strInstallDir & "ShowAuthor.asp?ChannelID=" & tempChannelID & "&AuthorName=" & rsAuthor("AuthorName") & """ alt=""" & GetSubStr(nohtml(PE_HtmlDecode(rsAuthor("Intro"))), IntroNum, False) & """>" & rsAuthor("AuthorName") & "</a></div>")
                rsAuthor.MoveNext
            Loop
        End Select
        If ShowMore <> "none" Then
            If ShowType < 5 Then
                strAuthor = strAuthor & ("<div align='right'><a class='LinkAuthorList' href='" & strInstallDir & "AuthorList.asp?ChannelID=" & ChannelID & "'>" & ShowMore & "</a></div>")
            Else
                strAuthor = strAuthor & ("<div class=""showauthorlist_more""><a href=""" & strInstallDir & "AuthorList.asp?ChannelID=" & ChannelID & """>" & ShowMore & "</a></div>")
            End If
        End If
    End If
    rsAuthor.Close
    Set rsAuthor = Nothing
    GetAuthorList = strAuthor
End Function

'==================================================
'函数名：GetProducerList
'作  用：显示厂商排行
'==================================================
Function GetProducerList(iChannelID, ProducerNum, ProducerType, ShowType, imgH, imgW, IntroNum, Imgstat, ShowMore, ColNum)
    Dim sqlProducer, rsProducer, strProducer, imgtmp, iCol
    If PE_CLng(ProducerNum) = 0 Then ProducerNum = 5
    sqlProducer = "select top " & PE_CLng(ProducerNum) & " * from PE_Producer Where Passed=" & PE_True
    If PE_CLng(iChannelID) > 0 Then sqlProducer = sqlProducer & " and ChannelID=" & PE_CLng(iChannelID)
    If PE_CLng(ProducerType) > 0 Then sqlProducer = sqlProducer & " and ProducerType=" & PE_CLng(ProducerType)
    If PE_CLng(ShowType) > 10 Then sqlProducer = sqlProducer & " and isElite=" & PE_True
    sqlProducer = sqlProducer & " order by onTop " & PE_OrderType & ",LastUseTime Desc"
    Set rsProducer = Server.CreateObject("ADODB.Recordset")
    rsProducer.Open sqlProducer, Conn, 1, 1
    If rsProducer.BOF And rsProducer.EOF Then
        strProducer = "没有任何厂商"
    Else
        iCol = 1
        If ShowType >= 10 Then ShowType = ShowType - 10
        Select Case ShowType
        Case 0
            strProducer = strProducer & "<table Class='ProducerList'><tr>"
            Do While Not rsProducer.EOF
                strProducer = strProducer & ("<td Class='ProducerList_intro'><a class='LinkProducerList' href='" & strInstallDir & "Shop/ShowProducer.asp?ChannelID=" & rsProducer("ChannelID") & "&ProducerName=" & rsProducer("ProducerName") & "' alt='" & GetSubStr(nohtml(PE_HtmlDecode(rsProducer("ProducerIntro"))), IntroNum, False) & "'>" & rsProducer("ProducerName") & "</a></td>")
                rsProducer.MoveNext
                If Not rsProducer.EOF Then
                    If iCol >= Int(ColNum) Then
                        strProducer = strProducer & "</tr><tr>"
                        iCol = 1
                    Else
                        iCol = iCol + 1
                    End If
                End If
            Loop
            strProducer = strProducer & "</tr></table>"
        Case 1
            strProducer = strProducer & "<table Class='ProducerList'><tr>"
            Do While Not rsProducer.EOF
                If Trim(rsProducer("ProducerPhoto") & "") = "" Then
                    imgtmp = strInstallDir & "Shop/ProducerPic/default.gif"
                Else
                    imgtmp = rsProducer("ProducerPhoto")
                End If
                Select Case Imgstat
                Case 0
                    strProducer = strProducer & ("<td Class='ProducerList_image'><a class='LinkProducerList' href='" & strInstallDir & "Shop/ShowProducer.asp?ChannelID=" & rsProducer("ChannelID") & "&ProducerName=" & rsProducer("ProducerName") & "'><img src='" & imgtmp & "' height='" & imgH & "' width='" & imgW & "' border=0></a></td><td Class='ProducerList_intro'>" & GetSubStr(nohtml(PE_HtmlDecode(rsProducer("ProducerIntro"))), IntroNum, False) & "</td>")
                Case 1
                    strProducer = strProducer & ("<td Class='ProducerList_image'><a class='LinkProducerList' href='" & strInstallDir & "Shop/ShowProducer.asp?ChannelID=" & rsProducer("ChannelID") & "&ProducerName=" & rsProducer("ProducerName") & "'><img src='" & imgtmp & "' height='" & imgH & "' width='" & imgW & "' border=0><br>" & GetSubStr(nohtml(PE_HtmlDecode(rsProducer("ProducerIntro"))), IntroNum, False) & "</a></td>")
                Case 2
                    strProducer = strProducer & ("<td Class='ProducerList_intro'>" & GetSubStr(nohtml(PE_HtmlDecode(rsProducer("ProducerIntro"))), IntroNum, False) & "</td><td Class='ProducerList_image'><a class='LinkProducerList' href='" & strInstallDir & "Shop/ShowProducer.asp?ChannelID=" & rsProducer("ChannelID") & "&ProducerName=" & rsProducer("ProducerName") & "'><img src='" & imgtmp & "' height='" & imgH & "' width='" & imgW & "' border=0></a></td>")
                Case Else
                    strProducer = strProducer & ("<td Class='ProducerList_image'><a class='LinkProducerList' href='" & strInstallDir & "Shop/ShowProducer.asp?ChannelID=" & rsProducer("ChannelID") & "&ProducerName=" & rsProducer("ProducerName") & "'>" & GetSubStr(nohtml(PE_HtmlDecode(rsProducer("ProducerIntro"))), IntroNum, False) & "<br><img src='" & imgtmp & "' height='" & imgH & "' width='" & imgW & "' border=0></a></td>")
                End Select
                rsProducer.MoveNext
                If Not rsProducer.EOF Then
                    If iCol >= Int(ColNum) Then
                        strProducer = strProducer & "</tr><tr>"
                        iCol = 1
                    Else
                        iCol = iCol + 1
                    End If
                End If
            Loop
            strProducer = strProducer & "</tr></table>"
        Case 2
            strProducer = strProducer & "<table Class='ProducerList'><tr>"
            Do While Not rsProducer.EOF
                If Trim(rsProducer("ProducerPhoto") & "") = "" Then
                    strProducer = strProducer & ("<td Class='ProducerList_image'><a class='LinkProducerList' href='" & strInstallDir & "Shop/ShowProducer.asp?ChannelID=" & rsProducer("ChannelID") & "&ProducerName=" & rsProducer("ProducerName") & "'><img src='" & strInstallDir & "Shop/ProducerPic/default.gif' height='" & imgH & "' width='" & imgW & "' border=0></a></td>")
                Else
                    strProducer = strProducer & ("<td Class='ProducerList_image'><a class='LinkProducerList' href='" & strInstallDir & "Shop/ShowProducer.asp?ChannelID=" & rsProducer("ChannelID") & "&ProducerName=" & rsProducer("ProducerName") & "'><img src='" & rsProducer("Photo") & "' height='" & imgH & "' width='" & imgW & "' border=0></a></td>")
                End If
                rsProducer.MoveNext
                If Not rsProducer.EOF Then
                    If iCol >= Int(ColNum) Then
                        strProducer = strProducer & "</tr><tr>"
                        iCol = 1
                    Else
                        iCol = iCol + 1
                    End If
                End If
            Loop
            strProducer = strProducer & "</tr></table>"
        Case 3
            Dim iChannelLink: iChannelLink = XmlText("BaseText", "ChannelLink", "&nbsp;|&nbsp;")
            
            strProducer = strProducer & iChannelLink
            Do While Not rsProducer.EOF
                strProducer = strProducer & ("<a class='LinkProducerList' href='" & strInstallDir & "Shop/ShowProducer.asp?ChannelID=" & rsProducer("ChannelID") & "&ProducerName=" & rsProducer("ProducerName") & "' alt='" & GetSubStr(nohtml(PE_HtmlDecode(rsProducer("ProducerIntro"))), IntroNum, False) & "' target='_self'>" & rsProducer("ProducerName") & "</a>" & iChannelLink)
                rsProducer.MoveNext
                If Not rsProducer.EOF Then
                    If iCol >= Int(ColNum) Then
                        strProducer = strProducer & ("<br>" & iChannelLink)
                        iCol = 1
                    Else
                        iCol = iCol + 1
                    End If
                End If
            Loop
        Case 4
            Do While Not rsProducer.EOF
                strProducer = strProducer & ("<li><a class='LinkProducerList' href='" & strInstallDir & "Shop/ShowProducer.asp?ChannelID=" & rsProducer("ChannelID") & "&ProducerName=" & rsProducer("ProducerName") & "' alt='" & GetSubStr(nohtml(PE_HtmlDecode(rsProducer("ProducerIntro"))), IntroNum, False) & "'>" & rsProducer("ProducerName") & "</a></li>")
                rsProducer.MoveNext
            Loop
        Case 5
            Do While Not rsProducer.EOF
                strProducer = strProducer & ("<div class=""showproducerlist""><a href=""" & strInstallDir & "Shop/ShowProducer.asp?ChannelID=" & rsProducer("ChannelID") & "&ProducerName=" & rsProducer("ProducerName") & """ alt=""" & GetSubStr(nohtml(PE_HtmlDecode(rsProducer("ProducerIntro"))), IntroNum, False) & """>" & rsProducer("ProducerName") & "</a></div>")
                rsProducer.MoveNext
            Loop
        End Select
        If ShowMore <> "none" Then
            If ShowType < 5 Then
                strProducer = strProducer & ("<div align='right'><a class='LinkProducerList' href='" & strInstallDir & "Shop/ShowProducer.asp?Action=List&ChannelID=" & ChannelID & "'>" & ShowMore & "</a></div>")
            Else
                strProducer = strProducer & ("<div class=""showproducerlist_more""><a href=""" & strInstallDir & "Shop/ShowProducer.asp?Action=List&ChannelID=" & ChannelID & """>" & ShowMore & "</a></div>")
            End If
        End If
    End If
    rsProducer.Close
    Set rsProducer = Nothing
    GetProducerList = strProducer
End Function

'==================================================
'过程名：PopAnnouceWindow
'作  用：弹出公告窗口
'参  数：Width-------弹出窗口宽度
'        Height------弹出窗口高度
'==================================================
Function PopAnnouceWindow(Width, Height)
    Dim popCount, rsAnnounce, strPop, strJS
    Set rsAnnounce = Conn.Execute("select count(ID) from PE_Announce where (ChannelID=-1 or ChannelID=" & ChannelID & ") and  IsSelected=" & PE_True & " and (ShowType=0 or ShowType=2) and (OutTime=0 or OutTime>DateDiff(" & PE_DatePart_D & ",DateAndTime, " & PE_Now & "))")
    If IsNull(rsAnnounce(0)) Then
        popCount = 0
    Else
        popCount = rsAnnounce(0)
    End If
    
    If popCount > 0 Then
        strJS = Replace(Replace(Replace(Replace(XmlText("Site", "PopAnnouceWindow/PopCode", "window.open ('{$strInstallDir}Announce.asp?ChannelID={$ChannelID}', 'newwindow', 'height={$Height}, width={$Width}, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"), "{$strInstallDir}", strInstallDir), "{$ChannelID}", ChannelID), "{$Height}", Height), "{$Width}", Width)
        If AnnounceCookieTime <= 0 Then
            strPop = "<script language='JavaScript'>" & vbCrLf
            strPop = strPop & strJS & vbCrLf
            strPop = strPop & "</script>" & vbCrLf
        Else
            strPop = "<script LANGUAGE='JavaScript'>" & vbCrLf
            strPop = strPop & "var PAW_CookieHour = " & AnnounceCookieTime & ";" & vbCrLf
            strPop = strPop & "PAW_Show_Check()" & vbCrLf
            strPop = strPop & "function PAW_Show_Check() {" & vbCrLf
            strPop = strPop & "  if (PAW_CookieCheck()) return false;" & vbCrLf
            strPop = strPop & "  " & strJS & vbCrLf
            strPop = strPop & "}" & vbCrLf
            strPop = strPop & "function PAW_CookieCheck() {" & vbCrLf
            strPop = strPop & "  if (!PAW_CookieHour) return false;" & vbCrLf
            strPop = strPop & "  var Now = new Date();" & vbCrLf
            strPop = strPop & "  var strToday = String(Now.getYear()) + String(Now.getMonth() + 1) + String(Now.getDate());" & vbCrLf
            strPop = strPop & "  var PAW_Cookie = 'PopAnnouceWindow'" & vbCrLf
            strPop = strPop & "  if (PAW_GetCookie(PAW_Cookie) == strToday)" & vbCrLf
            strPop = strPop & "    return true;" & vbCrLf
            strPop = strPop & "  else {" & vbCrLf
            strPop = strPop & "    Now.setTime(Now.getTime() + (parseFloat(typeof(PAW_CookieHour) == 'undefined' ? 0 : parseFloat(PAW_CookieHour)) * 60 * 60 * 1000));" & vbCrLf
            strPop = strPop & "    PAW_SetCookie(PAW_Cookie, strToday, Now);" & vbCrLf
            strPop = strPop & "    return false;" & vbCrLf
            strPop = strPop & "  }" & vbCrLf
            strPop = strPop & "}" & vbCrLf
            strPop = strPop & "function PAW_GetCookie(name) {" & vbCrLf
            strPop = strPop & "  var arg = name + '=';" & vbCrLf
            strPop = strPop & "  var alen = arg.length;" & vbCrLf
            strPop = strPop & "  var clen = document.cookie.length;" & vbCrLf
            strPop = strPop & "  var i = 0;" & vbCrLf
            strPop = strPop & "  while (i < clen) {" & vbCrLf
            strPop = strPop & "    var j = i + alen;" & vbCrLf
            strPop = strPop & "    if (document.cookie.substring(i, j) == arg)" & vbCrLf
            strPop = strPop & "      return PAW_GetCookieVal (j);" & vbCrLf
            strPop = strPop & "    i = document.cookie.indexOf(' ', i) + 1;" & vbCrLf
            strPop = strPop & "    if (i == 0) break; " & vbCrLf
            strPop = strPop & "  }" & vbCrLf
            strPop = strPop & "  return null;" & vbCrLf
            strPop = strPop & "}" & vbCrLf
            strPop = strPop & "function PAW_GetCookieVal(offset) {" & vbCrLf
            strPop = strPop & "  var endstr = document.cookie.indexOf (';', offset);" & vbCrLf
            strPop = strPop & "  if (endstr == -1)" & vbCrLf
            strPop = strPop & "    endstr = document.cookie.length;" & vbCrLf
            strPop = strPop & "  return unescape(document.cookie.substring(offset, endstr));" & vbCrLf
            strPop = strPop & "}" & vbCrLf
            strPop = strPop & "function PAW_SetCookie(name, value)" & vbCrLf
            strPop = strPop & "{" & vbCrLf
            strPop = strPop & "  var argv = PAW_SetCookie.arguments;" & vbCrLf
            strPop = strPop & "  var argc = PAW_SetCookie.arguments.length;" & vbCrLf
            strPop = strPop & "  var expires = (argc > 2) ? argv[2]: null;" & vbCrLf
            strPop = strPop & "  var path = (argc > 3) ? argv[3]: null;" & vbCrLf
            strPop = strPop & "  var domain = (argc > 4) ? argv[4]: null;" & vbCrLf
            strPop = strPop & "  var secure = (argc > 5) ? argv[5]: false;" & vbCrLf
            strPop = strPop & "  document.cookie = name + '=' + escape(value) + ((expires == null) ? '' : ('; expires=' + expires.toGMTString())) + ((path == null) ? '' : ('; path=' + path)) + ((domain == null) ? '' : ('; domain=' + domain)) + ((secure == true) ? '; secure' : '');" & vbCrLf
            strPop = strPop & "}" & vbCrLf
            strPop = strPop & "</script>"
        End If
    Else
        strPop = ""
    End If
    PopAnnouceWindow = strPop
End Function

Function GetSkin_CSS(tSkinID)
    If tSkinID = 0 Then
        GetSkin_CSS = "<link href='" & strInstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>"
    Else
        GetSkin_CSS = "<link href='" & strInstallDir & "Skin/Skin" & tSkinID & ".css' rel='stylesheet' type='text/css'>"
    End If
End Function

Function GetTemplate(iChannelID, TemplateType, TemplateID)
    Dim sqlTemplate, rsTemplate
    If IsNull(TemplateID) Or Trim(TemplateID) = "" Then TemplateID = 0
    
    If TemplateID = 0 Then
        Set rsTemplate = Conn.Execute("select TemplateContent from PE_Template where ChannelID=" & PE_CLng(iChannelID) & " and TemplateType=" & PE_CLng(TemplateType) & " and IsDefault=" & PE_True & "")
    Else
        Set rsTemplate = Conn.Execute("select TemplateContent from PE_Template where ChannelID=" & PE_CLng(iChannelID) & " and TemplateType=" & PE_CLng(TemplateType) & " and TemplateID=" & PE_CLng(TemplateID) & "")
    End If
    If rsTemplate.BOF And rsTemplate.EOF Then
        GetTemplate = XmlText("BaseText", "TemplateErr", "找不到模板")
    Else
        GetTemplate = rsTemplate(0)
    End If
    rsTemplate.Close
    Set rsTemplate = Nothing
End Function

'=================================================
'函数名：GetClass_Option
'作  用：显示栏目下拉列表的HTML代码
'参  数：无
'返回值：栏目下拉列表的HTML代码
'=================================================
Function GetClass_Option(CurrentID)
    Dim rsClass, sqlClass, strTemp, tmpDepth, i
    Dim arrShowLine(20)
    For i = 0 To UBound(arrShowLine)
        arrShowLine(i) = False
    Next
    sqlClass = "Select ClassID,ClassName,ClassType,Depth,NextID from PE_Class where ChannelID=" & ChannelID & " order by RootID,OrderID"
    Set rsClass = Conn.Execute(sqlClass)
    If rsClass.BOF And rsClass.EOF Then
        strTemp = "<option value=''>" & XmlText("Site", "GetClass_Option/option", "请先添加栏目") & "</option>"
    Else
        strTemp = ""
        Do While Not rsClass.EOF
            tmpDepth = rsClass(3)
            If rsClass(4) > 0 Then
                arrShowLine(tmpDepth) = True
            Else
                arrShowLine(tmpDepth) = False
            End If
            strTemp = strTemp & "<option value='" & rsClass(0) & "'"
            If CurrentID > 0 And rsClass(0) = CurrentID Then
                 strTemp = strTemp & " selected"
            End If
            strTemp = strTemp & ">"
            
            If tmpDepth > 0 Then
                For i = 1 To tmpDepth
                    strTemp = strTemp & "&nbsp;&nbsp;"
                    If i = tmpDepth Then
                        If rsClass(4) > 0 Then
                            strTemp = strTemp & "├&nbsp;"
                        Else
                            strTemp = strTemp & "└&nbsp;"
                        End If
                    Else
                        If arrShowLine(i) = True Then
                            strTemp = strTemp & "│"
                        Else
                            strTemp = strTemp & "&nbsp;"
                        End If
                    End If
                Next
            End If
            strTemp = strTemp & rsClass(1)
            If rsClass(2) = 2 Then
                strTemp = strTemp & XmlText("Site", "GetClass_Option/outside", "(外)")
            End If
            strTemp = strTemp & "</option>"
            rsClass.MoveNext
        Loop
    End If
    rsClass.Close
    Set rsClass = Nothing

    GetClass_Option = strTemp
End Function

Function ReplaceSql(TempSql) 
    Dim temp
    RegEx.Pattern = "\s(ChannelID)\s*=\s*0|\s(Elite)\s*=\s*0|\s(OnTop)\s*=\s*0|\s(ClassID)\s*=\s*0|\s(ChannelID)\s*(in)\s*\(\s*0\s*\)|\s(ClassID)\s*(in)\s*\(\s*0\s*\)" 
    temp = RegEx.replace(TempSql," 1=1 ") 
	temp = Replace(temp,"{$PE_False}",PE_False)
	temp = Replace(temp,"{$PE_True}",PE_True)
	temp = Replace(temp, "{$GetUserName}", GetUserName())
	ReplaceSql = temp
End Function 

Sub ReplaceCommonLabel()
    Dim rsLabel, nonLabel, LabelNum, Looptotalnum, PageNum, reFlashTime
    Dim DyTemp, LoopTemp, loopTempMatch, loopTempMatch2, InfoTemp, InfoTempMatch, FieldTemp, FieldArry, FieldTempText
    Dim rsLabelRe, InfoID, TempSql
    Dim arrAreaCode, arrAreaCode2, i
    Dim CaiTemp, rsArea
    Dim strChannel, strLogo, strBanner, strTopUser, strAuthorList, strProducerList, strFriendSite, strAnnounce, strPopAnnouce
    Dim strTemp, arrTemp
    Dim Match2, MatchesInfo, Match3, Matches3
    strHtml = PE_Replace(strHtml, "%7B", "{")
    strHtml = PE_Replace(strHtml, "%7D", "}")
    strHtml = PE_Replace(strHtml, "{$InstallDir}{$ChannelDir}", ChannelUrl)

    '以下这段代码放在最前面，是用于在自定义动态函数标签中可以解析个别标签
    '{$InstallDir}{$ChannelDir}的替换一定要放在单个{$ChannelDir}的前面
    If ChannelID > 0 Then
        strHtml = PE_Replace(strHtml, "{$InstallDir}{$ChannelDir}", ChannelUrl)
        strHtml = PE_Replace(strHtml, "{$ChannelID}", ChannelID)
        strHtml = PE_Replace(strHtml, "{$ChannelDir}", ChannelDir)
        strHtml = PE_Replace(strHtml, "{$ChannelUrl}", ChannelUrl)
        strHtml = PE_Replace(strHtml, "{$ChannelPicUrl}", ChannelPicUrl)
    End If
    
    '解决编辑器过滤为标签值为空问题
    regEx.Pattern = "(\s)+(value|title|src|href)(\s)*\=(\s)*\{\$(.[^\<\{]*)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        strTemp = Replace(Trim(Match.Value), "{$", """{$")
        strTemp = Replace(strTemp, "}", "}""")
        strHtml = Replace(strHtml, Match.Value, " " & strTemp)
    Next
    
    Looptotalnum = 10
    If InStr(strHtml, "{$MY_") > 0 Then
        LabelNum = 1
    Else
        LabelNum = 0
    End If
    Do While LabelNum > 0
        nonLabel = True
        '替换静态标签
        Set rsLabel = Conn.Execute("select LabelName,LabelContent from PE_Label where LabelType=0 order by Priority asc,LabelID asc")
        Do While Not rsLabel.EOF
            If InStr(strHtml, "{$" & rsLabel("LabelName")) > 0 Then
                If InStr(strHtml, "{$" & rsLabel("LabelName") & "}") > 0 Then
                    strHtml = Replace(strHtml, "{$" & rsLabel("LabelName") & "}", rsLabel("LabelContent"))
                Else
                    regEx.Pattern = "\{\$" & rsLabel("LabelName") & "\((.*?)\)\}"
                    Set Matches = regEx.Execute(strHtml)
                    For Each Match In Matches
                        InfoTemp = rsLabel("LabelContent")
                        arrTemp = Split(Match.SubMatches(0), ",")
                        For i = 0 To UBound(arrTemp)
                            InfoTemp = Replace(InfoTemp, "{input(" & i & ")}", arrTemp(i))
                        Next
                        strHtml = Replace(strHtml, Match.Value, InfoTemp)
                    Next
                End If
                nonLabel = False
            End If
            rsLabel.MoveNext
        Loop

        '替换动态标签
        Set rsLabel = Conn.Execute("select LabelID,LabelName,LabelType,PageNum,LabelIntro,LabelContent,reFlashTime from PE_Label where LabelType=1 or LabelType=3 order by Priority asc,LabelID asc")
        Do While Not rsLabel.EOF
        PageNum = rsLabel("PageNum")
        reFlashTime = rsLabel("reFlashTime")

        If InStr(strHtml, "{$" & rsLabel("LabelName")) > 0 Then
            Dim temptime, temptimetext, dyfoot
            LoopTemp = rsLabel("LabelContent")
            LoopTemp = Replace(LoopTemp,"{$InstallDir}{$Field(0,GetUrl","{$Field(0,GetUrl")		
            LoopTemp = Replace(Replace(Replace(Replace(LoopTemp, "{$Now}", Now()), "{$NowDay}", Day(Now())), "{$NowMonth}", Month(Now())), "{$NowYear}", Year(Now()))
            regEx.Pattern = "\{Loop\}([\s\S]*?)\{\/Loop\}"
            Set Matches = regEx.Execute(LoopTemp)
            For Each Match In Matches
                loopTempMatch = Match.Value
            Next
            LoopTemp = regEx.Replace(LoopTemp, "{$SqlReplaceText}")
            loopTempMatch = Replace(Replace(loopTempMatch, "{loop}", ""), "{/loop}", "")

            Select Case rsLabel("LabelType")
            Case 1 '标准动态标签的处理过程
                TempSql = Replace(Replace(Replace(Replace(Replace(rsLabel("LabelIntro"), "{$Now}", Now()), "{$NowDay}", Day(Now())), "{$NowMonth}", Month(Now())), "{$NowYear}", Year(Now())),"{$PE_DatePart_D}", PE_DatePart_D)
                TempSql = Replace(Replace(Replace(Replace(TempSql, "{$PE_True}", PE_True), "{$PE_False}", PE_False), "{$PE_Now}", PE_Now), "{$PE_OrderType}", PE_OrderType)
                If InStr(strHtml, "{$" & rsLabel("LabelName") & "}") > 0 Then

                If PE_Cache.CacheIsEmpty("{$" & rsLabel("LabelName") & "}") Then

                    '开始循环处理内容
                    InfoID = 0
                    On Error Resume Next
                    Set rsLabelRe = Server.CreateObject("adodb.recordset")
                    TempSql = Replacesql(TempSql)
                    rsLabelRe.Open TempSql, Conn, 1, 1
                    If Err Then
                        Err.Clear
                        DyTemp = "SQL查询错误"
                    Else
                        totalPut = rsLabelRe.RecordCount
                        If rsLabelRe.BOF And rsLabelRe.EOF Then
                            DyTemp = "尚无数据"
                        Else
                            loopTempMatch = Replace(Replace(Replace(Replace(loopTempMatch, "{$Now}", Now()), "{$NowDay}", Day(Now())), "{$NowMonth}", Month(Now())), "{$NowYear}", Year(Now()))
                            Do While Not rsLabelRe.EOF
                            regEx.Pattern = "\{Infobegin\}([\s\S]*?)\{Infoend\}"
                            Set Matches = regEx.Execute(loopTempMatch)
                     
                            If Matches.Count = 0 Then
                                rsLabelRe.MoveNext
                            Else
                                For Each Match In Matches
                                    If Not rsLabelRe.EOF Then
                                        InfoTemp = Match.Value
                                        InfoTempMatch = Replace(Replace(InfoTemp, "{Infobegin}", ""), "{Infoend}", "") '得到最终的单一字段内容
                                        regEx.Pattern = "\{\$Field\((.*?)\)\}"
                                        Set MatchesInfo = regEx.Execute(InfoTempMatch)
                                        For Each Match2 In MatchesInfo
                                            FieldTemp = Match2.Value
                                            FieldArry = Split(Match2.SubMatches(0), ",")
                                            If UBound(FieldArry) > 1 Then '参数正确,进行处理
                                                Select Case FieldArry(1)
                                                Case "Text" '按文本方式输出内容
                                                    If rsLabelRe(PE_CLng(FieldArry(0))) = "" Or IsNull(rsLabelRe(PE_CLng(FieldArry(0)))) Then
                                                        FieldTempText = ""
                                                    Else
                                                        If FieldArry(2) = 0 Then
                                                            Select Case FieldArry(3)
                                                            Case 1
                                                                FieldTempText = Replace(rsLabelRe(PE_CLng(FieldArry(0))), "<", "&lt;")
                                                            Case 2
                                                                FieldTempText = nohtml(rsLabelRe(PE_CLng(FieldArry(0))))
                                                            Case Else
                                                                FieldTempText = rsLabelRe(PE_CLng(FieldArry(0)))
                                                            End Select
                                                        Else
                                                            Select Case FieldArry(3)
                                                            Case 1
                                                                If FieldArry(4) = 0 Then
                                                                    FieldTempText = GetSubStr(Replace(rsLabelRe(PE_CLng(FieldArry(0))), "<", "&lt;"), PE_CLng(FieldArry(2)), True)
                                                                Else
                                                                    FieldTempText = GetSubStr(Replace(rsLabelRe(PE_CLng(FieldArry(0))), "<", "&lt;"), PE_CLng(FieldArry(2)), False)
                                                                End If
                                                            Case 2
                                                                If FieldArry(4) = 0 Then
                                                                    FieldTempText = GetSubStr(nohtml(rsLabelRe(PE_CLng(FieldArry(0)))), PE_CLng(FieldArry(2)), True)
																	FieldTempText = Replace(FieldTempText, Chr(10), "<br>")
                                                                Else
                                                                    FieldTempText = GetSubStr(nohtml(rsLabelRe(PE_CLng(FieldArry(0)))), PE_CLng(FieldArry(2)), False)
																	FieldTempText = Replace(FieldTempText, Chr(10), "<br>")
                                                                End If
                                                            Case Else
                                                                If FieldArry(4) = 0 Then
                                                                    FieldTempText = GetSubStr(rsLabelRe(PE_CLng(FieldArry(0))), PE_CLng(FieldArry(2)), True)
                                                                Else
                                                                    FieldTempText = GetSubStr(rsLabelRe(PE_CLng(FieldArry(0))), PE_CLng(FieldArry(2)), False)
                                                                End If
                                                            End Select
                                                        End If
                                                     End If
                                                Case "Num" '按数字方式输出内容
                                                    If rsLabelRe(PE_CLng(FieldArry(0))) = "" Or IsNull(rsLabelRe(PE_CLng(FieldArry(0)))) Then
                                                        FieldTempText = "0"
                                                    Else
                                                        Select Case FieldArry(2)
                                                        Case 0
                                                            If FieldArry(3) = "0" Then
                                                                FieldTempText = Int(rsLabelRe(PE_CLng(FieldArry(0))))
                                                            Else
                                                                FieldTempText = String(Int(rsLabelRe(PE_CLng(FieldArry(0)))), FieldArry(3))
                                                            End If
                                                        Case 1
                                                            FieldTempText = FormatNumber(rsLabelRe(PE_CLng(FieldArry(0))), FieldArry(3))
                                                        Case 2
                                                            FieldTempText = FormatPercent(rsLabelRe(PE_CLng(FieldArry(0))))
                                                        End Select
                                                   End If
                                                Case "Time" '按时间方式输出内容
                                                    If rsLabelRe(PE_CLng(FieldArry(0))) = "" Or IsNull(rsLabelRe(PE_CLng(FieldArry(0)))) Then
                                                        FieldTempText = ""
                                                    Else
                                                        If IsDate(rsLabelRe(PE_CLng(FieldArry(0)))) Then '判断字段类型是否正确
                                                            temptime = rsLabelRe(PE_CLng(FieldArry(0)))
                                                            Select Case FieldArry(2)
                                                            Case 0
                                                                FieldTempText = Replace(Replace(Replace(Replace(Replace(Replace(FieldArry(3), "{year}", Year(temptime)), "{month}", Month(temptime)), "{day}", Day(temptime)), "{Hour}", Hour(temptime)), "{Minute}", Minute(temptime)), "{Second}", Second(temptime))
                                                            Case 1, 2
                                                                If FieldArry(2) = 1 Then
                                                                    temptimetext = Replace(FieldArry(3), "{year}", Year(temptime))
                                                                Else
                                                                    temptimetext = Replace(FieldArry(3), "{year}", Right(Year(temptime), 2))
                                                                End If
                                                                If Len(Month(temptime)) = 1 Then
                                                                    temptimetext = Replace(temptimetext, "{month}", "0" & Month(temptime))
                                                                Else
                                                                    temptimetext = Replace(temptimetext, "{month}", Month(temptime))
                                                                End If
                                                                If Len(Day(temptime)) = 1 Then
                                                                    temptimetext = Replace(temptimetext, "{day}", "0" & Day(temptime))
                                                                Else
                                                                    temptimetext = Replace(temptimetext, "{day}", Day(temptime))
                                                                End If
                                                                FieldTempText = temptimetext
                                                            Case 3
                                                                FieldTempText = FormatDateTime(temptime, PE_CLng(FieldArry(3)))
                                                            End Select
                                                        Else
                                                            FieldTempText = "本字段非时间型"
                                                        End If
                                                    End If
                                                Case "yn" '按是否方式输出内容
                                                    If rsLabelRe(PE_CLng(FieldArry(0))) = "" Or IsNull(rsLabelRe(PE_CLng(FieldArry(0)))) Then
                                                        FieldTempText = ""
                                                    Else
                                                        If rsLabelRe(PE_CLng(FieldArry(0))) = True Then
                                                            FieldTempText = FieldArry(2)
                                                        Else
                                                            FieldTempText = FieldArry(3)
                                                        End If
                                                    End If
                                                Case "GetUrl"
                                                    FieldTempText = GetInfoUrl(rsLabelRe(PE_CLng(FieldArry(0))), FieldArry(2), FieldArry(3))
                                                Case "GetClass"
                                                    FieldTempText = GetInfoClass(rsLabelRe(PE_CLng(FieldArry(0))), FieldArry(2))
                                                Case "GetSpecil"
                                                    FieldTempText = GetInfoSpecil(rsLabelRe(PE_CLng(FieldArry(0))), FieldArry(2))
                                                Case "GetChannel"
                                                    FieldTempText = GetInfoChannel(rsLabelRe(PE_CLng(FieldArry(0))), FieldArry(2))
                                                Case Else
                                                    FieldTempText = "标签参数错误"
                                                End Select
                                            Else
                                                FieldTempText = "标签参数错误"
                                            End If
                                            If Trim(FieldTempText & "") = "" Then
                                                InfoTempMatch = Replace(InfoTempMatch, FieldTemp, "")
                                            Else
                                                InfoTempMatch = Replace(InfoTempMatch, FieldTemp, FieldTempText)
                                            End If
                                        Next
                                        DyTemp = DyTemp & Replace(InfoTempMatch, "{$AutoID}", InfoID + 1)
                                        rsLabelRe.MoveNext
                                        InfoID = InfoID + 1
                                        If PageNum > 0 Then
                                            If InfoID >= rsLabel("PageNum") Then Exit Do
                                        End If
                                    End If
                                    Next
                                End If
                            Loop
                        End If
                    End If
                    rsLabelRe.Close
                    LoopTemp = Replace(LoopTemp, "{$SqlReplaceText}", DyTemp)
                    LoopTemp = Replace(LoopTemp, "{$totalPut}", totalPut)
                    If (PageNum > 0 And totalPut > PageNum) Or (reFlashTime > 10 And totalPut > 0) Then '如有分页属性或者刷新时间大于10秒,则加入DIV
                        dyfoot = "<script language=""JavaScript"" type=""text/JavaScript"">ShowDynaPage(" & rsLabel("LabelID") & ",1," & reFlashTime & ",'" & strInstallDir & "','none');</script>"
                        LoopTemp = Replace(Replace(Replace(XmlText("BaseText", "DynaPage", "<div id=""dyna_body_{dyid}"">{dybody}</div><div id=""dyna_page_{dyid}"" style=""text-align: right;"">{dyfoot}</div>"), "{dyid}", rsLabel("LabelID")), "{dybody}", LoopTemp), "{dyfoot}", dyfoot)
                    End If
                    strHtml = Replace(strHtml, "{$" & rsLabel("LabelName") & "}", LoopTemp)
                    PE_Cache.SetValue "{$" & rsLabel("LabelName") & "}", LoopTemp
                Else
                    strHtml = Replace(strHtml, "{$" & rsLabel("LabelName") & "}", PE_Cache.GetValue("{$" & rsLabel("LabelName") & "}"))
                End If

                End If
            Case 3 '函数型动态标签的处理过程
                    Dim tempvalue, loopTemp2
                    loopTemp2 = LoopTemp
                    loopTempMatch2 = loopTempMatch
                    regEx.Pattern = "\{\$" & rsLabel("LabelName") & "\((.*?)\)\}"
                    Set Matches3 = regEx.Execute(strHtml)
                    For Each Match3 In Matches3
                        TempSql = Replace(Replace(Replace(Replace(Replace(rsLabel("LabelIntro"), "{$Now}", Now()), "{$NowDay}", Day(Now())), "{$NowMonth}", Month(Now())), "{$NowYear}", Year(Now())),"{$PE_DatePart_D}", PE_DatePart_D)
                        TempSql = Replace(Replace(Replace(Replace(TempSql, "{$PE_True}", PE_True), "{$PE_False}", PE_False), "{$PE_Now}", PE_Now), "{$PE_OrderType}", PE_OrderType)
                        LoopTemp = loopTemp2
                        loopTempMatch = loopTempMatch2
                        DyTemp = ""
                        tempvalue = ""						
                        Dim TempNum, IsEnd
                        IsEnd = True
                        TempNum = 0
                        arrTemp = Split(Match3.SubMatches(0), ",")
                        ReDim arrTem(CInt(UBound(arrTemp)))
                        For i = 0 To UBound(arrTemp)
                            If InStr(arrTemp(i), "[") > 0 Then
                                IsEnd = False
                                arrTemp(i) = Replace(arrTemp(i), "[", "")
                            End If
                            If InStr(arrTemp(i), "]") > 0 Then
                                IsEnd = True
                                arrTemp(i) = Replace(arrTemp(i), "]", "")
                            End If
                            If IsEnd = False Then
                                arrTem(TempNum) = arrTem(TempNum) & arrTemp(i) & ","
                            End If
                            If IsEnd = True Then
                                arrTem(TempNum) = arrTem(TempNum) & arrTemp(i)

                                TempSql = Replace(TempSql, "{input(" & TempNum & ")}", ReplaceLabelBadChar(arrTem(TempNum))) 
								loopTempMatch = Replace(loopTempMatch, "{input(" & TempNum & ")}", arrTem(TempNum))
                                tempvalue = tempvalue & arrTem(TempNum) & "|"
                                TempNum = TempNum + 1
                            End If			
                        Next
                        TempSql = Replacesql(TempSql)
                        '开始循环处理内容
                        InfoID = 0
                        On Error Resume Next
                        Set rsLabelRe = Server.CreateObject("adodb.recordset")
                        rsLabelRe.Open TempSql, Conn, 1, 1
                        If Err Then
                            Err.Clear
                            DyTemp = "SQL查询错误"
                        Else
                            totalPut = rsLabelRe.RecordCount
                            If rsLabelRe.BOF And rsLabelRe.EOF Then
                                DyTemp = "尚无数据"
                            Else
                                Do While Not rsLabelRe.EOF
                                regEx.Pattern = "\{Infobegin\}([\s\S]*?)\{Infoend\}"
                                Set Matches = regEx.Execute(loopTempMatch)
                                If Matches.Count = 0 Then
                                    rsLabelRe.MoveNext
                                Else
                                    For Each Match In Matches
                                        If Not rsLabelRe.EOF Then
                                            InfoTemp = Match.Value
                                            InfoTempMatch = Replace(Replace(InfoTemp, "{Infobegin}", ""), "{Infoend}", "") '得到最终的单一字段内容
                                            regEx.Pattern = "\{\$Field\((.*?)\)\}"
                                            Set MatchesInfo = regEx.Execute(InfoTempMatch)
                                            For Each Match2 In MatchesInfo
                                                FieldTemp = Match2.Value
                                                FieldArry = Split(Match2.SubMatches(0), ",")
                                                If UBound(FieldArry) > 1 Then '参数正确,进行处理
                                                    Select Case FieldArry(1)
                                                    Case "Text" '按文本方式输出内容
                                                        If rsLabelRe(PE_CLng(FieldArry(0))) = "" Or IsNull(rsLabelRe(PE_CLng(FieldArry(0)))) Then
                                                            FieldTempText = ""
                                                        Else
                                                            If FieldArry(2) = 0 Then
                                                                Select Case FieldArry(3)
                                                                Case 1
                                                                    FieldTempText = Replace(rsLabelRe(PE_CLng(FieldArry(0))), "<", "&lt;")
                                                                Case 2
                                                                    FieldTempText = nohtml(rsLabelRe(PE_CLng(FieldArry(0))))
																	FieldTempText = Replace(FieldTempText, Chr(10), "<br>")
                                                                Case Else
                                                                    FieldTempText = rsLabelRe(PE_CLng(FieldArry(0)))
                                                                End Select
                                                            Else
                                                                Select Case FieldArry(3)
                                                                Case 1
                                                                    If FieldArry(4) = 0 Then
                                                                        FieldTempText = GetSubStr(Replace(rsLabelRe(PE_CLng(FieldArry(0))), "<", "&lt;"), PE_CLng(FieldArry(2)), True)
                                                                    Else
                                                                        FieldTempText = GetSubStr(Replace(rsLabelRe(PE_CLng(FieldArry(0))), "<", "&lt;"), PE_CLng(FieldArry(2)), False)
                                                                    End If
                                                                Case 2
                                                                    If FieldArry(4) = 0 Then
                                                                        FieldTempText = GetSubStr(nohtml(rsLabelRe(PE_CLng(FieldArry(0)))), PE_CLng(FieldArry(2)), True)                                           
																		FieldTempText = Replace(FieldTempText, Chr(10), "<br>")
                                                                    Else
                                                                        FieldTempText = GetSubStr(nohtml(rsLabelRe(PE_CLng(FieldArry(0)))), PE_CLng(FieldArry(2)), False)
																		FieldTempText = Replace(FieldTempText, Chr(10), "<br>")
                                                                    End If
                                                                Case Else
                                                                    If FieldArry(4) = 0 Then
                                                                        FieldTempText = GetSubStr(rsLabelRe(PE_CLng(FieldArry(0))), PE_CLng(FieldArry(2)), True)
                                                                    Else
                                                                        FieldTempText = GetSubStr(rsLabelRe(PE_CLng(FieldArry(0))), PE_CLng(FieldArry(2)), False)
                                                                    End If
                                                                End Select
                                                            End If
                                                        End If
                                                    Case "Num" '按数字方式输出内容
                                                        If rsLabelRe(PE_CLng(FieldArry(0))) = "" Or IsNull(rsLabelRe(PE_CLng(FieldArry(0)))) Then
                                                            FieldTempText = "0"
                                                        Else
                                                            Select Case FieldArry(2)
                                                            Case 0
                                                                If FieldArry(3) = "0" Then
                                                                    FieldTempText = Int(rsLabelRe(PE_CLng(FieldArry(0))))
                                                                Else
                                                                    FieldTempText = String(Int(rsLabelRe(PE_CLng(FieldArry(0)))), FieldArry(3))
                                                                End If
                                                            Case 1
                                                                FieldTempText = FormatNumber(rsLabelRe(PE_CLng(FieldArry(0))), FieldArry(3))
                                                            Case 2
                                                                FieldTempText = FormatPercent(rsLabelRe(PE_CLng(FieldArry(0))))
                                                            End Select
                                                        End If
                                                    Case "Time" '按时间方式输出内容
                                                        If rsLabelRe(PE_CLng(FieldArry(0))) = "" Or IsNull(rsLabelRe(PE_CLng(FieldArry(0)))) Then
                                                            FieldTempText = ""
                                                        Else
                                                            If IsDate(rsLabelRe(PE_CLng(FieldArry(0)))) Then '判断字段类型是否正确
                                                                temptime = rsLabelRe(PE_CLng(FieldArry(0)))
                                                                Select Case FieldArry(2)
                                                                Case 0
                                                                    FieldTempText = Replace(Replace(Replace(Replace(Replace(Replace(FieldArry(3), "{year}", Year(temptime)), "{month}", Month(temptime)), "{day}", Day(temptime)), "{Hour}", Hour(temptime)), "{Minute}", Minute(temptime)), "{Second}", Second(temptime))
                                                                Case 1, 2
                                                                    If FieldArry(2) = 1 Then
                                                                        temptimetext = Replace(FieldArry(3), "{year}", Year(temptime))
                                                                    Else
                                                                        temptimetext = Replace(FieldArry(3), "{year}", Right(Year(temptime), 2))
                                                                    End If
                                                                    If Len(Month(temptime)) = 1 Then
                                                                        temptimetext = Replace(temptimetext, "{month}", "0" & Month(temptime))
                                                                    Else
                                                                        temptimetext = Replace(temptimetext, "{month}", Month(temptime))
                                                                    End If
                                                                    If Len(Day(temptime)) = 1 Then
                                                                        temptimetext = Replace(temptimetext, "{day}", "0" & Day(temptime))
                                                                    Else
                                                                        temptimetext = Replace(temptimetext, "{day}", Day(temptime))
                                                                    End If
                                                                    FieldTempText = temptimetext
                                                                Case 3
                                                                    FieldTempText = FormatDateTime(temptime, PE_CLng(FieldArry(3)))
                                                                End Select
                                                            Else
                                                                FieldTempText = "本字段非时间型"
                                                            End If
                                                        End If
                                                    Case "yn" '按是否方式输出内容
                                                        If rsLabelRe(PE_CLng(FieldArry(0))) = "" Or IsNull(rsLabelRe(PE_CLng(FieldArry(0)))) Then
                                                            FieldTempText = ""
                                                        Else
                                                            If rsLabelRe(PE_CLng(FieldArry(0))) = True Then
                                                                FieldTempText = FieldArry(2)
                                                            Else
                                                                FieldTempText = FieldArry(3)
                                                            End If
                                                        End If
                                                    Case "GetUrl"
                                                        FieldTempText = GetInfoUrl(rsLabelRe(PE_CLng(FieldArry(0))), FieldArry(2), FieldArry(3))
                                                    Case "GetClass"
                                                        FieldTempText = GetInfoClass(rsLabelRe(PE_CLng(FieldArry(0))), FieldArry(2))
                                                    Case "GetSpecil"
                                                        FieldTempText = GetInfoSpecil(rsLabelRe(PE_CLng(FieldArry(0))), FieldArry(2))
                                                    Case "GetChannel"
                                                        FieldTempText = GetInfoChannel(rsLabelRe(PE_CLng(FieldArry(0))), FieldArry(2))
                                                    Case Else
                                                        FieldTempText = "标签参数错误"
                                                    End Select
                                                Else
                                                    FieldTempText = "标签参数错误"
                                                End If
                                                InfoTempMatch = Replace(InfoTempMatch, FieldTemp, FieldTempText)
                                            Next
                                            DyTemp = DyTemp & Replace(InfoTempMatch, "{$AutoID}", InfoID + 1)
                                            rsLabelRe.MoveNext
                                            InfoID = InfoID + 1
                                            If PageNum > 0 Then
                                                If InfoID >= PageNum Then Exit Do
                                            End If
                                        End If
                                    Next
                                End If
                            Loop
                        End If
                        End If
                        rsLabelRe.Close

                        LoopTemp = Replace(LoopTemp, "{$SqlReplaceText}", DyTemp)
                        LoopTemp = Replace(LoopTemp, "{$totalPut}", totalPut)
                        For i = 0 To UBound(arrTemp)
                            LoopTemp = Replace(LoopTemp, "{input(" & i & ")}", arrTem(i))
                        Next
                        If (PageNum > 0 And totalPut > PageNum) Or (reFlashTime > 10 And totalPut > 0) Then '如有分页属性或者刷新时间大于10秒,则加入DIV
                            dyfoot = "<script language=""JavaScript"" type=""text/JavaScript"">ShowDynaPage(" & rsLabel("LabelID") & ",1," & reFlashTime & ",'" & strInstallDir & "','" & tempvalue & "');</script>"
                            LoopTemp = Replace(Replace(Replace(XmlText("BaseText", "DynaPage", "<div id=""dyna_body_{dyid}"">{dybody}</div><div id=""dyna_page_{dyid}"" style=""text-align: right;"">{dyfoot}</div>"), "{dyid}", rsLabel("LabelID")), "{dybody}", LoopTemp), "{dyfoot}", dyfoot)
                        End If
                        strHtml = Replace(strHtml, Match3.Value, LoopTemp)
                    Next
            End Select
            DyTemp = ""
            nonLabel = False
        End If
        rsLabel.MoveNext
        Loop
        Set rsLabelRe = Nothing
        Set MatchesInfo = Nothing

        Dim Xmlhttptime
        Xmlhttptime = PE_CLng(XmlText("Site", "Xmlhttptime", "10"))
    
        '替换采集标签
        Dim caistat
        Set rsLabel = Conn.Execute("select LabelName,LabelType,PageNum,LabelIntro,LabelContent,AreaCollectionID from PE_Label where LabelType=2 order by Priority asc,LabelID asc")
        Do While Not rsLabel.EOF
            If InStr(strHtml, "{$" & rsLabel("LabelName") & "}") > 0 Then
                caistat = False

                CaiTemp = Split(PE_Cache.GetValue("{$" & rsLabel("LabelName") & "}"), "|$|$|")

                If UBound(CaiTemp) = 1 Then
                    If IsDate(CaiTemp(0)) And DateDiff("n", CaiTemp(0), Now()) < Xmlhttptime Then
                            caistat = True
                        End If
                End If

                If caistat = False Then
                    If rsLabel("AreaCollectionID") = 0 Then
                        CaiTemp = GetHttpPage(rsLabel("LabelIntro"), 0)
                    Else
                        Set rsArea = Conn.Execute("select Top 1 AreaID,Code,StringReplace,LableStart,LableEnd,FilterProperty,UpFileType,Type from PE_AreaCollection where AreaID=" & rsLabel("AreaCollectionID") & " and Type=1")
                        If Not rsArea.EOF Then
                            CaiTemp = GetHttpPage(rsLabel("LabelIntro"), rsArea("Code"))
                            If CaiTemp <> "" Then
                                CaiTemp = GetBody(CaiTemp, rsArea("LableStart"), rsArea("LableEnd"), True, True)
                                CaiTemp = ReplaceStringPath(CaiTemp, rsLabel("LabelIntro"), rsArea("UpFileType"))

                                If rsArea("StringReplace") <> "" Then
                                    arrAreaCode = Split(rsArea("StringReplace"), "$$$")
                                    For i = 0 To UBound(arrAreaCode)
                                        arrAreaCode2 = Split(arrAreaCode(i), "|||")
                                        CaiTemp = Replace(CaiTemp, arrAreaCode2(0), arrAreaCode2(1))
                                    Next
                                End If
                                CaiTemp = FilterScript(CaiTemp, rsArea("FilterProperty"))
                            End If
                        Else
                            CaiTemp = GetHttpPage(rsLabel("LabelIntro"), 0)
                        End If
                        rsArea.Close
                    End If
                    strHtml = Replace(strHtml, "{$" & rsLabel("LabelName") & "}", CaiTemp)

                    PE_Cache.SetValue "{$" & rsLabel("LabelName") & "}", Now() & "|$|$|" & CaiTemp
                Else
                    strHtml = Replace(strHtml, "{$" & rsLabel("LabelName") & "}", CaiTemp(1))
                End If

                nonLabel = False
            End If
            rsLabel.MoveNext
        Loop
        CaiTemp = Null
        Set rsArea = Nothing

        If nonLabel = True Then
            LabelNum = 0
            Exit Do
        Else
            If InStr(strHtml, "{$MY_") > 0 Then
                LabelNum = 1
                If Looptotalnum > 0 Then
                    Looptotalnum = Looptotalnum - 1
                Else
                    LabelNum = 0
                    Exit Do
                End If
            Else
                LabelNum = 0
                Exit Do
            End If
        End If
        rsLabel.Close
        Set rsLabel = Nothing
    Loop
    strHtml = PE_Replace(strHtml, "%7B", "{")
    strHtml = PE_Replace(strHtml, "%7D", "}") 
    strHtml = Replace(strHtml, "<!--{$", "{$")
    strHtml = Replace(strHtml, "}-->", "}") 
    regEx.Pattern = "\{\$InstallDir\}(?!\{\$ChannelDir\})"
    strHtml = regEx.Replace(strHtml, strInstallDir)
    strHtml = PE_Replace(strHtml, "{$ADDir}", ADDir)
    strHtml = PE_Replace(strHtml, "{$SiteUrl}", SiteUrl)
    strHtml = PE_Replace(strHtml, "{$SiteName}", SiteName)
    strHtml = PE_Replace(strHtml, "{$WebmasterEmail}", WebmasterEmail)
    strHtml = PE_Replace(strHtml, "{$WebmasterName}", WebmasterName)
    strHtml = PE_Replace(strHtml, "{$Copyright}", Copyright)
    strHtml = PE_Replace(strHtml, "{$Meta_Keywords}", Meta_Keywords)
    strHtml = PE_Replace(strHtml, "{$Meta_Description}", Meta_Description)
    strHtml = Replace(strHtml, "{$ShowAD}", "")
    If InStr(strHtml, "{$ShowLogo}") > 0 Then strHtml = Replace(strHtml, "{$ShowLogo}", GetLogo(180, 60))
    If InStr(strHtml, "{$ShowBanner}") > 0 Then strHtml = Replace(strHtml, "{$ShowBanner}", GetBanner(480, 60))
    If InStr(strHtml, "{$ShowSiteCountAll}") > 0 Then strHtml = Replace(strHtml, "{$ShowSiteCountAll}", GetSiteCountAll())
    If InStr(strHtml, "{$ShowChannel}") > 0 Then strHtml = Replace(strHtml, "{$ShowChannel}", GetChannelList(0))    
    If InStr(strHtml, "{$GetUserName}") > 0 Then strHtml = Replace(strHtml, "{$GetUserName}", GetUserName())  
    If InStr(strHtml, "{$AdminDir}") > 0 Then strHtml = Replace(strHtml, "{$AdminDir}", AdminDir)  	 
    If InStr(strHtml, "{$ShowVoteJS_Comment}") > 0 Then strHtml = Replace(strHtml, "{$ShowVoteJS_Comment}", ShowVoteJS_Comment())	
    If ShowAdminLogin = True Then
        strHtml = Replace(strHtml, "{$ShowAdminLogin}", " <a class='Bottom' href='" & strInstallDir & AdminDir & "/Admin_Index.asp' target='_blank'>" & XmlText("Site", "ReplaceCommon/AdminLogin", "管理登录") & "</a>&nbsp;" & XmlText("Site", "ReplaceCommon/a2", "|") & "&nbsp;")
    Else
        strHtml = Replace(strHtml, "{$ShowAdminLogin}", "")
    End If	
    '替换{$YN(Condition,Fir,Sec)}标签
    Dim strYN
    regEx.Pattern = "\{\$YN\((.*?)\)\}[^\]]"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
    Dim TempNum1
    IsEnd = True
    TempNum1 = 0
    arrTemp = Split(Match.SubMatches(0), ",")
    ReDim arrTem(CInt(UBound(arrTemp)))
    For i = 0 To UBound(arrTemp)
        If InStr(arrTemp(i), "[") > 0 Then
            IsEnd = False
            arrTemp(i) = Replace(arrTemp(i), "[", "")
        End If
        If InStr(arrTemp(i), "]") > 0 Then
            IsEnd = True
            arrTemp(i) = Replace(arrTemp(i), "]", "")
        End If
        If IsEnd = False Then
            arrTem(TempNum1) = arrTem(TempNum1) & arrTemp(i) & ","
        End If
        If IsEnd = True Then
            arrTem(TempNum1) = arrTem(TempNum1) & arrTemp(i)
            tempSql = Replace(tempSql, "{input(" & TempNum1 & ")}", arrTem(TempNum1))
            TempNum1 = TempNum1 + 1
        End If
    Next		
        If TempNum1 <> 3 Then 
            strYN = "函数式标签：{$YN(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Else
            strYN = YN(arrTem(0),arrTem(1),arrTem(2)) 
        End If
        strHtml = Replace(strHtml, Left(Match.value,len(Match.value)-1), strYN)
    Next
	
    '替换{$GetLanguage(BigNode,SmallNode,DefChar)}标签
    Dim strLanguage
    regEx.Pattern = "\{\$GetLanguage\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) <> 2 Then
            strLanguage = "函数式标签：{$GetLanguage(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Else
            strLanguage = XmlText(arrTemp(0), arrTemp(1), arrTemp(2))
        End If
        strHtml = Replace(strHtml, Match.Value, strLanguage)
    Next
	
    Dim strSlidePicJs
    regEx.Pattern = "\{\$SlidePicJs\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) <> 9 and UBound(arrTemp)<>10 Then
            strSlidePicJs = "函数式标签：{$SlidePicJs(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Elseif UBound(arrTemp) = 9 then
            strSlidePicJs = SlidePicJs(PE_Clng(arrTemp(0)),PE_CLng(arrTemp(1)),PE_CLng(arrTemp(2)),PE_CLng(arrTemp(3)),PE_CLng(arrTemp(4)),arrTemp(5),0,PE_CLng(arrTemp(6)),PE_CLng(arrTemp(7)),PE_CLng(arrTemp(8)),PE_CLng(arrTemp(9)))
        Else
            strSlidePicJs = SlidePicJs(PE_Clng(arrTemp(0)),PE_CLng(arrTemp(1)),PE_CLng(arrTemp(2)),PE_CLng(arrTemp(3)),PE_CLng(arrTemp(4)),arrTemp(5),arrTemp(6),PE_CLng(arrTemp(7)),PE_CLng(arrTemp(8)),PE_CLng(arrTemp(9)),PE_CLng(arrTemp(10)))			
        End If
	strHtml = Replace(strHtml, Match.value, strSlidePicJs)	
    Next
	
    Dim strIsLogin
    regEx.Pattern = "\{\$IsLogin\(([\s\S]*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
	    If UBound(arrTemp) <> 1 Then
            strIsLogin = "函数式标签：{$IsLogin(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Else
        strIsLogin = IsLogin(arrTemp(0),arrTemp(1))
        End If
	strHtml = Replace(strHtml, Match.value, strIsLogin)	
    Next
    
    '替换频道导航
    regEx.Pattern = "\{\$ShowChannel\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        strChannel = GetChannelList(PE_CLng(arrTemp(0)))
        strHtml = Replace(strHtml, Match.Value, strChannel)
    Next

    '替换Logo
    regEx.Pattern = "\{\$ShowLogo\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) <> 1 Then
            strLogo = "函数式标签：{$ShowLogo(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Else
            strLogo = GetLogo(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)))
        End If
        strHtml = Replace(strHtml, Match.Value, strLogo)
    Next
    
    '替换banner
    regEx.Pattern = "\{\$ShowBanner\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) <> 1 Then
            strBanner = "函数式标签：{$ShowBanner(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Else
            strBanner = GetBanner(arrTemp(0), arrTemp(1))
        End If
        strHtml = Replace(strHtml, Match.Value, strBanner)
    Next
    
    
    '替换广告
    regEx.Pattern = "\{\$ShowAD\((.*?)\)\}"
    strHtml = regEx.Replace(strHtml, "")
    
    '替换指定ID广告
    regEx.Pattern = "\{\$GetAD\((.*?)\)\}"
    strHtml = regEx.Replace(strHtml, "")
    
    
    '替换调查
    If InStr(strHtml, "{$ShowVote}") > 0 Then strHtml = Replace(strHtml, "{$ShowVote}", GetVote())

    '替换用户排行
    regEx.Pattern = "\{\$ShowTopUser\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        Select Case UBound(arrTemp)
        Case 0
            strTopUser = GetTopUser(PE_CLng(arrTemp(0)), 1, True, True, False, False, "more...", 1)
        Case 6
            strTopUser = GetTopUser(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CBool(arrTemp(2)), PE_CBool(arrTemp(3)), PE_CBool(arrTemp(4)), PE_CBool(arrTemp(5)), arrTemp(6), 1)
        Case 7
            strTopUser = GetTopUser(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CBool(arrTemp(2)), PE_CBool(arrTemp(3)), PE_CBool(arrTemp(4)), PE_CBool(arrTemp(5)), arrTemp(6), arrTemp(7))
        Case Else
            strTopUser = "标签{$ShowTopUser(参数列表)}的参数个数不对"
        End Select
        strHtml = Replace(strHtml, Match.Value, strTopUser)
    Next

    '替换聚合列表
    regEx.Pattern = "\{\$ShowSpaceList\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) < 8 Then
            strAuthorList = "函数式标签：{$ShowSpaceList(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Else
            strAuthorList = GetBlogList(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CBool(arrTemp(3)), PE_CBool(arrTemp(4)), PE_CLng(arrTemp(5)), PE_CLng(arrTemp(6)), PE_CLng(arrTemp(7)), PE_CLng(arrTemp(8)), arrTemp(9), PE_CLng(arrTemp(10)))
        End If
        strHtml = Replace(strHtml, Match.Value, strAuthorList)
    Next

    '替换作者列表
    regEx.Pattern = "\{\$ShowAuthorList\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) < 8 Then
            strAuthorList = "函数式标签：{$ShowAuthorList(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Else
            If UBound(arrTemp) = 8 Then
                strAuthorList = GetAuthorList(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)), PE_CLng(arrTemp(5)), PE_CLng(arrTemp(6)), PE_CLng(arrTemp(7)), arrTemp(8), 1)
            Else 
                strAuthorList = GetAuthorList(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)), PE_CLng(arrTemp(5)), PE_CLng(arrTemp(6)), PE_CLng(arrTemp(7)), arrTemp(8), PE_CLng(arrTemp(9)))
            End If
        End If
        strHtml = Replace(strHtml, Match.Value, strAuthorList)
    Next

    '替换厂商列表
    regEx.Pattern = "\{\$ShowProducerList\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) < 9 Then
            strProducerList = "函数式标签：{$ShowProducerList(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Else
            strProducerList = GetProducerList(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)), PE_CLng(arrTemp(5)), PE_CLng(arrTemp(6)), PE_CLng(arrTemp(7)), arrTemp(8), PE_CLng(arrTemp(9)))
        End If
        strHtml = Replace(strHtml, Match.Value, strProducerList)
    Next

    '替换友情链接
    regEx.Pattern = "\{\$ShowFriendSite\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) < 3 Then
            strFriendSite = "函数式标签：{$ShowFriendSite(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Else
            If UBound(arrTemp) = 5 Then
                strFriendSite = ShowFriendSite(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), arrTemp(4), arrTemp(5), 88, False, True, 0)
            ElseIf UBound(arrTemp) = 6 Then
                strFriendSite = ShowFriendSite(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), arrTemp(4), arrTemp(5), PE_CLng(arrTemp(6)), False, True, 0)
            ElseIf UBound(arrTemp) = 9 Then
                strFriendSite = ShowFriendSite(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), arrTemp(4), arrTemp(5), PE_CLng(arrTemp(6)), PE_CBool(arrTemp(7)), PE_CBool(arrTemp(8)), PE_CLng(arrTemp(9)))
            Else
                strFriendSite = ShowFriendSite(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), 0, 0, 88, False, True, 0)
            End If
        End If
        strHtml = Replace(strHtml, Match.Value, strFriendSite)
    Next

    '替换公告
    regEx.Pattern = "\{\$ShowAnnounce\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) < 1 Then
            strAnnounce = "函数式标签：{$ShowAnnounce(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Else
            Select Case UBound(arrTemp)
            Case 1
                strAnnounce = ShowAnnounce(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), True, True, 100)
            Case 3
                strAnnounce = ShowAnnounce(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CBool(arrTemp(2)), PE_CBool(arrTemp(3)), 100)
            Case 4
                strAnnounce = ShowAnnounce(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CBool(arrTemp(2)), PE_CBool(arrTemp(3)), PE_CLng(arrTemp(4)))
            Case Else
                strAnnounce = ShowAnnounce(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), True, True, 100)
            End Select
        End If
        strHtml = Replace(strHtml, Match.Value, strAnnounce)
    Next

    '替换指定专题列表
    Dim strSpecial, arrTemp2
    regEx.Pattern = "\{\$ShowSpecialList\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp2 = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp2) + 1 <> 6 Then
            strSpecial = "函数式标签：{$ShowSpecialList(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Else
            strSpecial = ShowSpecialList(PE_CLng(arrTemp2(0)), PE_CBool(arrTemp2(1)), PE_CLng(arrTemp2(2)), PE_CLng(arrTemp2(3)), PE_CLng(arrTemp2(4)), PE_CLng(arrTemp2(5)))
        End If
        strHtml = Replace(strHtml, Match.Value, strSpecial)
    Next
    '替换弹出式公告
    regEx.Pattern = "\{\$PopAnnouceWindow\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) <> 1 Then
            strPopAnnouce = "函数式标签：{$PopAnnouceWindow(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Else
            strPopAnnouce = PopAnnouceWindow(arrTemp(0), arrTemp(1))
        End If
        strHtml = Replace(strHtml, Match.Value, strPopAnnouce)
    Next

    If ChannelID > 0 Then
        strHtml = PE_Replace(strHtml, "{$InstallDir}{$ChannelDir}", ChannelUrl)
        strHtml = PE_Replace(strHtml, "{$ChannelID}", ChannelID)
        strHtml = PE_Replace(strHtml, "{$ChannelDir}", ChannelDir)
        strHtml = PE_Replace(strHtml, "{$ChannelUrl}", ChannelUrl)
        strHtml = PE_Replace(strHtml, "{$ChannelName}", ChannelName)
        strHtml = PE_Replace(strHtml, "{$ChannelShortName}", ChannelShortName)
        strHtml = PE_Replace(strHtml, "{$UploadDir}", UploadDir)
        strHtml = PE_Replace(strHtml, "{$ChannelPicUrl}", ChannelPicUrl)
        strHtml = PE_Replace(strHtml, "{$Meta_Keywords_Channel}", Meta_Keywords_Channel)
        strHtml = PE_Replace(strHtml, "{$Meta_Description_Channel}", Meta_Description_Channel)
        '自设内容
        strHtml = CustomContent("Channel", Custom_Content_Channel, strHtml)
    End If
    If strInstallDir<>"/" Then strHtml = PE_Replace(strHtml, strInstallDir & strInstallDir, strInstallDir)'兼容自动填充{$InstallDir}出现地址错误以及兼容频道变子站时标签内置方法获取标签路径写法
    If InStr(strHtml, "{$MenuJS}") > 0 Then strHtml = PE_Replace(strHtml, "{$MenuJS}", GetMenuJS(ChannelDir, ShowClassTreeGuide))
    If InStr(strHtml, "{$Skin_CSS}") > 0 Then strHtml = PE_Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(SkinID))
    
    '替换底部栏目导航标签
    Dim strNavigation
    regEx.Pattern = "\{\$ShowClassNavigation\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) <> 2 Then
            strNavigation = "函数式标签：{$ShowClassNavigation(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Else
            strNavigation = GetClass_Navigation(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)))
        End If
        strHtml = Replace(strHtml, Match.Value, strNavigation)
    Next

    Dim strBroClass
    regEx.Pattern = "\{\$GetBrotherClass\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) <> 4 Then
            strBroClass = "函数式标签：{$GetBrotherClass(参数列表)}的参数个数不对。请检查模板中的此标签。"
        Else
            strBroClass = GetBrotherClass(PE_CLng(arrTemp(0)), PE_CLng(arrTemp(1)), PE_CLng(arrTemp(2)), PE_CLng(arrTemp(3)), PE_CLng(arrTemp(4)))
        End If
        strHtml = Replace(strHtml, Match.Value, strBroClass)
    Next
    
    Dim strChildClass
    regEx.Pattern = "\{\$ShowChildClass\((.*?)\)\}"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        arrTemp2 = Split(Match.SubMatches(0), ",")
        '此处判断为兼容旧版本标签GetChildClass(1, 0)
        '函数名：GetChildClass
        '参  数：ShowType--------显示方式，1为竖向列表，2为横向列表
        '        Cols ----- 每行显示多少个栏目，竖向列表时无效
        If UBound(arrTemp2) + 1 = 2 Then
            Select Case PE_CLng(arrTemp2(0))
            Case 1
                strChildClass = GetChildClass(0, 0, 3, 3, 0, True)
            Case 2
                strChildClass = GetChildClass(0, 0, 3, 3, PE_CLng(arrTemp2(1)), True)
            Case Else
                strChildClass = GetChildClass(0, 0, 3, 3, 0, True)
            End Select
        ElseIf UBound(arrTemp2) + 1 = 6 Then
            Select Case LCase(arrTemp2(0))
            Case "classid"
                strTemp = ClassID
            Case "parentid"
                strTemp = ParentID
            Case Else
                strTemp = PE_CLng(LCase(arrTemp2(0)))
            End Select
            strChildClass = GetChildClass(strTemp, PE_CLng(arrTemp2(1)), PE_CLng(arrTemp2(2)), PE_CLng(arrTemp2(3)), PE_CLng(arrTemp2(4)), PE_CBool(arrTemp2(5)))
        Else
            strChildClass = "函数式标签：{$ShowChildClass(参数列表)}的参数个数不对。请检查模板中的此标签。"
        End If
        strHtml = Replace(strHtml, Match.Value, strChildClass)
    Next
End Sub
'=================================================
'函数名：ShowSpecialList
'作  用：显示指定频道专题
'参  数：
'1       ChannelID ---- 频道ID,0为全站专题，-1为所有频道专题
'2       IsElite ---- 是否是推荐专题，True为只显示推荐专题，False为显示所有专题
'3       SpecialNum  ------最多显示多少个专题名称
'4       ShowPropertyType ---- 显示前的小图标，0为不显示，1为符号，其他为小图片：/images/Special_List*.gif
'5       OpenType ---- 打开方式，0为在原窗口打开，1为在新窗口打开
'6       Cols ---- 每行的列数。超过此列数就换行。
'=================================================

Function ShowSpecialList(ChannelID, IsElite, SpecialNum, ShowPropertyType, OpenType, Cols)
    Dim sqlSpecial, rsSpecial, strSpecial, i
    If SpecialNum <= 0 Or SpecialNum > 100 Then
        SpecialNum = 10
    End If
    If Cols = 0 Then Cols = 1
    If ChannelID = -1 Then
        sqlSpecial = "select S.ChannelID,S.SpecialID,S.SpecialName,S.SpecialDir,C.ChannelDir,C.FileExt_List,C.UseCreateHTML,S.Tips from PE_Special S left join PE_Channel C on S.ChannelID=C.ChannelID where 1=1"
    ElseIf ChannelID = 0 Then
        sqlSpecial = "select ChannelID,SpecialID,SpecialName,SpecialDir,Tips from PE_Special where ChannelID=0"
    Else
        sqlSpecial = "select S.ChannelID,S.SpecialID,S.SpecialName,S.SpecialDir,S.Tips,C.ChannelDir,C.FileExt_List,C.UseCreateHTML from PE_Special S left join PE_Channel C on S.ChannelID=C.ChannelID where S.ChannelID=" & ChannelID & ""
    End If
    If IsElite = True Then
        If ChannelID = 0 Then
            sqlSpecial = sqlSpecial & " and IsElite=" & PE_True & " order by OrderID"
        Else
            sqlSpecial = sqlSpecial & " and S.IsElite=" & PE_True & " order by S.OrderID"
        End If
    End If
    
    Set rsSpecial = Conn.Execute(sqlSpecial)
    If rsSpecial.BOF And rsSpecial.EOF Then
        strSpecial = "&nbsp;没有任何专题栏目"
    Else
        i = 0
        Do While Not rsSpecial.EOF
            If i > 0 Then
                If i Mod Cols = 0 Then
                    strSpecial = strSpecial & "<br>"
                Else
                    strSpecial = strSpecial & "&nbsp;&nbsp;"
                End If
            
            End If

            If ShowPropertyType = 0 Then
                strSpecial = strSpecial & ""
            ElseIf ShowPropertyType = 1 Then
                strSpecial = strSpecial & "·"
            Else
                strSpecial = strSpecial & "<img src='" & strInstallDir & "images/Special_List" & ShowPropertyType & ".gif' border='0'>"
            End If
            If rsSpecial("ChannelID") <> 0 Then
                If rsSpecial("UseCreateHTML") = 1 Or rsSpecial("UseCreateHTML") = 3 Then
                    strSpecial = strSpecial & "&nbsp;<a href='" & strInstallDir & rsSpecial("ChannelDir") & "/Special/" & rsSpecial("SpecialDir") & "/Index" & arrFileExt(rsSpecial("FileExt_List")) & "'"
                Else
                    strSpecial = strSpecial & "&nbsp;<a href='" & strInstallDir & rsSpecial("ChannelDir") & "/ShowSpecial.asp?SpecialID=" & rsSpecial("SpecialID") & "'"
                End If
            Else
                If FileExt_SiteSpecial <> ".asp" Then
                    strSpecial = strSpecial & "&nbsp;<a href='" & strInstallDir & "Special/" & rsSpecial("SpecialDir") & "/Index" & FileExt_SiteSpecial & "'"
                Else
                    strSpecial = strSpecial & "&nbsp;<a href='" & strInstallDir & "ShowSpecial.asp?SpecialID=" & rsSpecial("SpecialID") & "'"
                End If
            End If
            strSpecial = strSpecial & " title='" & Trim(nohtml(rsSpecial("Tips"))) & "'"
            If OpenType = 0 Then
                strSpecial = strSpecial & " target=""_self"">"
            Else
                strSpecial = strSpecial & " target=""_blank"">"
            End If
            strSpecial = strSpecial & rsSpecial("SpecialName") & "</a>"
            rsSpecial.MoveNext
            i = i + 1
            If i >= SpecialNum Then Exit Do
        Loop
    End If
    If Not rsSpecial.EOF Then
        If ChannelID = -1 Or ChannelID = 0 Then
            strSpecial = strSpecial & "<br><p align='right'><a href='" & strInstallDir & "SpecialList.asp'>更多专题</a></p>"
        Else
            strSpecial = strSpecial & "<br><p align='right'><a href='" & strInstallDir & rsSpecial("ChannelDir") & "/SpecialList.asp'>更多专题</a></p>"
        End If
    End If
    rsSpecial.Close
    Set rsSpecial = Nothing
    ShowSpecialList = strSpecial
End Function

'==================================================
'函数名：GetInfoChannel
'作  用：获取对象的频道参数
'参  数：InfoID ------对象ID
'      ：OutType -----输出方式
'==================================================
Function GetInfoChannel(InfoID, OutType)
    If IsNull(InfoID) = True Or IsNull(OutType) = True Then
        GetInfoChannel = ""
        Exit Function
    End If
    Dim sqlInfo, rsInfo, rsChannel2, strTemp
    sqlInfo = "select top 1 ChannelID,ChannelName,LinkUrl,ChannelDir,Disabled,UploadDir from PE_Channel Where ChannelID=" & PE_CLng(InfoID)
    Set rsInfo = Conn.Execute(sqlInfo)
    If Not (rsInfo.BOF And rsInfo.EOF) Then
        If rsInfo("Disabled") = True Then
                strTemp = ""
        Else
            Select Case OutType
            Case 1
                If IsNull(rsInfo("ChannelDir")) Then
                    strTemp = rsInfo("LinkUrl")
                Else
                    strTemp = rsInfo("ChannelDir")
                End If
            Case 2
                strTemp = rsInfo("ChannelName")
            Case 3
                strTemp = rsInfo("UploadDir")
            Case Else
                strTemp = "标签参数错"
            End Select
        End If
    End If
    rsInfo.Close
    Set rsInfo = Nothing
    GetInfoChannel = strTemp
End Function

'==================================================
'函数名：GetInfoUrl
'作  用：获取对象的路径
'参  数：InfoID ------对象ID
'      ：DataType ------数据库名称
'==================================================
Function GetInfoUrl(InfoID, DataType, OutType)
    If IsNull(InfoID) = True Or IsNull(DataType) = True Or IsNull(OutType) = True Then
        GetInfoUrl = ""
        Exit Function
    End If
    Dim sqlInfo, rsInfo, rsChannel2, strTemp
    Dim ChannelDir, StructureType, FileNameType, FileExtType, iUseCreateHTML, CacheTemp, ChannelTemp,ChannelUrl
    Select Case DataType
    Case "Article"
        sqlInfo = "select top 1 A.ArticleID,A.ChannelID,A.ClassID,A.Title,A.UpdateTime,A.InfoPoint,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Article A Left join PE_Class C on A.ClassID=C.ClassID Where A.ArticleID=" & PE_CLng(InfoID)
    Case "Soft"
        sqlInfo = "select top 1 A.SoftID,A.ChannelID,A.ClassID,A.SoftName,A.UpdateTime,A.InfoPoint,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Soft A Left join PE_Class C on A.ClassID=C.ClassID Where A.SoftID=" & PE_Clng(InfoID)
    Case "Photo"
        sqlInfo = "select top 1 A.PhotoID,A.ChannelID,A.ClassID,A.PhotoName,A.UpdateTime,A.InfoPoint,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Photo A Left join PE_Class C on A.ClassID=C.ClassID Where A.PhotoID=" & PE_CLng(InfoID)
    Case "Product"
        sqlInfo = "select top 1 A.ProductID,A.ChannelID,A.ClassID,A.ProductName,A.UpdateTime,A.Stocks,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Product A Left join PE_Class C on A.ClassID=C.ClassID Where A.ProductID=" & PE_CLng(InfoID)
    Case Else
        GetInfoUrl = InfoID
        Exit Function
    End Select
    Set rsInfo = Conn.Execute(sqlInfo)
    If Not (rsInfo.BOF And rsInfo.EOF) Then
        If PE_Cache.CacheIsEmpty("InfoUrl_" & DataType) Then
            Set rsChannel2 = Conn.Execute("select ChannelID,ChannelDir,StructureType,FileNameType,FileExt_Item,UseCreateHTML from PE_Channel Where ChannelID=" & rsInfo(1) & " and Disabled=" & PE_False)
            If Not (rsChannel2.BOF And rsChannel2.EOF) Then
                ChannelDir = rsChannel2("ChannelDir")
                StructureType = rsChannel2("StructureType")
                FileNameType = rsChannel2("FileNameType")
                FileExtType = rsChannel2("FileExt_Item")
                iUseCreateHTML = rsChannel2("UseCreateHTML")
                CacheTemp = rsChannel2("ChannelID") & "|||" & rsChannel2("ChannelDir") & "|||" & rsChannel2("StructureType") & "|||" & rsChannel2("FileNameType") & "|||" & rsChannel2("FileExt_Item") & "|||" & rsChannel2("UseCreateHTML")
                PE_Cache.SetValue "InfoUrl_" & DataType, CacheTemp
            Else
                strTemp = InfoID
            End If
            rsChannel2.Close
            Set rsChannel2 = Nothing
        Else
            ChannelTemp = Split(PE_Cache.GetValue("InfoUrl_" & DataType), "|||")
            If rsInfo(1) = ChannelTemp(0) Then
                ChannelDir = ChannelTemp(1)
                StructureType = ChannelTemp(2)
                FileNameType = ChannelTemp(3)
                FileExtType = ChannelTemp(4)
                iUseCreateHTML = ChannelTemp(5)
            Else
                Set rsChannel2 = Conn.Execute("select ChannelID,ChannelDir,StructureType,FileNameType,FileExt_Item,UseCreateHTML from PE_Channel Where ChannelID=" & rsInfo(1) & " and Disabled=" & PE_False)
                If Not (rsChannel2.BOF And rsChannel2.EOF) Then
                    ChannelDir = rsChannel2("ChannelDir")
                    StructureType = rsChannel2("StructureType")
                    FileNameType = rsChannel2("FileNameType")
                    FileExtType = rsChannel2("FileExt_Item")
                    iUseCreateHTML = rsChannel2("UseCreateHTML")
                    CacheTemp = rsChannel2("ChannelID") & "|||" & rsChannel2("ChannelDir") & "|||" & rsChannel2("StructureType") & "|||" & rsChannel2("FileNameType") & "|||" & rsChannel2("FileExt_Item") & "|||" & rsChannel2("UseCreateHTML")
                    PE_Cache.SetValue "InfoUrl_" & DataType, CacheTemp
                Else
                    strTemp = InfoID
                End If
                rsChannel2.Close
                Set rsChannel2 = Nothing
            End If
        End If
        If strTemp <> InfoID Then
            Select Case OutType
            Case 1
                ChannelUrl = strInstallDir & ChannelDir
                If Enable_SubDomain = True And rsInfo("ChannelID")>0 Then
                    ChannelUrl = Conn.Execute("select LinkUrl from PE_Channel where ChannelID="&rsInfo("ChannelID"))(0)
                    If IsNull(ChannelUrl) Or Trim(ChannelUrl) = "" Or Left(strInstallDir, 7) <> "http://" Then
                        ChannelUrl = strInstallDir & ChannelDir
                    Else
                        ChannelUrl = ChannelUrl
                    End If 									
                End If
                If iUseCreateHTML > 0 Then
                    If DataType = "Product" Then
                        strTemp = ChannelUrl & GetItemPath(StructureType, rsInfo(7), rsInfo(6), rsInfo(4)) & GetItemFileName(FileNameType, ChannelDir, rsInfo(4), InfoID) & arrFileExt(FileExtType)
                    Else
                        If (rsInfo(8) = 0 And rsInfo(5) = 0) Or (rsInfo(2) = -1 And rsInfo(5) = 0) Then
                            strTemp = ChannelUrl & GetItemPath(StructureType, rsInfo(7), rsInfo(6), rsInfo(4)) & GetItemFileName(FileNameType, ChannelDir, rsInfo(4), InfoID) & arrFileExt(FileExtType)
                        Else
                            strTemp = ChannelUrl & "/Show" & DataType & ".asp?" & DataType & "ID=" & rsInfo(0)
                        End If
                    End If
                Else
                    strTemp = ChannelUrl & "/Show" & DataType & ".asp?" & DataType & "ID=" & rsInfo(0)
                End If
            Case 2
                strTemp = rsInfo(3)
            Case 3
                If iUseCreateHTML > 0 Then
                    If DataType = "Product" Then
                        strTemp = "<a href='" & strInstallDir & ChannelDir & GetItemPath(StructureType, rsInfo(7), rsInfo(6), rsInfo(4)) & GetItemFileName(FileNameType, ChannelDir, rsInfo(4), InfoID) & arrFileExt(FileExtType) & "'>" & rsInfo(3) & "</a>"
                    Else
                        If (rsInfo(8) = 0 And rsInfo(5) = 0) Or (rsInfo(2) = -1 And rsInfo(5) = 0)  Then
                            strTemp = "<a href='" & strInstallDir & ChannelDir & GetItemPath(StructureType, rsInfo(7), rsInfo(6), rsInfo(4)) & GetItemFileName(FileNameType, ChannelDir, rsInfo(4), InfoID) & arrFileExt(FileExtType) & "'>" & rsInfo(3) & "</a>"
                        Else
                            strTemp = "<a href='" & strInstallDir & ChannelDir & "/Show" & DataType & ".asp?" & DataType & "ID=" & rsInfo(0) & "'>" & rsInfo(3) & "</a>"
                        End If
                    End If
                Else
                    strTemp = "<a href='" & strInstallDir & ChannelDir & "/Show" & DataType & ".asp?" & DataType & "ID=" & rsInfo(0) & "'>" & rsInfo(3) & "</a>"
                End If
            Case Else
                strTemp = "标签参数错误"
            End Select
        End If
    End If
    rsInfo.Close
    Set rsInfo = Nothing
    GetInfoUrl = strTemp
End Function

'==================================================
'函数名：GetInfoClass
'作  用：获取对象的分类
'参  数：InfoID ------对象ID
'      ：DataType ------数据库名称
'==================================================
Function GetInfoClass(InfoID, OutType)
    If IsNull(InfoID) = True Or IsNull(OutType) = True Then
        GetInfoClass = ""
        Exit Function
    End If
    Dim sqlInfo, rsInfo, rsChannel2, strTemp, PriChannelID
    Dim ChannelDir, ModuleType, StructureType, ListFileType, FileExtList, iUseCreateHTML
    sqlInfo = "select top 1 ClassID,ChannelID,ClassName,ClassDir,ParentDir,ClassPurview from PE_Class Where ClassID=" & PE_CLng(InfoID)
    Set rsInfo = Conn.Execute(sqlInfo)
    If Not (rsInfo.BOF And rsInfo.EOF) Then
        If rsInfo("ChannelID") <> PriChannelID Then
            Set rsChannel2 = Conn.Execute("select ChannelID,ChannelDir,ModuleType,StructureType,ListFileType,FileExt_List,UseCreateHTML from PE_Channel Where ChannelID=" & rsInfo("ChannelID") & " and Disabled=" & PE_False)
            If Not (rsChannel2.BOF And rsChannel2.EOF) Then
                ChannelDir = rsChannel2("ChannelDir")
                ModuleType = rsChannel2("ModuleType")
                StructureType = rsChannel2("StructureType")
                ListFileType = rsChannel2("ListFileType")
                FileExtList = rsChannel2("FileExt_List")
                iUseCreateHTML = rsChannel2("UseCreateHTML")
                PriChannelID = rsInfo("ChannelID")
            Else
                strTemp = "栏目不存在"
            End If
            rsChannel2.Close
            Set rsChannel2 = Nothing
        End If

        If strTemp <> "栏目不存在" Then
            Select Case OutType
            Case 1
                If iUseCreateHTML = 1 Or iUseCreateHTML = 3 Then
                    If ModuleType = 5 Then
                        strTemp = ChannelDir & GetListPath(StructureType, ListFileType, rsInfo("ParentDir"), rsInfo("ClassDir")) & GetListFileName(ListFileType, rsInfo("ClassID"), 1, 1) & arrFileExt(FileExtList)
                    Else
                        If rsInfo("ClassPurview") < 2 Then
                            strTemp = ChannelDir & GetListPath(StructureType, ListFileType, rsInfo("ParentDir"), rsInfo("ClassDir")) & GetListFileName(ListFileType, rsInfo("ClassID"), 1, 1) & arrFileExt(FileExtList)
                        Else
                            strTemp = ChannelDir & "/ShowClass.asp?ClassID=" & rsInfo("ClassID")
                        End If
                    End If
                Else
                    strTemp = ChannelDir & "/ShowClass.asp?ClassID=" & rsInfo("ClassID")
                End If
            Case 2
                strTemp = rsInfo("ClassName")
            Case 3
                If iUseCreateHTML = 1 Or iUseCreateHTML = 3 Then
                    If ModuleType = 5 Then
                        strTemp = "<a href='" & strInstallDir & ChannelDir & GetListPath(StructureType, ListFileType, rsInfo("ParentDir"), rsInfo("ClassDir")) & GetListFileName(ListFileType, rsInfo("ClassID"), 1, 1) & arrFileExt(FileExtList) & "'>" & rsInfo("ClassName") & "</a>"
                    Else
                        If rsInfo("ClassPurview") < 2 Then
                            strTemp = "<a href='" & strInstallDir & ChannelDir & GetListPath(StructureType, ListFileType, rsInfo("ParentDir"), rsInfo("ClassDir")) & GetListFileName(ListFileType, rsInfo("ClassID"), 1, 1) & arrFileExt(FileExtList) & "'>" & rsInfo("ClassName") & "</a>"
                        Else
                            strTemp = "<a href='" & strInstallDir & ChannelDir & "/ShowClass.asp?ClassID=" & rsInfo("ClassID") & "'>" & rsInfo("ClassName") & "</a>"
                        End If
                    End If
                Else
                    strTemp = "<a href='" & strInstallDir & ChannelDir & "/ShowClass.asp?ClassID=" & rsInfo("ClassID") & "'>" & rsInfo("ClassName") & "</a>"
                End If
            Case Else
                strTemp = "标签参数错"
            End Select
            GetInfoClass = strTemp
        Else
            GetInfoClass = ""
        End If
    End If
    rsInfo.Close
    Set rsInfo = Nothing
End Function

'==================================================
'函数名：GetInfoSpecil
'作  用：获取对象的专题
'参  数：InfoID ------对象ID
'      ：DataType ------数据库名称
'==================================================
Function GetInfoSpecil(InfoID, OutType)
    If IsNull(InfoID) = True Or IsNull(OutType) = True Then
        GetInfoSpecil = ""
        Exit Function
    End If
    Dim sqlInfo, rsInfo, rsChannel2, strTemp, PriChannelID
    Dim ChannelDir, iUseCreateHTML
    sqlInfo = "select top 1 A.ChannelID,I.SpecialID,SP.SpecialName,SP.SpecialDir from PE_Article A right join (PE_InfoS I left join PE_Special SP on I.SpecialID=SP.SpecialID) on A.ArticleID=I.ItemID Where A.ArticleID=" & PE_CLng(InfoID)
    Set rsInfo = Conn.Execute(sqlInfo)
    If Not (rsInfo.BOF And rsInfo.EOF) Then
        If rsInfo(0) <> PriChannelID Then
            Set rsChannel2 = Conn.Execute("select ChannelID,ChannelDir,UseCreateHTML from PE_Channel Where ChannelID=" & rsInfo(0) & " and Disabled=" & PE_False)
            If Not (rsChannel2.BOF And rsChannel2.EOF) Then
                ChannelDir = rsChannel2("ChannelDir")
                iUseCreateHTML = rsChannel2("UseCreateHTML")
                PriChannelID = rsInfo(0)
            Else
                strTemp = "专题不存在"
            End If
            rsChannel2.Close
            Set rsChannel2 = Nothing
        End If

        If strTemp <> "专题不存在" Then
            Select Case OutType
            Case 1
                If iUseCreateHTML = 1 Or iUseCreateHTML = 3 Then
                    strTemp = ChannelDir & "/" & rsInfo(3) & "Index.html"
                Else
                    strTemp = ChannelDir & "/ShowSpecial.asp?SpecialID=" & rsInfo(1)
                End If
            Case 2
                strTemp = rsInfo(2)
            Case 3
                If iUseCreateHTML = 1 Or iUseCreateHTML = 3 Then
                    strTemp = "<a href='" & strInstallDir & ChannelDir & "/" & rsInfo(3) & "Index.html" & "'>" & rsInfo(2) & "</a>"
                Else
                    strTemp = "<a href='" & strInstallDir & ChannelDir & "/ShowSpecial.asp?SpecialID=" & rsInfo(1) & "'>" & rsInfo(2) & "</a>"
                End If
            Case Else
                strTemp = "标签参数错"
            End Select
            GetInfoSpecil = strTemp
        Else
            GetInfoSpecil = ""
        End If
    End If
    rsInfo.Close
    Set rsInfo = Nothing
End Function



Function GetSiteCountAll()
    Dim sqlCount, rsCount, iCount, strCount
    If PE_Cache.CacheIsEmpty("SiteCountAll") Then
        sqlCount = "select ChannelName,ChannelShortName,ItemCount,ChannelItemUnit,ModuleType from PE_Channel where ChannelType<=1 and ChannelID<>4  and ChannelID<>997 and Disabled=" & PE_False & " order by OrderID"
        Set rsCount = Conn.Execute(sqlCount)
        Do While Not rsCount.EOF
            If IsNull(rsCount("ItemCount")) Then
                iCount = 0
            Else
                iCount = rsCount("ItemCount")
            End If
            strCount = strCount & (rsCount("ChannelName") & "：" & iCount & " " & rsCount("ChannelItemUnit") & rsCount("ChannelShortName") & "<br>")
            rsCount.MoveNext
        Loop
        rsCount.Close

        sqlCount = "select count(UserID) from PE_User"
        Set rsCount = Conn.Execute(sqlCount)
        strCount = strCount & Replace(XmlText("Site", "SiteCountAll", "注册会员：{$Count}位"), "{$Count}", rsCount(0)) & "<br>"
        rsCount.Close
        Set rsCount = Nothing
        PE_Cache.SetValue "SiteCountAll", strCount
    Else
        strCount = PE_Cache.GetValue("SiteCountAll")
    End If
    GetSiteCountAll = strCount
End Function


'==================================================
'过程名：GetMenuJS
'作  用：生成下拉菜单相关的JS代码
'参  数：无
'==================================================
Function GetMenuJS(sChannelDir, ShowClassTreeGuide)
    Dim strMenu
    strMenu = "<script language='JavaScript' type='text/JavaScript' src='" & strInstallDir & "js/menu.js'></script>" & vbCrLf
    If ChannelID > 0 And ChannelID <> 4 Then
        '无限级下拉菜单的JS代码文件
        strMenu = strMenu & "<script type='text/javascript' language='JavaScript1.2' src='" & strInstallDir & "js/stm31.js'></script>"
        If ShowClassTreeGuide = True Then
            strMenu = strMenu & "<script language='JavaScript' type='text/JavaScript' src='" & strInstallDir & "js/TreeGuide.js'></script>" & vbCrLf
            strMenu = strMenu & "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
            strMenu = strMenu & "//树形导航的JS代码" & vbCrLf
            strMenu = strMenu & "var expandState = 0;" & vbCrLf
            strMenu = strMenu & "function expand(){" & vbCrLf
            strMenu = strMenu & "  if(expandState == 0){setPace('master', 0, 10, 10); if(ie){document.menutop.src = '" & strInstallDir & "images/menui.gif'}; expandState = 1;}" & vbCrLf
            strMenu = strMenu & "  else{setPace('master', -200, 10, 10); if(ie){document.menutop.src='" & strInstallDir & "images/menuo.gif'}; expandState = 0;}" & vbCrLf
            strMenu = strMenu & "}" & vbCrLf
            strMenu = strMenu & "document.write(""<style type=text/css>#master {LEFT: -200px; POSITION: absolute; TOP: 25px; VISIBILITY: visible; Z-INDEX: 999}</style>"")" & vbCrLf
            strMenu = strMenu & "document.write(""<table id=master width='218' border='0' cellspacing='0' cellpadding='0'><tr><td><img border=0 height=6 src=" & strInstallDir & "images/menutop.gif  width=200></td><td rowspan='2' valign='top'><img id=menu onMouseOver=javascript:expand() border=0 height=70 name=menutop src=" & strInstallDir & "images/menuo.gif width=18></td></tr>"");" & vbCrLf
            strMenu = strMenu & "document.write(""<tr><td valign='top'><table width='100%' border='0' cellspacing='5' cellpadding='0'><tr><td height='400' valign='top'><table width=100% height='100%' border=1 cellpadding=0 cellspacing=5 bordercolor='#666666' bgcolor=#ecf6f5 style=FILTER: alpha(opacity=90)><tr>"");" & vbCrLf
            strMenu = strMenu & "document.write(""<td height='10' align='center' bordercolor='#ecf6f5'><font color=999900><strong>" & XmlText("Site", "GetMenuJS", "栏 目 树 形 导 航") & "</strong></font></td></tr><tr><td valign='top' bordercolor='#ecf6f5'>"");" & vbCrLf
            strMenu = strMenu & "document.write(""<iframe width='100%' height='350' src='" & strInstallDir & sChannelDir & "/ClassTree.asp" & "' frameborder=0 allowTransparency='true'></iframe></td></tr></table></td></tr></table></td></tr></table>"");" & vbCrLf
            strMenu = strMenu & "var ie = document.all ? 1 : 0" & vbCrLf
            strMenu = strMenu & "var ns = document.layers ? 1 : 0" & vbCrLf
            strMenu = strMenu & "var master = new Object('element')" & vbCrLf
            strMenu = strMenu & "master.curLeft = -200;   master.curTop = 10;" & vbCrLf
            strMenu = strMenu & "master.gapLeft = 0;      master.gapTop = 0;" & vbCrLf
            strMenu = strMenu & "master.timer = null;" & vbCrLf
            strMenu = strMenu & "if(ie){var sidemenu = document.all.master;}" & vbCrLf
            strMenu = strMenu & "if(ns){var sidemenu = document.master;}" & vbCrLf
            strMenu = strMenu & "setInterval('FixY()',100);" & vbCrLf
            strMenu = strMenu & "</script>" & vbCrLf
        End If
    End If
    GetMenuJS = strMenu
End Function


Function GetLinkType_Option()
    Dim strOption
    strOption = "<select name='JumpType' id='JumpType' onchange=""if(this.options[this.selectedIndex].value!=''){location='Index.asp?LinkType='+this.options[this.selectedIndex].value+'&KindID=" & KindID & "&SpecialID=" & SpecialID & "';}"">"
    strOption = strOption & "<option value='0'"
    If LinkType = 0 Then
        strOption = strOption & " selected"
    End If
    strOption = strOption & ">" & XmlText("Site", "ShowFriendSiteList/t6", "所有类型") & "</option>"
    strOption = strOption & "<option value='1'"
    If LinkType = 1 Then
        strOption = strOption & " selected"
    End If
    strOption = strOption & ">" & XmlText("Site", "ShowFriendSiteList/t4", "LOGO链接") & "</option>"
    strOption = strOption & "<option value='2'"
    If LinkType = 2 Then
        strOption = strOption & " selected"
    End If
    strOption = strOption & ">" & XmlText("Site", "ShowFriendSiteList/t5", "文字链接") & "</option>"
    strOption = strOption & "</select>"
    GetLinkType_Option = strOption
End Function

Function GetFsKind_Option(KindType)
    Dim sqlFsKind, rsFsKind, strOption, FsKindID, strID, strName
    If KindType = 1 Then
        FsKindID = KindID
        strName = "类别"
        strOption = "<select name='JumpKind' id='JumpKind' onchange=""if(this.options[this.selectedIndex].value!=''){location='Index.asp?LinkType=" & LinkType & "&KindID='+this.options[this.selectedIndex].value+'&SpecialID=" & SpecialID & "';}"">"
    ElseIf KindType = 2 Then
        FsKindID = SpecialID
        strName = "专题"
        strOption = "<select name='JumpKind' id='JumpKind' onchange=""if(this.options[this.selectedIndex].value!=''){location='Index.asp?LinkType=" & LinkType & "&KindID=" & KindID & "&SpecialID='+this.options[this.selectedIndex].value;}"">"
    End If
    strOption = strOption & "<option value='0'"
    If FsKindID = "" Then
        strOption = strOption & " selected"
    End If
    strOption = strOption & ">" & XmlText("Site", "ShowFriendSiteList/t6", "所有类型") & strName & "</option>"
    sqlFsKind = "select * from PE_FsKind"
    If KindType > 0 Then
        sqlFsKind = sqlFsKind & " where KindType=" & KindType
    End If
    sqlFsKind = sqlFsKind & " order by KindID"
    Set rsFsKind = Conn.Execute(sqlFsKind)
    Do While Not rsFsKind.EOF
        If rsFsKind("KindID") = FsKindID Then
            strOption = strOption & "<option value='" & rsFsKind("KindID") & "' selected>" & rsFsKind("KindName") & "</option>"
        Else
            strOption = strOption & "<option value='" & rsFsKind("KindID") & "'>" & rsFsKind("KindName") & "</option>"
        End If
        rsFsKind.MoveNext
    Loop
    rsFsKind.Close
    Set rsFsKind = Nothing
    strOption = strOption & "</select>"
    GetFsKind_Option = strOption
End Function

Function SetSearchString(strField)
    Dim arrTemp, i
    Dim strTemp, j
    If Keyword = "" Then
        SetSearchString = ""
        Exit Function
    End If
    
    strTemp = " And ("
    arrTemp = Split(Keyword, ",")
    If UBound(arrTemp) > 2 Then
        j = 2
    Else
        j = UBound(arrTemp)
    End If
    For i = 0 To j
        If i = 0 Then
            If strField = "Keyword" Then
                strTemp = strTemp & strField & " like '%|" & arrTemp(i) & "|%' "
            Else
                strTemp = strTemp & strField & " like '%" & arrTemp(i) & "%' "
            End If
        Else
            If strField = "Keyword" Then
                strTemp = strTemp & " and " & strField & " like '%|" & arrTemp(i) & "|%' "
            Else
                strTemp = strTemp & " and " & strField & " like '%" & arrTemp(i) & "%' "
            End If
        End If
    Next
    strTemp = strTemp & ")"

    SetSearchString = strTemp
End Function

Function GetResultTitle()
    Dim strTitle
    Dim arrTemp, i, sTemp, j
       
    If Keyword = "" Then
        If Trim(Request.ServerVariables("QUERY_STRING")) <> "" Then
            strTitle = "本次高级搜索结果"
        Else
            strTitle = "所有" & ChannelShortName
        End If
    Else
        Keyword = Replace(Replace(Keyword, " ", ","), ",,", ",")
        sTemp = Replace(Keyword, ",", " 和 ")
                
        Select Case strField
        Case "ProductNum"
            strTitle = "编号含有 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
        Case "Title"
            strTitle = "标题含有 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
        Case "ArticleID"
            strTitle = "ID为 <font color=red>" & sTemp & "</font> 的" & ChannelShortName			
        Case "Content"
            strTitle = "内容含有 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
        Case "SoftName"
            strTitle = "名称含有 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
        Case "SoftIntro"
            strTitle = "简介含有 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
        Case "PhotoName"
            strTitle = "名称含有 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
        Case "PhotoIntro"
            strTitle = "简介含有 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
        Case "Author"
            strTitle = "作者姓名中含有 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
        Case "Inputer"
            strTitle = "<font color=red>" & sTemp & "</font> 录入的" & ChannelShortName
        Case "ProductName"
            strTitle = "名称含有 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
        Case "ProductIntro"
            strTitle = "简介含有 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
        Case "ProductExplain"
            strTitle = "介绍含有 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
        Case "ProducerName"
            strTitle = "厂商为 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
        Case "TrademarkName"
            strTitle = "品牌/商标为 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
        Case "Keywords"
            strTitle = "关键字含有 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
        Case Else
            Dim rsField
            Set rsField = Conn.Execute("select Title from PE_Field where (ChannelID=-1 or ChannelID=" & ChannelID & ") and FieldName='" & ReplaceBadChar(strField) & "'")
            If rsField.BOF And rsField.EOF Then
                strTitle = "标题含有 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
            Else
                strTitle = rsField(0) & "中含有 <font color=red>" & sTemp & "</font> 的" & ChannelShortName
            End If
            rsField.Close
            Set rsField = Nothing
        End Select
    End If
    GetResultTitle = strTitle
End Function

Function GetValidConsumeLogID(iUserName, iModuleType, InfoID, iChargeType, PitchTime, ReadTimes)
    Dim trs
    Select Case PE_CLng(iChargeType)
    Case 0  '不重复收费
        Set trs = Conn.Execute("select top 1 LogID from PE_ConsumeLog where UserName='" & iUserName & "' and ModuleType=" & iModuleType & " and InfoID=" & InfoID & " and Income_Payout=2 order by LogID desc")
    Case 1  '距离上次收费时间 N 小时后重新收费
        Set trs = Conn.Execute("select top 1 LogID from PE_ConsumeLog where UserName='" & iUserName & "' and ModuleType=" & iModuleType & " and InfoID=" & InfoID & " and Income_Payout=2 and DateDiff(" & PE_DatePart_H & ",LogTime," & PE_Now & ")<" & PitchTime & " order by LogID desc")
    Case 2  '会员重复查看此文章 N 次后重新收费
        Set trs = Conn.Execute("select top 1 LogID from PE_ConsumeLog where UserName='" & iUserName & "' and ModuleType=" & iModuleType & " and InfoID=" & InfoID & " and Income_Payout=2 and Times<" & ReadTimes & " order by LogID desc")
    Case 3  '上述两者都满足时重新收费
        Set trs = Conn.Execute("select top 1 LogID from PE_ConsumeLog where UserName='" & iUserName & "' and ModuleType=" & iModuleType & " and InfoID=" & InfoID & " and Income_Payout=2 and (DateDiff(" & PE_DatePart_H & ",LogTime," & PE_Now & ")<" & PitchTime & " or Times<" & ReadTimes & ") order by LogID desc")
    Case 4  '上述两者任一个满足时就重新收费
        Set trs = Conn.Execute("select top 1 LogID from PE_ConsumeLog where UserName='" & iUserName & "' and ModuleType=" & iModuleType & " and InfoID=" & InfoID & " and Income_Payout=2 and (DateDiff(" & PE_DatePart_H & ",LogTime," & PE_Now & ")<" & PitchTime & " and Times<" & ReadTimes & ") order by LogID desc")
    Case 5  '每阅读一次就重复收费一次
        Set trs = Conn.Execute("select top 1 LogID from PE_ConsumeLog where 1=0 order by LogID desc")
    End Select
    If trs.BOF And trs.EOF Then
        GetValidConsumeLogID = 0
    Else
        GetValidConsumeLogID = trs(0)
    End If
    Set trs = Nothing
End Function


'*******************************************************
'函 数 名：GetVoteOfContent()
'参    数：无
'作    用：返回投票标签的内容
'**********************************************************
Function GetVoteOfContent(iItemID)
    If IsNull(iItemID) Then
        GetVote = ""
        Exit Function
    End If
    Dim rsVote, rsVote2, strtmp, i
    Set rsVote = Conn.Execute("select top 1 VoteID from " & SheetName & " where " & ModuleName & "ID=" & PE_CLng(iItemID))
    If IsNull(rsVote("VoteID")) Or rsVote("VoteID") = "" Or rsVote("VoteID") = 0 Then
        GetVoteOfContent = ""
    Else
        Set rsVote2 = Conn.Execute("select top 1 * from PE_Vote where ID=" & rsVote("VoteID"))
        If rsVote2.BOF And rsVote2.EOF Then
            GetVoteOfContent = ""
        Else
            If Now() > rsVote2("EndTime") Then
                GetVoteOfContent = "<a href='" & strInstallDir & "Vote.asp?ID=" & rsVote("VoteID") & "&Action=Show' target='_blank'>本调查已过期,点击查看结果</a>"
            Else
                If rsVote2("VoteType") = "Single" Then
                    For i = 1 To 8
                        If Trim(rsVote2("Select" & i) & "") = "" Then Exit For
                        strtmp = strtmp & "<input type='radio' name='VoteOption' value='" & i & "' style='border:0'>" & rsVote2("Select" & i) & "<br>"
                    Next
                Else
                    For i = 1 To 8
                        If Trim(rsVote2("Select" & i) & "") = "" Then Exit For
                        strtmp = strtmp & "<input type='checkbox' name='VoteOption' value='" & i & "' style='border:0'>" & rsVote2("Select" & i) & "<br>"
                    Next
                End If
                GetVoteOfContent = Replace(Replace(Replace(Replace(Replace(XmlText("Article", "ShowVote", "<hr><form name='VoteForm' method='post' action='{$strInstallDir}vote.asp' target='_blank'><table><tr><td><h4>您对<font color=red>""{$Title}""</font>的看法是</h2></tr><tr><td>{$VoteBody}<input name='VoteType' type='hidden'value='{$VoteType}'><input name='Action' type='hidden' value='Vote'><input name='ID' type='hidden' value='{$ID}'></td></tr><tr align='center'><td><a href='javascript:VoteForm.submit();'><img src='{$strInstallDir}images/voteSubmit.gif' width='52' height='18' border='0'></a>&nbsp;&nbsp;<a href='{$strInstallDir}Vote.asp?ID={$ID}&Action=Show' target='_blank'><img src='{$strInstallDir}images/voteView.gif' width='52' height='18' border='0'></a></td></tr></table></form>"), "{$strInstallDir}", strInstallDir), "{$Title}", rsVote2("Title")), "{$VoteBody}", strtmp), "{$VoteType}", rsVote2("VoteType")), "{$ID}", rsVote2("ID"))
            End If
        End If
    End If
    Set rsVote = Nothing
    Set rsVote2 = Nothing
End Function

'**************************************************
'函数名：自设内容
'作  用：创建文件夹
'参  数：foldername ----文件夹名
'返回值：True  ----已创建
'**************************************************
Function CustomContent(ByVal LabelType, ByVal Custom_Content, ByVal strHtml)
    Dim arrCustom, i
    If IsNull(Custom_Content) = True Or Custom_Content = "" Then
        For i = 1 To 20
            strHtml = PE_Replace(strHtml, "{$" & LabelType & "_Custom_Content" & i & "}", "")
        Next
    Else
        arrCustom = Split(Custom_Content, "{#$$$#}")
        For i = 0 To UBound(arrCustom)
            strHtml = PE_Replace(strHtml, "{$" & LabelType & "_Custom_Content" & i + 1 & "}", arrCustom(i))
        Next
        For i = UBound(arrCustom) To 20
            strHtml = PE_Replace(strHtml, "{$" & LabelType & "_Custom_Content" & i & "}", "")
        Next
    End If
    CustomContent = strHtml
End Function

Function GetTDWidth_Date(DateType)
    Select Case DateType
    Case 0  '不显示
        GetTDWidth_Date = 0
    Case 1  '2006-09-11
        GetTDWidth_Date = 60
    Case 2  '9月11日
        GetTDWidth_Date = 50
    Case 3  '09-11
        GetTDWidth_Date = 40
    Case 4  '2006年9月11日
        GetTDWidth_Date = 80
    Case 5  '2006-9-11 10:20:30
        GetTDWidth_Date = 120
    End Select
End Function


'**************************************************
'函数名：ShowPage_Html
'作  用：显示“上一页 下一页”等信息
'参  数：strPath ----HTMl文件的路径
'        iClassID  ----栏目ID
'        FileExt ----- 扩展名
'        sfilename  ---- 文件名
'        TotalNumber ----总数量
'        MaxPerPage  ----每页数量
'        ShowTotal   ----是否显示总数量
'        ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'        strUnit     ----计数单位
'返回值：“上一页 下一页”等信息的HTML代码
'**************************************************
Function ShowPage_Html(ByVal strPath, iClassID, FileExt, sfilename, totalnumber, MaxPerPage, CurrentPage, ShowTotal, ShowAllPages, strUnit)
    Dim NextPage, PrevPage, EndPage
    Dim TotalPage, strTemp, strUrl, i
    If totalnumber = 0 Or MaxPerPage = 0 Or IsNull(MaxPerPage) Then
        ShowPage_Html = ""
        Exit Function
    End If
    If totalnumber Mod MaxPerPage = 0 Then
        TotalPage = totalnumber \ MaxPerPage
    Else
        TotalPage = totalnumber \ MaxPerPage + 1
    End If
    If CurrentPage > TotalPage Then CurrentPage = TotalPage
    PrevPage = TotalPage - CurrentPage + 2
    NextPage = TotalPage - CurrentPage
    EndPage = 1
    If sfilename <> "" Then
        strUrl = JoinChar(sfilename)
    Else
        strUrl = ""
    End If
    If Right(strPath, 1) = "/" Then
        strPath = Left(strPath, Len(strPath) - 1)
    End If
    strTemp = strTemp & "<div class=""showpage"">"
    If ShowTotal = True Then
        strTemp = strTemp & "共 <b>" & totalnumber & "</b> "& strUnit & "&nbsp;&nbsp;"
    End If
    If CurrentPage = 1 Then
        strTemp = strTemp & "首页 | 上一页 |"
    Else
        If iClassID > 0 Then
            strTemp = strTemp & "<a href='" & strPath & "/List_" & iClassID & FileExt & "'>首页</a> |"
        Else
            strTemp = strTemp & "<a href='" & strPath & "/" & "'>首页</a> |"
        End If
        If CurrentPage = 2 Then
            If iClassID > 0 Then
                strTemp = strTemp & " <a href='" & strPath & "/List_" & iClassID & FileExt & "'>上一页</a> |"
            Else
                strTemp = strTemp & " <a href='" & strPath & "/" & "'>上一页</a> |"
            End If
        Else
            If strUrl <> "" Then
                strTemp = strTemp & " <a href='" & strUrl & "page=" & (CurrentPage - 1) & "'>上一页</a> |"
            Else
                If iClassID > 0 Then
                    strTemp = strTemp & " <a href='" & strPath & "/List_" & iClassID & "_" & PrevPage & FileExt & "'>上一页</a> |"
                Else
                    strTemp = strTemp & " <a href='" & strPath & "/List_" & PrevPage & FileExt & "'>上一页</a> |"
                End If
            End If
        End If
    End If
    strTemp = strTemp & " "
    If ShowAllPages = True Then
        Dim Jmaxpages
        If (CurrentPage - 4) <= 0 Or TotalPage < 10 Then
            Jmaxpages = 1
            Do While (Jmaxpages < 10)
                If Jmaxpages = CurrentPage Then
                    strTemp = strTemp & "<font color=""FF0000"">" & Jmaxpages & "</font> "
                ElseIf Jmaxpages = 1 Then
                    If iClassID > 0 Then
                        strTemp = strTemp & "<a href=""" & strPath & "/List_" & iClassID & FileExt & """>" & Jmaxpages & "</a> "
                    Else
                        strTemp = strTemp & "<a href=""" & strPath & "/" & """>" & Jmaxpages & "</a> "
                    End If
                Else
                    If strUrl <> "" Then
                        strTemp = strTemp & "<a href=""" & strUrl & "page=" & Jmaxpages & """>" & Jmaxpages & "</a> "
                    Else
                        If iClassID > 0 Then
                            strTemp = strTemp & "<a href=""" & strPath & "/List_" & iClassID & "_" & TotalPage - Jmaxpages + 1 & FileExt & """>" & Jmaxpages & "</a> "
                        Else
                            strTemp = strTemp & "<a href=""" & strPath & "/List_" & TotalPage - Jmaxpages + 1 & FileExt & """>" & Jmaxpages & "</a> "
                        End If
                    End If
                End If
                If Jmaxpages = TotalPage Then Exit Do
                Jmaxpages = Jmaxpages + 1
            Loop
        ElseIf (CurrentPage + 4) >= TotalPage Then
            Jmaxpages = TotalPage - 8
            Do While (Jmaxpages <= TotalPage)
                If Jmaxpages = CurrentPage Then
                    strTemp = strTemp & "<font color=""FF0000"">" & Jmaxpages & "</font> "
                ElseIf Jmaxpages = 1 Then
                    If iClassID > 0 Then
                        strTemp = strTemp & "<a href=""" & strPath & "/List_" & iClassID & FileExt & """>" & Jmaxpages & "</a> "
                    Else
                        strTemp = strTemp & "<a href=""" & strPath & "/" & """>" & Jmaxpages & "</a> "
                    End If
                Else
                    If strUrl <> "" Then
                        strTemp = strTemp & "<a href=""" & strUrl & "page=" & Jmaxpages & """>" & Jmaxpages & "</a> "
                    Else
                        If iClassID > 0 Then
                            strTemp = strTemp & "<a href=""" & strPath & "/List_" & iClassID & "_" & TotalPage - Jmaxpages + 1 & FileExt & """>" & Jmaxpages & "</a> "
                        Else
                            strTemp = strTemp & "<a href=""" & strPath & "/List_" & TotalPage - Jmaxpages + 1 & FileExt & """>" & Jmaxpages & "</a> "
                        End If
                    End If
                End If
                Jmaxpages = Jmaxpages + 1
            Loop
        Else
            Jmaxpages = CurrentPage - 4
            Do While (Jmaxpages < CurrentPage + 5)
                If Jmaxpages = CurrentPage Then
                    strTemp = strTemp & "<font color=""FF0000"">" & Jmaxpages & "</font> "
                ElseIf Jmaxpages = 1 Then
                    If iClassID > 0 Then
                        strTemp = strTemp & "<a href=""" & strPath & "/List_" & iClassID & FileExt & """>" & Jmaxpages & "</a> "
                    Else
                        strTemp = strTemp & "<a href=""" & strPath & "/" & """>" & Jmaxpages & "</a> "
                    End If
                Else
                    If strUrl <> "" Then
                        strTemp = strTemp & "<a href=""" & strUrl & "page=" & Jmaxpages & """>" & Jmaxpages & "</a> "
                    Else
                        If iClassID > 0 Then
                            strTemp = strTemp & "<a href=""" & strPath & "/List_" & iClassID & "_" & TotalPage - Jmaxpages + 1 & FileExt & """>" & Jmaxpages & "</a> "
                        Else
                            strTemp = strTemp & "<a href=""" & strPath & "/List_" & TotalPage - Jmaxpages + 1 & FileExt & """>" & Jmaxpages & "</a> "
                        End If
                    End If
                End If
                Jmaxpages = Jmaxpages + 1
            Loop
        End If
    End If

    If CurrentPage >= TotalPage Then
        strTemp = strTemp & "| 下一页 | 尾页  "
    Else
        If strUrl <> "" Then
            strTemp = strTemp & "| <a href='" & strUrl & "page=" & (CurrentPage + 1) & "'>下一页</a> "
            strTemp = strTemp & "| <a href='" & strUrl & "page=" & TotalPage & "'>尾页  </a>"
        Else
            If iClassID > 0 Then
                strTemp = strTemp & "| <a href='" & strPath & "/List_" & iClassID & "_" & NextPage & FileExt & "'>下一页</a> "
                strTemp = strTemp & "| <a href='" & strPath & "/List_" & iClassID & "_" & EndPage & FileExt & "'>尾页  </a>"
            Else
                strTemp = strTemp & "| <a href='" & strPath & "/List_" & NextPage & FileExt & "'>下一页</a> "
                strTemp = strTemp & "| <a href='" & strPath & "/List_" & EndPage & FileExt & "'>尾页  </a>"
            End If
        End If
    End If
	strTemp = strTemp & "&nbsp;<b>" & MaxPerPage & "</b>" & strUnit & "/页"
    If ShowAllPages = True Then
        strTemp = strTemp & "&nbsp;&nbsp;转到第<Input type='text' name='page' size='3' maxlength='5' value='" & CurrentPage & "' onKeyPress=""gopage(this.value,"&TotalPage&")"">页"
        strTemp = strTemp & "<script language='javascript'>" & vbCrLf
        strTemp = strTemp & "function gopage(page,totalpage){" & vbCrLf
        strTemp = strTemp & "  if (event.keyCode==13){" & vbCrLf
        strTemp = strTemp & "    if(Math.abs(page)>totalpage) page=totalpage;" & vbCrLf
        If iClassID > 0 Then
            If strUrl <> "" Then
                strTemp = strTemp & "    if(Math.abs(page)>1) window.location='" & strUrl & "page='" & "+Math.abs(page);" & vbCrLf
            Else
                strTemp = strTemp & "    if(Math.abs(page)>1) window.location='" & strPath & "/List_" & iClassID & "_'" & "+(totalpage-Math.abs(page)+1)+'" & FileExt & "';" & vbCrLf
            End If
            strTemp = strTemp & "    else  window.location='" & strPath & "/List_" & iClassID & FileExt & "';" & vbCrLf
        Else
            If strUrl <> "" Then
                strTemp = strTemp & "    if(Math.abs(page)>1) window.location='" & strUrl & "page='" & "+Math.abs(page);" & vbCrLf
            Else
                strTemp = strTemp & "    if(Math.abs(page)>1) window.location='" & strPath & "/List_'+(totalpage-Math.abs(page)+1)+'" & FileExt & "';" & vbCrLf
            End If
            strTemp = strTemp & "    else  window.location='" & strPath & "/Index" & FileExt & "';" & vbCrLf
        End If
        strTemp = strTemp & "  }" & vbCrLf
        strTemp = strTemp & "}" & vbCrLf
        strTemp = strTemp & "</script>" & vbCrLf
    End If
    strTemp = strTemp & "</div>" & vbCrLf
   ShowPage_Html = strTemp
End Function


'**************************************************
'函数名：ShowPage_en_Html
'作  用：显示英文“上一页 下一页”等信息
'参  数：strPath ----HTMl文件的路径
'        iClassID  ----栏目ID
'        FileExt ----- 扩展名
'        sfilename  ---- 文件名
'        TotalNumber ----总数量
'        MaxPerPage  ----每页数量
'        ShowTotal   ----是否显示总数量
'        ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'        strUnit     ----计数单位
'返回值：“上一页 下一页”等信息的HTML代码
'**************************************************
Function ShowPage_en_Html(ByVal strPath, iClassID, FileExt, sfilename, totalnumber, MaxPerPage, CurrentPage, ShowTotal, ShowAllPages, strUnit)
    Dim NextPage, PrevPage, EndPage
    Dim TotalPage, strTemp, strUrl, i
    If totalnumber = 0 Or MaxPerPage = 0 Or IsNull(MaxPerPage) Then
        ShowPage_en_Html = ""
        Exit Function
    End If
    If totalnumber Mod MaxPerPage = 0 Then
        TotalPage = totalnumber \ MaxPerPage
    Else
        TotalPage = totalnumber \ MaxPerPage + 1
    End If
    If CurrentPage > TotalPage Then CurrentPage = TotalPage
    
    PrevPage = TotalPage - CurrentPage + 2
    NextPage = TotalPage - CurrentPage
    EndPage = 1

    If sfilename <> "" Then
        strUrl = JoinChar(sfilename)
    Else
        strUrl = ""
    End If
    
    If Right(strPath, 1) = "/" Then
        strPath = Left(strPath, Len(strPath) - 1)
    End If
    
    strTemp = "<!-- ShowPage Begin -->"
    strTemp = strTemp & "<div class=""show_page"">"
    If ShowTotal = True Then
        strTemp = strTemp & "Total <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
    End If
    If CurrentPage = 1 Then
        strTemp = strTemp & "FirstPage PreviousPage&nbsp;"
    Else
        If iClassID > 0 Then
            strTemp = strTemp & "<a href='" & strPath & "/List_" & iClassID & FileExt & "'>FirstPage</a>&nbsp;"
        Else
            strTemp = strTemp & "<a href='" & strPath & "/Index" & FileExt & "'>FirstPage</a>&nbsp;"
        End If
        If CurrentPage = 2 Then
            If iClassID > 0 Then
                strTemp = strTemp & "<a href='" & strPath & "/List_" & iClassID & FileExt & "'>PreviousPage</a>&nbsp;"
            Else
                strTemp = strTemp & "<a href='" & strPath & "/Index" & FileExt & "'>PreviousPage</a>&nbsp;"
            End If
        Else
            If strUrl <> "" Then
                strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage - 1) & "'>PreviousPage</a>&nbsp;"
            Else
                If iClassID > 0 Then
                    strTemp = strTemp & "<a href='" & strPath & "/List_" & iClassID & "_" & PrevPage & FileExt & "'>PreviousPage</a>&nbsp;"
                Else
                    strTemp = strTemp & "<a href='" & strPath & "/List_" & PrevPage & FileExt & "'>PreviousPage</a>&nbsp;"
                End If
            End If
        End If
    End If

    If CurrentPage >= TotalPage Then
        strTemp = strTemp & "NextPage LastPage"
    Else
        If strUrl <> "" Then
            strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage + 1) & "'>NextPage</a>&nbsp;"
            strTemp = strTemp & "<a href='" & strUrl & "page=" & TotalPage & "'>LastPage</a>"
        Else
            If iClassID > 0 Then
                strTemp = strTemp & "<a href='" & strPath & "/List_" & iClassID & "_" & NextPage & FileExt & "'>NextPage</a>&nbsp;"
                strTemp = strTemp & "<a href='" & strPath & "/List_" & iClassID & "_" & EndPage & FileExt & "'>LastPage</a>"
            Else
                strTemp = strTemp & "<a href='" & strPath & "/List_" & NextPage & FileExt & "'>NextPage</a>&nbsp;"
                strTemp = strTemp & "<a href='" & strPath & "/List_" & EndPage & FileExt & "'>LastPage</a>"
            End If
        End If
    End If
    strTemp = strTemp & "&nbsp;CurrentPage:<strong><font color=red>" & CurrentPage & "</font>/" & TotalPage & "</strong> "
    strTemp = strTemp & "&nbsp;<b>" & MaxPerPage & "</b>" & strUnit & "/Page"
    If ShowAllPages = True Then
        If TotalPage > 20 Then
            strTemp = strTemp & "&nbsp;&nbsp;GoTo Page:<Input type='text' name='page' size='3' maxlength='5' value='" & CurrentPage & "' onKeyPress='gopage(this.value);'>"
        Else
            strTemp = strTemp & "&nbsp;Goto:<select name='page' size='1' onchange=""javascript:window.location=this.options[this.selectedIndex].value;"">"
            If iClassID > 0 Then
                strTemp = strTemp & "<option value='" & strPath & "/List_" & iClassID & FileExt & "'>Page1</option>"
            Else
                strTemp = strTemp & "<option value='" & strPath & "/Index" & FileExt & "'>Page1</option>"
            End If
            For i = 2 To TotalPage
                If strUrl <> "" Then
                   strTemp = strTemp & "<option value='" & strUrl & "page=" & i & "'"
                Else
                    If iClassID > 0 Then
                        strTemp = strTemp & "<option value='" & strPath & "/List_" & iClassID & "_" & TotalPage - i + 1 & FileExt & "'"
                    Else
                        strTemp = strTemp & "<option value='" & strPath & "/List_" & TotalPage - i + 1 & FileExt & "'"
                    End If
                End If
                If CurrentPage = i Then strTemp = strTemp & " selected "
                strTemp = strTemp & ">Page" & i & "</option>"
            Next
            strTemp = strTemp & "</select>"
        End If
    End If
    strTemp = strTemp & "</div>" & vbCrLf
    If ShowAllPages = True And TotalPage > 20 Then
        strTemp = strTemp & "<script language='javascript'>" & vbCrLf
        strTemp = strTemp & "function gopage(page){" & vbCrLf
        strTemp = strTemp & "  if (event.keyCode==13){" & vbCrLf
        strTemp = strTemp & "    if(Math.abs(page)>totalpage) page=totalpage;" & vbCrLf
        If iClassID > 0 Then
            If strUrl <> "" Then
                strTemp = strTemp & "    if(Math.abs(page)>1) window.location='" & strUrl & "page='" & "+Math.abs(page);" & vbCrLf
            Else
                strTemp = strTemp & "    if(Math.abs(page)>1) window.location='" & strPath & "/List_" & iClassID & "_'" & "+(totalpage-Math.abs(page)+1)+'" & FileExt & "';" & vbCrLf
            End If
            strTemp = strTemp & "    else  window.location='" & strPath & "/List_" & iClassID & FileExt & "';" & vbCrLf
        Else
            If strUrl <> "" Then
                strTemp = strTemp & "    if(Math.abs(page)>1) window.location='" & strUrl & "page='" & "+Math.abs(page);" & vbCrLf
            Else
                strTemp = strTemp & "    if(Math.abs(page)>1) window.location='" & strPath & "/List_'+(totalpage-Math.abs(page)+1)+'" & FileExt & "';" & vbCrLf
            End If
            strTemp = strTemp & "    else  window.location='" & strPath & "/Index" & FileExt & "';" & vbCrLf
        End If
        strTemp = strTemp & "  }" & vbCrLf
        strTemp = strTemp & "}" & vbCrLf
        strTemp = strTemp & "</script>" & vbCrLf
    End If
    strTemp = strTemp & "<!-- ShowPage End -->"
    ShowPage_en_Html = strTemp
End Function

Function ReplaceSpace(ByVal iText)
    If IsNull(iText) Then
        ReplaceSpace = "未知"
    Else
        ReplaceSpace = iText
    End If
End Function
%>
