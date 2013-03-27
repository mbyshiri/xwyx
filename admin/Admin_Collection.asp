<!--#include file="Admin_Common.asp"-->
<!--#include file="Admin_CommonCode_Collection.asp"-->
<!--#include file="../Include/PowerEasy.FSO.asp"-->
<!--#include file="../Include/PowerEasy.XmlHttp.asp"-->
<!--#include file="../Include/PowerEasy.CreateThumb.asp"-->

<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Const NeedCheckComeUrl = True   '是否需要检查外部访问

Const PurviewLevel = 2      '0--不检查，1--超级管理员，2--普通管理员
Const PurviewLevel_Channel = 0   '0--不检查，1--频道管理员，2--栏目总编，3--栏目管理员
Const PurviewLevel_Others = "Collection"   '其他权限

Private rs, sql, rsItem, i '通用变量
'项目公用变量
Private ItemID, ItemName, ClassID, SpecialID, ItemNum, ListNum, ItemEnd, ListEnd
Private Status, OnTop, Elite, Hot, Hits, Stars, InfoPoint, MaxCharPerPage, ShowCommentLink
Private FilterProperty, Script_Iframe, Script_Object, Script_Script, Script_Class, Script_Div, Script_Span, Script_Img, Script_Font, Script_A, Script_Html, Script_Table, Script_Tr, Script_Td
Private SaveFiles, SaveFlashUrlToFile, CollecOrder, CreateImmediate
Private ListStr, LsString, LoString, ListPaingType, LPsString, LPoString, ListPaingStr1, ListPaingStr2, ListPaingID1, ListPaingID2, ListPaingStr3, HsString, HoString, HttpUrlType, HttpUrlStr
Private TsString, ToString, CsString, CoString, AuthorType, AsString, AoString, AuthorStr, CopyFromType, FsString, FoString
Private CopyFromStr, KeyType, KsString, KoString, KeyStr, KeyScatterNum, NewsPaingType, NPsString, NPoString, NewsPaingStr1, NewsPaingStr2
Private PsString, PoString, PhsString, PhoString
Private IsString, IoString, IntroType, IntroStr, IntroNum, Intro
Private DateType, DsString, DoString
Private IncludePicYn, DefaultPicYn, PaginationType
'自定义字段采集变量
Private IsField, Field, iField
Private arrField, arrField2, FieldID, FieldName, FieldType, FisSting, FioSting, FieldStr
'登录验证变量
Private LoginType, LoginUrl, LoginPostUrl, LoginUser, LoginPass, LoginFalse
'采集选项变量
Private CollecTest, Content_view, IsTitle, IsLink
'采集相关的变量
Private CollecNewsi, CollecNewsj, CollectionModify, ItemIDStr
Private ItemIDArray, ContentTemp, NewsPaingNextCode, NewsPaingNext
Private Arr_j, Arr_i, NewsUrl, NewsCode
Private LoginData, LoginResult, CollecNewsA, OrderTemp, StartTime
'图片类型及保存路径
Private FilesOverStr, FilesPath, FilesArray, ImagesNum
'文章保存变量
Private ArticleID, Title, Content, Author, CopyFrom, Key, UpDateType, UpdateTime, IncludePic, UploadFiles, DefaultPicUrl
'历史记录
Private His_Title, His_NewsCollecDate, His_Result, His_Repeat, His_i
'采集列表处理变量
Private WebUrl, ListUrl, ListCode, ListUrlArray, NewsArrayCode, NewsArray, ListArray, ListPaingNext
Private tempStr, ItemIDtemp, TimeNum, rnd_temp, ArticleList, CollectionNum, CollectionType
Private AddWatermark, AddThumb, ItemSucceedNum, ItemSucceedNum2, ImagesNumAll, PaingNum, dirMonth, dtNow
Private Arr_Item, Arr_Histrolys, CollecType, Arr_Filters, Filteri, FilterStr, SwfTime, CollectionCreateHTML '采集缓存
'采集正文分页变量
Private PageListCode, PageArrayCode, PageArray
'定时生成
Private Timing_AreaCollection, TimingCreate
'收费文章属性
Private InfoPurview, arrGroupID, PitchTime, ReadTimes, DividePercent
'列表缩略图
Private ThumbnailType, ThsString, ThoString
Private ThumbnailArrayCode, ThumbnailArray, ThumbnailUrl
'转换路径
Private ConversionTrails


'采集定时刷新会包含变量http://
If InStr(ComeUrl, "?") > 0 Then
    ComeUrl = Left(ComeUrl, InStr(ComeUrl, "?"))
End If

    
'获得当前时间当前年月
dtNow = Now()
dirMonth = Year(dtNow) & Right("0" & Month(dtNow), 2)


XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))



Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>采集系统</title>" & vbCrLf
Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""Admin_Style.css"">" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">" & vbCrLf

'这些代码暂时放到采集以后会移开
If Action = "CreateItemHtml" Then
Else
    Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""border"">" & vbCrLf
    Call ShowPageTitle("采 集 系 统 项 目 管 理", 10051)
    Response.Write "</table>" & vbCrLf
End If

Select Case Action
Case "Start"                    '开始采集
    Call Start
Case "main"                     '采集管理
    Call main
Case "CheckItem"
    Call CheckItem              '批量检测项目
Case "StopCollection"           '停止采集处理
    ItemEnd = True
    CollecNewsi = PE_CLng(Trim(Request("CollecNewsi")))         'CollecNewsi    显示采集成功数
    CollecNewsj = PE_CLng(Trim(Request("CollecNewsj")))         'CollecNewsj    显示采集失败数
    ArticleList = Replace(CStr(Trim(Request("ArticleList"))),"|","/")            'ArticleList 用于缓存不同的列表
    ItemSucceedNum2 = PE_CLng(Trim(Request("ItemSucceedNum2"))) 'ItemSucceedNum2 成功采集项目数
    ImagesNumAll = PE_CLng(Trim(Request("ImagesNumAll")))       'ImagesNumAll    项目总数
    CollecType = PE_CLng(Trim(Request("CollecType")))           'CollecType    采集模式 0 稳定 1 快速
    CollectionCreateHTML = Trim(Request("CollectionCreateHTML")) 'CollectionCreateHTML    生成html数组
    CreateImmediate = Trim(Request("CreateImmediate"))          'CreateImmediate 采集项目是否生成
    UseCreateHTML = PE_CLng(Trim(Request("UseCreateHTML")))     'UseCreateHTML   频道是否生成

    If CollectionCreateHTML = "" Then
        If CreateImmediate = "True" And UseCreateHTML <> 0 And ItemSucceedNum2 <> 0 Then
            CollectionCreateHTML = PE_CLng(Trim(Request("ChannelID"))) & "$" & PE_CLng(Trim(Request("ClassID"))) & "$" & ReplaceBadChar(Trim(Request("SpecialID"))) & "$" & ItemSucceedNum2
        End If
    Else '如果是多项目网站停止

        If CreateImmediate = "True" And UseCreateHTML <> 0 And ItemSucceedNum2 <> 0 Then
            CollectionCreateHTML = CollectionCreateHTML & "|" & PE_CLng(Trim(Request("ChannelID"))) & "$" & PE_CLng(Trim(Request("ClassID"))) & "$" & ReplaceBadChar(Trim(Request("SpecialID"))) & "$" & ItemSucceedNum2
        End If
    End If

    ErrMsg = "<br>已经停止当前项目,目前已完成！"
    ErrMsg = ErrMsg & "<li>成功采集： <font color=red>" & CollecNewsi & "</font>  篇,失败：<font color=blue> " & CollecNewsj & "</font>  篇,图片：<font color=green>" & ImagesNumAll & "</font> 个。</li>"
    Call PE_Cache.DelAllCache
    Call WriteSuccessMsg2(ErrMsg)
Case "CreateItemHtml"
    Call CreateItemHtml         '采集后自动生成Html
Case Else
    Call main
End Select
Response.Write "</body></html>"
Call CloseConn


'=================================================
'过程名：Main
'作  用：文章采集
'=================================================
Sub main()

    Dim sql, rs, SqlH, RsH, Flag, Action
    Dim iChannelID, ClassID, SpecialID, ItemID, ItemName, ListUrl, WebName, NewsCollecDate
    Dim SkinID, LayoutID, SkinCount, LayoutCount, MaxPerPage
        
    If Request("page") <> "" Then
        CurrentPage = CInt(Request("page"))
    Else
        CurrentPage = 1
    End If
    iChannelID = PE_CLng(Trim(Request("iChannelID")))
    MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))
    If MaxPerPage <= 0 Then MaxPerPage = 10
    
    strFileName = "Admin_Collection.asp?Action=Main&iChannelID=" & iChannelID

    Response.Write "<a name='submit'></a>" & vbCrLf
    Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"">" & vbCrLf
    Response.Write "  <tr class=""tdbg""> " & vbCrLf
    Response.Write "    <td width=""70"" height=""30""><strong>管理导航：</strong></td>" & vbCrLf
    Response.Write "    <td height=""30""><a href=Admin_Collection.asp?Action=Main>管理首页</a> | <a href=""Admin_CollectionManage.asp?Action=Step1"">添加新项目</a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>"
    
    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "    function CheckAll(thisform){" & vbCrLf
    Response.Write "        for (var i=0;i<thisform.elements.length;i++){" & vbCrLf
    Response.Write "            var e = thisform.elements[i];" & vbCrLf
    Response.Write "            if (e.Name != ""chkAll""&&e.disabled!=true&&e.zzz!=1)" & vbCrLf
    Response.Write "                e.checked = thisform.chkAll.checked;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    function mysub(){" & vbCrLf
    Response.Write "        window.location='#submit';" & vbCrLf
    Response.Write "        esave.style.visibility=""visible"";" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "</script>" & vbCrLf
    Response.Write "<br>" & vbCrLf
    
    If IsObjInstalled("MSXML2.XMLHTTP") = False Then
        Call WriteErrMsg("<li>您的系统没有安装XMLHTTP 组件,请到微软网站下载MSXML 4.0", ComeUrl)
        Exit Sub
    End If

    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "<tr class='title'><td colspan='2'> | "
    sql = "SELECT DISTINCT I.ChannelID, C.ChannelName,C.ModuleType FROM PE_Item I LEFT OUTER JOIN PE_Channel C ON I.ChannelID = C.ChannelID"
    sql = sql & " WHERE C.ModuleType=1"
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
    Else
        Do While Not rs.EOF
            Response.Write "<a href='Admin_Collection.asp?Action=Main&iChannelID=" & rs("ChannelID") & "'><FONT style='font-size:12px'"
            If rs("ChannelID") = iChannelID Then Response.Write "color='red'"
            Response.Write ">" & rs("ChannelName") & "</FONT></a> | "
            rs.MoveNext
        Loop
        Response.Write "<a href='Admin_Collection.asp?Action=Main&iChannelID=0'><FONT style='font-size:12px'"
        If iChannelID = 0 Then Response.Write "color='red'"
        Response.Write "> 所有频道 </FONT></a> | "
    End If
    Response.Write "</td></tr>"
    Response.Write "</table>"
    rs.Close
    Set rs = Nothing
    Response.Write "<br>"
    Response.Write GetManagePath(iChannelID)
    Response.Write "<br>"
    Response.Write "<table class=""border"" border=""0"" cellspacing=""1"" width=""100%"" cellpadding=""0"">" & vbCrLf
    Response.Write "<form name=""myform"" method=""POST"" action=""Admin_Collection.asp"">" & vbCrLf
    Response.Write "  <tr class=""title"" style=""padding: 0px 2px;"">" & vbCrLf
    Response.Write "    <td width=""40"" height=""22"" align=""center""><strong>选择</strong></td>        " & vbCrLf
    Response.Write "    <td width=""100"" align=""center""><strong>项目名称</strong></td>" & vbCrLf
    Response.Write "    <td width=""100"" align=""center""><strong>采集地址</strong></td>" & vbCrLf
    Response.Write "    <td width=""100"" height=""22"" align=""center""><strong>所属频道</strong></td> " & vbCrLf
    Response.Write "    <td width=""100"" height=""22"" align=""center""><strong>所属栏目</strong></td> " & vbCrLf
    Response.Write "    <td width=""40"" align=""center""><strong>可运行</strong></td>        " & vbCrLf
    Response.Write "    <td width=""120"" height=""22"" align=""center""><strong>上次采集时间</strong></td>" & vbCrLf
    Response.Write "    <td width=""60"" height=""22"" align=""center""><strong>成功记录</strong></td>" & vbCrLf
    Response.Write "    <td width=""60"" height=""22"" align=""center""><strong>失败记录</strong></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf

    sql = "SELECT I.*,C.ChannelName,CL.ClassName,C.Disabled,C.ModuleType "
    sql = sql & " FROM (PE_Item I left JOIN PE_Channel C ON I.ChannelID =C.ChannelID)"
    sql = sql & " Left JOIN PE_Class CL ON I.ClassID = CL.ClassID"
    sql = sql & " where C.ModuleType=1 and I.Flag=" & PE_True
    If iChannelID <> 0 Then sql = sql & " And I.ChannelID=" & iChannelID
    sql = sql & " ORDER BY I.Flag "
        If SystemDatabaseType = "SQL" Then
        sql = sql & "desc"
    Else
        sql = sql & "asc"
    End If
    sql = sql & ", I.ItemID DESC, I.NewsCollecDate DESC"

    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 1
    If rs.BOF And rs.EOF Then
        Response.Write "<tr class='tdbg' height='50'><td colspan='9' align='center'>系统中暂无采集项目！</td></tr></table>"
    Else
        totalPut = rs.RecordCount
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > totalPut Then
            If (totalPut Mod MaxPerPage) = 0 Then
                CurrentPage = totalPut \ MaxPerPage
            Else
                CurrentPage = totalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                rs.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        Dim VisitorNum
        VisitorNum = 0
        Do While Not rs.EOF
            ChannelID = rs("ChannelID")
            ClassID = PE_CLng(rs("ClassID"))
            ItemID = rs("ItemID")
            ItemName = rs("ItemName")
            ListUrl = rs("ListStr")
            WebName = rs("WebName")
            NewsCollecDate = rs("NewsCollecDate")
            Flag = rs("Flag")
            
            Response.Write "<tr class=""tdbg"" onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbCrLf
            Response.Write "  <td width=""40"" align=""center"" height='25'>" & vbCrLf

            Response.Write "    <input type=""checkbox"" value=" & ItemID & " name=""ItemID"""
            If rs("Disabled") = True Or Flag <> True Or IsNull(rs("ChannelName")) = True Then
                Response.Write " disabled"
            End If
            Response.Write "> " & vbCrLf
            Response.Write "  </td>" & vbCrLf
            Response.Write "  <td width=""100"" align=""center"">" & ItemName & "</td> " & vbCrLf
            Response.Write "  <td width=""100"" align=""center""><a href=" & ListUrl & " target=""_bank"">" & WebName & "</a></td>  " & vbCrLf
            Response.Write "  <td width=""100"" height=""22"" align=""center"">"
            If IsNull(rs("ChannelName")) = True Then
                Response.Write "还没有指定频道"
            Else
                If rs("Disabled") = True Then
                    Response.Write rs("ChannelName") & "<font color=red>&nbsp;已禁用</font>"
                Else
                    Response.Write rs("ChannelName")
                End If
            End If
            Response.Write "</td> " & vbCrLf
            Response.Write "  <td width=""100"" align=""center"">"
            If IsNull(rs("ClassName")) = True Then
                Response.Write "还没有指定栏目"
            Else
                Response.Write rs("ClassName")
            End If
            Response.Write "</td>" & vbCrLf
            Response.Write "  <td width=""40"" align=""center"">" & vbCrLf
            If Flag = True Then
                Response.Write "<b>√</b>"
            Else
                Response.Write "<FONT color='red'><b>×</b></FONT>"
            End If
            Response.Write "  </td>" & vbCrLf
            Response.Write "  <td width=""120"" align=""center"">" & vbCrLf
            If DateDiff("d", NewsCollecDate, Now()) = 0 Then
                Response.Write "<font color=red>" & NewsCollecDate & "</font>"
            Else
                Response.Write NewsCollecDate
            End If
            Response.Write "  </td>" & vbCrLf
            Response.Write "  <td width=""60"" align=""center"">" & vbCrLf
            Response.Write "   <a href='Admin_CollectionHistory.asp?Action=main&SelectCollateItemID=" & ItemID & "&HistrolyResult=true'>&nbsp;"
            Call HistrolyNum(ItemID, PE_True)
            Response.Write "</a>" & vbCrLf
            Response.Write "  </td>" & vbCrLf
            Response.Write "  <td width=""60"" align=""center"">" & vbCrLf
            Response.Write "   <a href='Admin_CollectionHistory.asp?Action=main&SelectCollateItemID=" & ItemID & "&HistrolyResult=false'>&nbsp;"
            Call HistrolyNum(ItemID, PE_False)
            Response.Write "</a>" & vbCrLf
            Response.Write "  </td>" & vbCrLf
            Response.Write "</tr> " & vbCrLf

            VisitorNum = VisitorNum + 1
            If VisitorNum >= MaxPerPage Then Exit Do
            rs.MoveNext
        Loop
        Response.Write "<tr class='tdbg'>" & vbCrLf
        Response.Write "  <td colspan='7' align=""right"">合计：</td>" & vbCrLf
        Response.Write "  <td align='center' width='60'>" & vbCrLf
        SqlH = "select count(HistrolyNewsID) from PE_HistrolyNews where  Result=" & PE_True
        Set RsH = Conn.Execute(SqlH)
        If RsH.BOF And RsH.EOF Then
            Response.Write "&nbsp;<font color='green'>0</font>"
        Else
            Response.Write "&nbsp;<font color='blue'>" & RsH(0) & "</font>"
        End If
        RsH.Close
        Set RsH = Nothing
        Response.Write "  </td>" & vbCrLf
        Response.Write "  <td align='center' width='60'>" & vbCrLf
        SqlH = "select count(HistrolyNewsID) from PE_HistrolyNews where  Result=" & PE_False
        Set RsH = Conn.Execute(SqlH)
        If RsH.BOF And RsH.EOF Then
            Response.Write "&nbsp;<font color='green'>0</font>"
        Else
            Response.Write "&nbsp;<font color='red'>" & RsH(0) & "</font>"
        End If
        RsH.Close
        Set RsH = Nothing
        Response.Write "  </td>" & vbCrLf
        Response.Write "</tr>" & vbCrLf
        Response.Write "<tr class='tdbg'>" & vbCrLf
        Response.Write "  <td colspan='2'>" & vbCrLf
        Response.Write "     <input name=""chkAll"" type=""checkbox"" id=""chkAll"" onclick=CheckAll(this.form) value=""checkbox"" > &nbsp;全选 &nbsp;&nbsp;"
        Response.Write "  </td>" & vbCrLf
        Response.Write "  <td colspan='7'>" & vbCrLf
        Response.Write "    <input name=""ItemNum"" type=""hidden"" id=""ItemNum"" value=""1"">"
        Response.Write "    <input name=""ListNum"" type=""hidden"" id=""ListNum"" value=""1"">"
        Response.Write "    <input name=""Arr_i"" type=""hidden"" id=""Arr_i"" value=""0"">"
        Response.Write "    <input name=""CollecNewsA"" type=""hidden"" id=""CollecNewsA"" value=""0"">"
        Response.Write "    <input name=""CollecNewsi"" type=""hidden"" id=""CollecNewsi"" value=""0"">"
        Response.Write "    <input name=""ItemSucceedNum"" type=""hidden"" id=""ItemSucceedNum"" value=""0"">"
        Response.Write "    <input name=""ItemSucceedNum2"" type=""hidden"" id=""ItemSucceedNum2"" value=""0"">"
        Response.Write "    <input name=""CollecNewsj"" type=""hidden"" id=""CollecNewsj"" value=""0"">"
        Response.Write "    <input name=""ImagesNumAll"" type=""hidden"" id=""ImagesNumAll"" value=""0"">"
        Response.Write "    <input name=""ItemIDtemp"" type=""hidden"" id=""ItemIDtemp"" value=""0"">"
        Response.Write "    <input name=""Action"" type=""hidden"" id=""Action"" value=""Start"">"
        Response.Write "    <input name=""CollecType"" type=""hidden"" id=""ItemNum"" value=""1"">"
        Response.Write "    <INPUT TYPE='checkbox' NAME='CollecTest' value='yes' zzz='1' onclick=""javascript:document.myform.Content_View.checked=true""> 不录入数据库，只测试采集功能是否正常<br>" & vbCrLf
        Response.Write "    <INPUT TYPE='checkbox' NAME='Content_View' value='yes' zzz='1'> 采集过程中预览文章内容<br>" & vbCrLf
        Response.Write "    <INPUT TYPE='checkbox' NAME='IsTitle' value='yes' zzz='1'> 不采集保存栏目中已有的相同标题文章<br>" & vbCrLf
        Response.Write "    <INPUT TYPE='checkbox' NAME='IsLink' value='yes' zzz='1'> 内部链接采集（此选项只针对链接采集）<br>" & vbCrLf
        Response.Write "  </td>" & vbCrLf
        Response.Write "</tr>" & vbCrLf
        Response.Write "<tr class='tdbg'>" & vbCrLf
        Response.Write "  <td colspan='9' height='32' align='center'>" & vbCrLf
        Response.Write "    <input type=""submit"" value=""快 速 采 集"" name=""submit"" onclick=""javascript:mysub();document.myform.Action.value='Start';document.myform.CollecType.value=1"" >&nbsp;&nbsp;&nbsp;"
        Response.Write "    <input type=""submit"" value=""稳 定 采 集"" name=""submit"" onclick=""javascript:mysub();document.myform.Action.value='Start';document.myform.CollecType.value=0"" >&nbsp;&nbsp;&nbsp;"
        Response.Write "    <input type=""submit"" value=""链 接 采 集"" name=""submit"" onclick=""javascript:if (confirm('链接采集，就是只采集对方网站的链接，不采集正文，这里建议您设置好采集项目的标题和简介，在按扭的上方可设置是内部链接还是外部链接，内部链接就是文章内容只保存对方的URL您可以在模板加内联页加以扩展，外部链接就是列表点击后转向链接，您确定使用链接采集么？')){mysub();document.myform.Action.value='Start';document.myform.CollecType.value=2;}else{return false;};"" >&nbsp;&nbsp;&nbsp;"
        Response.Write "    <input type=""submit"" value=""断 点 续 采 "" name=""submit"""
        '得到断点记录
        Dim rsBreakpoint
        sql = "select top 1 Timing_Breakpoint from PE_config"
        Set rsBreakpoint = Server.CreateObject("adodb.recordset")
        rsBreakpoint.Open sql, Conn, 1, 3
        If rsBreakpoint("Timing_Breakpoint") = "" Then
            Response.Write " disabled"
        End If
        Response.Write "    onclick=""javascript:if (confirm('上次采集因为您停止了采集项目或XMLHTTP组件服务器故障导致中止，现在您是否继续上次的采集项目？')){mysub();document.myform.Action.value='Start';document.myform.CollecType.value=3;}else{return false;};"" >&nbsp;&nbsp;&nbsp;"
        rsBreakpoint.Close
        Set rsBreakpoint = Nothing
        Response.Write "    <input type=""submit"" value=""检测采集项目"" name=""CheckItem"" onclick=""javascript:if (confirm('当你的采集项目比较多，而且长时间未使用采集时，你可能不能确定哪些采集项目还能正常使用，在此情况下你可以使用本功能来检测。此功能非常耗时，请尽量少用。确定要进行检测吗？')){mysub();document.myform.Action.value='CheckItem'}else{return false;};"" ></td>"
        Response.Write "  </td></tr>" & vbCrLf
        Response.Write "</form>" & vbCrLf
        Response.Write "</table>" & vbCrLf

        If totalPut > 0 Then
            Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "个项目记录", True)
        End If
        Response.Write "<br>" & vbCrLf
        Response.Write "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""1"" class=""border"" >" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "   <td colspan='2' height=""20"" align=""center""><font color=#ff6600><strong>声明：因使用本系统提供的采集功能所引起或导致的一切法律或经济责任都由使用者承担，本系统开发商不承担任何责任！</strong></font></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
    End If
    rs.Close
    Set rs = Nothing
    Response.Write " <div id=""esave"" style=""position:absolute; top:50px; left:200px; z-index:1;visibility:hidden""> " & vbCrLf
    Response.Write "    <TABLE WIDTH=400 BORDER=0 CELLSPACING=0 CELLPADDING=0>" & vbCrLf
    Response.Write "      <TR><td width=""20%""></td>" & vbCrLf
    Response.Write "    <TD width=""60%""> " & vbCrLf
    Response.Write "    <TABLE WIDTH=100% height=100 BORDER=0 CELLSPACING=1 CELLPADDING=0>" & vbCrLf
    Response.Write "    <TR> " & vbCrLf
    Response.Write "      <td bgcolor=""#0033FF"" align=center><b><marquee align=""middle"" behavior=""alternate"" scrollamount=""5""><font color=#FFFFFF>正在加载采集项目,请稍候...</font></marquee></b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    Response.Write "    </td><td width='20%'></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    </table>" & vbCrLf
    Response.Write "  </div>" & vbCrLf
    Response.Write " <table WIDTH=400 height=130 BORDER=0 CELLSPACING=0 CELLPADDING=0><tr><td></td></tr></table>" & vbCrLf
    Call CloseConn
End Sub

'=================================================
'过程名：Start
'作  用：保存批量采集文章
'=================================================
Sub Start()
    FoundErr = False                    '是否有错务
    ItemEnd = False                     '是否采集项目完成
    ListEnd = False                     '是否采集列表完成
    ErrMsg = ""                         '错务说明
    TimeNum = 3     '等待时间
    ItemNum = PE_CLng(Trim(Request("ItemNum")))                     'ItemNum    项目数
    ListNum = PE_CLng(Trim(Request("ListNum")))                     'ListNum    列表数
    Arr_i = PE_CLng(Trim(Request("Arr_i")))                         'Arr_i      当前列表的第几文章数
    CollecNewsi = PE_CLng(Trim(Request("CollecNewsi")))             'CollecNewsi    显示采集成功数
    CollecNewsj = PE_CLng(Trim(Request("CollecNewsj")))             'CollecNewsj    显示采集失败数
    ListPaingNext = Replace(Trim(Request("ListPaingNext")),"|","/") 'ListPaingNext  显示采集列表下一页
    ItemIDStr = Replace(ReplaceBadChar(Trim(Request("ItemID"))), " ", "") 'ItemIDStr 项目数组
    CollecNewsA = PE_CLng(Trim(Request("CollecNewsA")))             'CollecNewsA 采集文章数
    ItemIDtemp = PE_CLng(Trim(Request("ItemIDtemp")))               'ItemIDtemp  项目是否首次加载
    rnd_temp = CStr(Trim(Request("rnd_temp")))                      'rnd_temp    用于随机数不同的缓存
    ArticleList = Replace(CStr(Trim(Request("ArticleList"))),"|","/")                      'ArticleList 用于缓存不同的列表
    ItemSucceedNum = PE_CLng(Trim(Request("ItemSucceedNum")))       'ItemSucceedNum 项目采集成功数为记录不同项目采集成功数用于多项目指定的采集数量
    ItemSucceedNum2 = PE_CLng(Trim(Request("ItemSucceedNum2")))     'ItemSucceedNum2 成功采集项目数
    ImagesNumAll = PE_CLng(Trim(Request("ImagesNumAll")))           'ImagesNumAll    项目总数
    CollecType = PE_CLng(Trim(Request("CollecType")))               'CollecType    采集模式 0 稳定 1 快速 2 链接 3 断点续采
    CollectionCreateHTML = Trim(Request("CollectionCreateHTML"))    'CollectionCreateHTML    生成html数组
    TimingCreate = Trim(Request("TimingCreate"))                    'TimingCreate  定时生成html
    Timing_AreaCollection = Trim(Request("Timing_AreaCollection"))  'Timing_AreaCollection  定时区域采集

    If CollecType = 3 Then
        '断点续采
        sql = "select top 1 Timing_Breakpoint from PE_config"
        Set rs = Server.CreateObject("adodb.recordset")
        rs.Open sql, Conn, 1, 3
        Response.Write rs("Timing_Breakpoint")
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If Trim(Request("CollecTest")) = "yes" Then
        CollecTest = True
    Else
        CollecTest = False
    End If
    If Trim(Request("Content_view")) = "yes" Then
        Content_view = True
    Else
        Content_view = False
    End If

    If Trim(Request("IsTitle")) = "yes" Then
        IsTitle = True
    Else
        IsTitle = False
    End If

    If Trim(Request("IsLink")) = "yes" Then
        IsLink = True
    Else
        IsLink = False
    End If
	If IsValidID(ItemIDStr) = False Then
		ItemIDStr = ""
	End If

    If ItemIDStr = "" Then
        FoundErr = True
        ErrMsg = "<li>参数错误,请选择项目！</li>"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    ElseIf ItemIDStr = "0" Then '为定时管理跳转
        Call Refresh("Admin_Timing.asp?Action=DoTiming&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,0)	
        'Response.Write " <meta http-equiv=""refresh"" content=0;url=""Admin_Timing.asp?Action=DoTiming&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
        Exit Sub
    End If
     
    '是否全部采集完成
    'ItemNum 当前项目数 ItemIDStr 项目数组 分割数组 ItemIDArray 得到项目数 ItemIDArray 得到每一个项目数
    ItemIDArray = Split(ItemIDStr, ",")
    If (ItemNum - 1) > UBound(ItemIDArray) Then
        ItemEnd = True
        ErrMsg = "<br>全部项目采集任务完成！"
        ErrMsg = ErrMsg & "<li>成功采集： <font color=red>" & CollecNewsi & "</font>  篇,失败：<font color=blue> " & CollecNewsj & "</font>  篇,图片：<font color=green>" & ImagesNumAll & "</font> 个。</li>"

        '清空断点记录
        sql = "select Timing_Breakpoint from PE_config"
        Set rs = Server.CreateObject("adodb.recordset")
        rs.Open sql, Conn, 1, 3
        rs("Timing_Breakpoint") = ""
        rs.Update
        rs.Close
        Set rs = Nothing
        
        '清除缓存
        Call PE_Cache.DelAllCache
        Call WriteSuccessMsg2(ErrMsg)
        Exit Sub
    End If
    
    '加载初始项目入缓存
    If ItemIDtemp = 0 Then
        Call SetCache
        ItemIDtemp = 1
    Else
        If PE_Cache.CacheIsEmpty("Collection" & rnd_temp) Then
            Call SetCache
            ArticleList = ""
        End If
    End If
    
    '加载缓存
    Arr_Item = PE_Cache.GetValue("Collection" & rnd_temp)
    Arr_Filters = PE_Cache.GetValue("Arr_Filters" & rnd_temp)
    Arr_Histrolys = PE_Cache.GetValue("Arr_Histrolys" & rnd_temp)
    Call loadItem
    
    If CollectionNum <> "" Then   '是否到了指定的成功采集数
        If CollectionType = 0 Then  '是否到了指定的成功采集数
            If ItemSucceedNum = PE_CLng(CollectionNum) Then
                ErrMsg = "<li>已经成功采集了" & ItemName & "项目<font color=red>" & CollectionNum & "</font>篇指定采集数。</li>"
                ErrMsg = ErrMsg & "<br><font color=red>" & ItemName & "</font> 项目采集任务完成！</li>"
                Call WriteSuccessMsg2(ErrMsg)
                Exit Sub
            End If
        End If
        If CollectionType = 1 Then  '是否到了每页的要的采集数
            If ListNum > PE_CLng(CollectionNum) Then
                ErrMsg = "<li>已经成功采集了" & ItemName & "项目<font color=red>" & CollectionNum & "</font>篇指定列数。</li>"
                ErrMsg = ErrMsg & "<br><font color=red>" & ItemName & "</font> 项目采集任务完成！</li>"
                Call WriteSuccessMsg2(ErrMsg)
                Exit Sub
            End If
        End If
    End If
    
    '更新项目记录时间
    If ListNum = 1 And CollecTest = False Then
        sql = "select top 1 * from PE_Item where ItemID=" & ItemID
        Set rs = Server.CreateObject("adodb.recordset")
        rs.Open sql, Conn, 1, 3
        rs("NewsCollecDate") = Now()
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
    
    If LoginType = 1 And ListNum = 1 Then '采集登录
        '登录网站
        LoginData = UrlEncoding(LoginUser & "&" & LoginPass)
        LoginResult = PostHttpPage(LoginUrl, LoginPostUrl, LoginData, PE_CLng(WebUrl))
        If InStr(LoginResult, LoginFalse) > 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>在登录网站时发生错误,请确保登录信息的正确性！</li>"
            ItemNum = ItemNum + 1
            ListNum = 1
            ArticleList = ""
            Call Refresh("Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&CollecNewsj=" & CollecNewsj & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,TimeNum)			
            'Response.Write "   <meta http-equiv=""refresh"" content=" & TimeNum & ";url=""Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&CollecNewsj=" & CollecNewsj & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"

            Call WriteErrMsg(ErrMsg, ComeUrl)
            Exit Sub
        End If
    End If
    
    If FoundErr <> True And ItemEnd <> True Then
        '继续采集列表
        If ArticleList = "" Then  '加载列表处理
            '判断列表类型
            '不作分页设置
            If ListPaingType = 0 Then
                If ListNum = 1 Then
                    '列表链接=列表索引页面
                    ListUrl = ListStr
                Else
                    ListEnd = True
                End If
            '设置标签
            ElseIf ListPaingType = 1 Then
                '判断列表为1时加载链接地址
                If ListNum = 1 Then
                    ListUrl = ListStr
                Else
                    If ListPaingNext = "" Or ListPaingNext = "$False$" Then
                        ListEnd = True
                    Else
                        If InStr(ListPaingNext, "{$ID}") > 0 Then
                            ListPaingNext = Replace(ListPaingNext, "{$ID}", "&")
                        End If
                        ListUrl = ListPaingNext
                    End If
                End If
            '批量生成
            ElseIf ListPaingType = 2 Then
                If ListPaingID1 > ListPaingID2 Then
                    If (ListPaingID1 - ListNum + 1) < ListPaingID2 Or (ListPaingID1 - ListNum + 1) < 0 Then
                        ListEnd = True
                    Else
                        ListUrl = Replace(ListPaingStr2, "{$ID}", CStr(ListPaingID1 - ListNum + 1))
                    End If
                Else
                    If (ListPaingID1 + ListNum - 1) > ListPaingID2 Then
                        ListEnd = True
                    Else
                        ListUrl = Replace(ListPaingStr2, "{$ID}", CStr(ListPaingID1 + ListNum - 1))
                    End If
                End If
            '手动添加
            ElseIf ListPaingType = 3 Then
                ListArray = Split(ListPaingStr3, vbCrLf)
                If (ListNum - 1) > UBound(ListArray) Then
                    ListEnd = True
                Else
                    ListUrl = ListArray(ListNum - 1)
                End If
            End If
            If ListEnd <> True Then
                If CheckUrl(ListStr) = False Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>列表网址不对！</li>"
                End If
                ArticleList = ListUrl
                If InStr(ListUrl, "{$ID}") > 0 Then
                    ListUrl = Replace(ListUrl, "{$ID}", "&")
                End If
                If FoundErr <> True Then
                    ListCode = GetHttpPage(ListUrl, PE_CLng(WebUrl)) '获取网页源代码 ListCode
                    '类型为设置标签时
                    If ListPaingType = 1 Then
                        ListPaingNext = GetPaing(ListCode, LPsString, LPoString, False, False)
                        If ListPaingNext <> "$False$" Then
                            If ListPaingStr1 <> "" Then
                                ListPaingNext = Replace(ListPaingStr1, "{$ID}", ListPaingNext)
                            Else
                                ListPaingNext = DefiniteUrl(ListPaingNext, ListUrl)
                            End If
                            If InStr(ListPaingNext, "&") > 0 Then
                                ListPaingNext = Replace(ListPaingNext, "&", "{$ID}")
                            End If
                        End If
                    Else
                        ListPaingNext = "$False$"
                    End If
                    
                    If ListCode = "$False$" Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>在获取：" & ListUrl & "网页源码时发生错误！</li>"
                    Else
                        ListCode = GetBody(ListCode, LsString, LoString, False, False) '截取列表字符串
                        If ListCode = "$False$" Or ListCode = "" Then
                            FoundErr = True
                            ErrMsg = ErrMsg & "<li>在截取：" & ListUrl & "列表时发生错误！</li>"
                        End If
                    End If
                End If
                If FoundErr <> True Then
                    NewsArrayCode = GetArray(ListCode, HsString, HoString, False, False) 'NewsArrayCode=在列表中提取链接地址
                    If ThumbnailType = 1 Then
                        ThumbnailArrayCode = GetArray(ListCode, ThsString, ThoString, False, False) '缩略图地址
                    End If
                End If

                If NewsArrayCode = "$False$" Or FoundErr = True Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>在分析：" & ListUrl & "新闻列表时发生错误！</li>"
                    ItemNum = ItemNum + 1
                    ListNum = 1
                    ArticleList = ""
                    '生成Html
                    Call GetArrOfCreateHTML
                    Call Refresh("Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&CollecNewsj=" & CollecNewsj & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|")  & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,TimeNum)	
                    'Response.Write "   <meta http-equiv=""refresh"" content=" & TimeNum & ";url=""Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&CollecNewsj=" & CollecNewsj & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
                    Call WriteErrMsg(ErrMsg, ComeUrl)

                    Exit Sub
                Else
                    '分割链接文章地址
                    NewsArray = Split(NewsArrayCode, "$Array$")
                    For Arr_j = 0 To UBound(NewsArray)
                        '当链接地址要从新定位时？
                        If HttpUrlType = 1 Then
                            NewsArray(Arr_j) = Trim(Replace(HttpUrlStr, "{$ID}", NewsArray(Arr_j)))
                        Else
                            '过滤空格并将相对地址转换为绝对地址
                            NewsArray(Arr_j) = Trim(DefiniteUrl(NewsArray(Arr_j), ListUrl))
                        End If
                    Next
                    If PE_CLng(CollecOrder) = 1 Then '如果是倒序采集
                        '颠倒当前数组的顺序
                        For Arr_j = 0 To Fix(UBound(NewsArray) / 2)
                            OrderTemp = NewsArray(Arr_j)
                            NewsArray(Arr_j) = NewsArray(UBound(NewsArray) - Arr_j)
                            NewsArray(UBound(NewsArray) - Arr_j) = OrderTemp
                        Next
                    End If
                    '列表缩略图地址
                    If ThumbnailType = 1 Then
                        '分割链接文章地址
                        ThumbnailArray = Split(ThumbnailArrayCode, "$Array$")
                        For Arr_j = 0 To UBound(ThumbnailArray)
                            '过滤空格并将相对地址转换为绝对地址
                            ThumbnailArray(Arr_j) = Trim(DefiniteUrl(ThumbnailArray(Arr_j), ListUrl))
                        Next
                        If PE_CLng(CollecOrder) = 1 Then '如果是倒序采集
                            '颠倒当前数组的顺序
                            For Arr_j = 0 To Fix(UBound(ThumbnailArray) / 2)
                                OrderTemp = ThumbnailArray(Arr_j)
                                ThumbnailArray(Arr_j) = ThumbnailArray(UBound(ThumbnailArray) - Arr_j)
                                ThumbnailArray(UBound(ThumbnailArray) - Arr_j) = OrderTemp
                            Next
                        End If
                        PE_Cache.SetValue "ThumbnailList" & rnd_temp, ThumbnailArray '加载缓存
                    End If
                    PE_Cache.SetValue "ArticleList" & rnd_temp, NewsArray '加载缓存

                    '更新断点记录
                    sql = "select Timing_Breakpoint from PE_config"
                    Set rs = Server.CreateObject("adodb.recordset")
                    rs.Open sql, Conn, 1, 3
                    rs("Timing_Breakpoint") = " <meta http-equiv=""refresh"" content=0;url=""Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
                    rs.Update
                    rs.Close
                    Set rs = Nothing
                    Call Refresh("Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,0)		
                    'Response.Write " <meta http-equiv=""refresh"" content=0;url=""Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
                    Exit Sub
                End If
            Else
                '生成Html
                Call GetArrOfCreateHTML

                ErrMsg = ErrMsg & "<br><font color=red>" & ItemName & "</font> 项目采集任务完成！</li>"
                Call WriteSuccessMsg2(ErrMsg)
                Exit Sub
            End If
        Else
            NewsArray = PE_Cache.GetValue("ArticleList" & rnd_temp)
            If ThumbnailType = 1 Then
                ThumbnailArray = PE_Cache.GetValue("ThumbnailList" & rnd_temp)
            End If
        End If
        
        '加载导航信息
        Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"" >" & vbCrLf
        Response.Write "    <tr> " & vbCrLf
        Response.Write "      <td height=""22"" colspan=""2"" class=""tdbg"" align=""left"">&nbsp;&nbsp;采集需要一定的时间,请耐心等待,如果网站出现暂时无法访问的情况这是正常的,采集过程正常结束后即可恢复。" & vbCrLf
        Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button""  name=""Stop""  value=""停止采集""  onCLICK=""location.href='Admin_Collection.asp?Action=StopCollection&rnd_temp=" & rnd_temp & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&CollectionCreateHTML=" & CollectionCreateHTML & "&ChannelID=" & ChannelID & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&CreateImmediate=" & CreateImmediate & "&UseCreateHTML=" & UseCreateHTML & "&TimingCreate=" & TimingCreate & "'"">" & vbCrLf
        Response.Write "      </td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        '加载显示采集全局信息
        Response.Write "    <tr>" & vbCrLf
        Response.Write "      <td height=""22"" colspan=""2"" class=""tdbg"" align=""left"">&nbsp;&nbsp;本次运行：" & UBound(ItemIDArray) + 1 & " 个项目,正在采集第 <font color=red>" & ItemNum & "</font> 个项目  <font color=red>" & ItemName & "</font>  的第   <font color=red>" & ListNum & "</font> 页列表,该列表待采集新闻  <font color=red>" & UBound(NewsArray) + 1 & "</font> 条,中的第 <font color=red>" & Arr_i + 1 & "</font> 条。" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr>"
        Response.Write "      <td height=""22"" colspan=""2"" class=""tdbg"" align=""left"">&nbsp;&nbsp;采集统计：成功采集--" & CollecNewsi & "  条新闻,失败--" & CollecNewsj & "  条,图片--" & ImagesNumAll & "</td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
        Response.Write "<br>"
        StartTime = Timer()

        If CollecType = 0 Then
            '执行核心采集过程
            Call StartCollection
            '………………………………………………
            '采集数处理 当前列表采集完成 or 指定数完成
            If CollectionType = 0 And CollectionNum <> "" Then
                If CLng(ItemSucceedNum2) >= CLng(CollectionNum) Then
                    '生成Html
                    Call GetArrOfCreateHTML
                    ListNum = ListNum + 1
                    ArticleList = ""
                    ItemSucceedNum2 = 0 '统计数采集项目都清0为下一个采集项目准备
                Else
                    Arr_i = Arr_i + 1 '移动到下一采集文章
                End If
            ElseIf CollectionType = 1 And CollectionNum <> "" Then
                If ListNum > PE_CLng(CollectionNum) Then
                    ArticleList = ""      '采集列表完成
                    '生成Html
                    Call GetArrOfCreateHTML
                    ItemSucceedNum2 = 0   '统计数采集项目都清0为下一个采集项目准备
                Else
                    Arr_i = Arr_i + 1 '移动到下一采集文章
                End If
            Else
                Arr_i = Arr_i + 1 '移动到下一采集文章
            End If
            If Arr_i > UBound(NewsArray) Then
                Arr_i = 0
                ListNum = ListNum + 1
                ArticleList = ""      '采集列表完成
            End If
        Else
            For Arr_i = 0 To UBound(NewsArray)
                FoundErr = False
                Call StartCollection  '执行核心采集过程

                '采集数处理 当前列表采集完成 or 指定数完成
                If CollectionType = 0 And CollectionNum <> "" Then
                    If CLng(ItemSucceedNum2) >= CLng(CollectionNum) Then
                        ListNum = ListNum + 1
                        ArticleList = ""
                        '生成Html
                        Call GetArrOfCreateHTML
                        ItemSucceedNum2 = 0 '统计数采集项目都清0为下一个采集项目准备
                        Exit For
                    End If
                ElseIf PE_CLng(CollectionType) = 1 And CollectionNum <> "" Then
                    If ListNum = PE_CLng(CollectionNum) And Arr_i >= UBound(NewsArray) Then
                        ListNum = ListNum + 1
                        ArticleList = ""      '采集列表完成
                        '生成Html
                        Call GetArrOfCreateHTML
                        ItemSucceedNum2 = 0   '统计数采集项目都清0为下一个采集项目准备
                        Exit For
                    End If
                End If
                If Arr_i >= UBound(NewsArray) Then
                    Arr_i = 0
                    ListNum = ListNum + 1
                    ArticleList = ""      '采集列表完成
                    Exit For
                End If
            Next
        End If
    End If

    Response.Write "<br>"
    Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"" >" & vbCrLf
    Response.Write "  <tr>"
    Response.Write "   <td height=""22"" align=""left"" class=""tdbg"">"
    Response.Write "&nbsp;&nbsp;数据整理中," & TimeNum & " 秒后继续......" & TimeNum & "秒后如果还没反应请点击 <a href='Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & "'><font color=red>这里</font></a> 继续<br>"
    Call Refresh("Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,TimeNum)
    'Response.Write "   <meta http-equiv=""refresh"" content=" & TimeNum & ";url=""Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
    Response.Write "&nbsp;&nbsp;执行时间：" & CStr(FormatNumber((Timer() - StartTime) * 1000, 2)) & " 毫秒"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr> " & vbCrLf
    Response.Write "    <td height=""22""  class=""tdbg"" align=""left"">&nbsp;&nbsp;采集需要一定的时间,请耐心等待,如果网站出现暂时无法访问的情况这是正常的,采集过程正常结束后即可恢复。" & vbCrLf
    Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button""  name=""Stop""  value=""停止采集""  onCLICK=""location.href='Admin_Collection.asp?Action=StopCollection&rnd_temp=" & rnd_temp & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&CollectionCreateHTML=" & CollectionCreateHTML & "&ChannelID=" & ChannelID & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&CreateImmediate=" & CreateImmediate & "&UseCreateHTML=" & UseCreateHTML & "&TimingCreate=" & TimingCreate & "'"">" & vbCrLf
    Response.Write "    </td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End Sub
'==================================================
'过程名：StartCollection
'作  用：开始采集
'参  数：无
'==================================================
Sub StartCollection()

    '………………………………………………
    '内容页变量初始化
    CollecNewsA = CollecNewsA + 1 '已经采集数（包含成功和失败）
    DefaultPicUrl = ""   '要采集的绝对路径
    ImagesNum = 0        '本次采集采集到的图片数量
    NewsCode = ""        '获得内容也的源代码
    Title = ""           '标题
    Content = ""         '正文
    Author = ""          '作者
    CopyFrom = ""        '来源
    Key = ""             '关键字
    His_Repeat = False   '是否采集过
    NewsUrl = Trim(NewsArray(Arr_i)) '要采集的正文链接页
    If ThumbnailType = 1 Then
        ThumbnailUrl = Trim(ThumbnailArray(Arr_i))
    End If

    PaingNum = 1               '正文中有多少分页
    UploadFiles = ""           '上传的图片地址
    ErrMsg = ""

    '………………………………………………
    '检测客户连接是否仍然有效
    If Response.IsClientConnected Then
        Response.Flush '强迫输出Html 到浏览器
    Else
        Exit Sub
    End If

    If CollecTest = False Then
        His_Repeat = CheckRepeat(NewsUrl)
    Else
        His_Repeat = False
    End If

    If His_Repeat = True Then
        FoundErr = True
    End If

    '标题 正文 获取过滤
    If FoundErr <> True Then
        'NewsCode 获取内容页Html
        NewsCode = GetHttpPage(NewsUrl, PE_CLng(WebUrl))

        If NewsCode = "$False$" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>在获取：" & NewsUrl & "新闻源码时发生错误！</li>"
        End If
    End If
    If FoundErr <> True Then
        Title = FpHtmlEnCode(Trim(GetBody(NewsCode, TsString, ToString, False, False))) '获得标题代码
        If Title = "$False$" Or Title = "" Or Len(Title) > 200 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>在采集：" & NewsUrl & "新闻标题时发生错误</li>"
        End If
        If CollecTest = False And IsTitle = True And FoundErr <> True Then
            If PE_CLng(Conn.Execute("Select count(*) From PE_Article Where Title='" & Title & "' And ClassID =" & ClassID)(0)) > 0 Then
                FoundErr = True
            End If
        End If
    End If
    If FoundErr <> True Then
        If CollecType <> 2 Then
            Content = Trim(GetBody(NewsCode, CsString, CoString, False, False)) '获得正文代码
        End If
        If Content = "$False$" Or Content = "" And CollecType <> 2 Then '如果标题和正文产生错误
            FoundErr = True
            ErrMsg = ErrMsg & "<li>在采集：" & NewsUrl & "新闻正文时发生错误</li>"

            If CollecTest = False Then '不为测试时
                '写入历史记录
                sql = "INSERT INTO PE_HistrolyNews(ItemID,ChannelID,ClassID,NewsCollecDate,NewsUrl,Result) VALUES ('" & ItemID & "','" & ChannelID & "','" & ClassID & "','" & Now() & "','" & NewsUrl & "'," & PE_False & ")"
                Conn.Execute (sql)
            End If

        Else
            If CollecType <> 2 Then
                If NewsPaingType = 1 Then '新闻分页 正文分页为 设置标签时
                    NewsPaingNext = GetPaing(NewsCode, NPsString, NPoString, False, False) '获取分页地址
                    
                    '影响了部分内容页分页暂时中止
                    'If Left(NewsPaingNext,1) = "/" then
                        ConversionTrails = NewsUrl
                    'Else
                    '    ConversionTrails = ListStr
                    'End If
                    NewsPaingNext = DefiniteUrl(NewsPaingNext, ConversionTrails) '将相对路径转绝对路径

                    Dim NewsPaingNextTemp

                    Do While NewsPaingNext <> "$False$" And NewsPaingNext <> ""

                        If NewsPaingNextTemp <> "" Then
                            If FoundInArr(NewsPaingNextTemp, NewsPaingNext, "$$$") = True Then
                                Exit Do
                            Else
                                NewsPaingNextTemp = NewsPaingNextTemp & "$$$" & NewsPaingNext
                            End If

                        Else
                            NewsPaingNextTemp = NewsPaingNext
                        End If
                        
                        If CheckUrl(NewsPaingNext) = False Then
                            Response.Write "<font color=red>内容页分页代码不正确,不是有效的网页链接代码。</a>"
                            Exit Do
                        End If

                        NewsPaingNextCode = GetHttpPage(NewsPaingNext, PE_CLng(WebUrl)) '获得分页html代码

                        If NewsPaingNextCode = "$False$" Or NewsPaingNextCode = "" Then Exit Do
                        ContentTemp = GetBody(NewsPaingNextCode, CsString, CoString, False, False) '截取正文代码

                        If ContentTemp = "$False$" Or ContentTemp = "" Then
                            Exit Do
                        Else
                            PaingNum = PaingNum + 1

                            If PaginationType = 2 Then '加一行段落链接上一正文
                                Content = Content & "<p> </p>[NextPage]<p> </p>" & ContentTemp
                            Else
                                Content = Content & "<p> </p>" & ContentTemp
                            End If
                            '得到下一分页链接代码
                            NewsPaingNext = GetPaing(NewsPaingNextCode, NPsString, NPoString, False, False) '获取分页地址
                            ''影响了部分内容页分页暂时中止
                            'If Left(NewsPaingNext,1) = "/" then
                                ConversionTrails = NewsUrl
                            'Else
                            '    ConversionTrails = ListStr
                            'End If

                            NewsPaingNext = DefiniteUrl(NewsPaingNext, ConversionTrails) '将相对路径转绝对路径
                        End If

                    Loop

                ElseIf NewsPaingType = 2 Then
                    PageListCode = GetBody(NewsCode, PsString, PoString, False, False) '获取列表页

                    If PageListCode <> "$False$" Then
                        PageArrayCode = GetArray(PageListCode, PhsString, PhoString, False, False) '获取链接地址

                        If PageArrayCode <> "$False$" Then
                            If InStr(PageArrayCode, "$Array$") > 0 Then
                                '去掉地址开始
                                Dim tempk, TempPaingNext
                                PageArray = Split(PageArrayCode, "$Array$") '分割得到地址
                                TempPaingNext = ""
                                For tempk = 0 To UBound(PageArray)
                                    If InStr(LCase(TempPaingNext), LCase(PageArray(tempk))) < 1 Then
                                        TempPaingNext = TempPaingNext & "$Array$" & PageArray(tempk)
                                    End If
                                Next
                                TempPaingNext = Right(TempPaingNext, Len(TempPaingNext) - 7)
                                PageArray = Split(TempPaingNext, "$Array$")
                                '去掉地址结束

                                For i = 0 To UBound(PageArray)
                                    NewsPaingNextCode = GetHttpPage(DefiniteUrl(PageArray(i), NewsUrl), PE_CLng(WebUrl)) '获得分页html代码

                                    If NewsPaingNextCode <> "$False$" Or NewsPaingNextCode <> "" Then
                                        ContentTemp = GetBody(NewsPaingNextCode, CsString, CoString, False, False) '截取正文代码

                                        If ContentTemp <> "$False$" Or ContentTemp <> "" Then
                                            PaingNum = PaingNum + 1

                                            If PaginationType = 2 Then '加一行段落链接上一正文
                                                Content = Content & "<p> </p>[NextPage]<p> </p>" & ContentTemp
                                            Else
                                                Content = Content & "<p> </p>" & ContentTemp
                                            End If
                                        End If
                                    End If

                                Next

                            Else
                                NewsPaingNextCode = GetHttpPage(DefiniteUrl(PageArrayCode, NewsUrl), PE_CLng(WebUrl)) '获得分页html代码

                                If NewsPaingNextCode <> "$False$" Or NewsPaingNextCode <> "" Then
                                    ContentTemp = GetBody(NewsPaingNextCode, CsString, CoString, False, False) '截取正文代码

                                    If ContentTemp <> "$False$" Or ContentTemp <> "" Then
                                        PaingNum = PaingNum + 1

                                        If PaginationType = 2 Then '加一行段落链接上一正文
                                            Content = Content & "<p> </p>[NextPage]<p> </p>" & ContentTemp
                                        Else
                                            Content = Content & "<p> </p>" & ContentTemp
                                        End If
                                    End If
                                End If
                            End If

                        Else
                            Response.Write "<li>在获取分页链接列表时出错。</li>"
                        End If

                    Else
                        Response.Write "<li>在截取分页代码发生错误。</li>"
                    End If
                End If
            End If
            Call Filters ' 标题过滤 正文过滤 广告
        End If
    End If

    If FoundErr <> True Then

        '………………………………………………
        '时间
        If UpDateType = 0 Or UpDateType = "" Then
            UpdateTime = Now()
        ElseIf UpDateType = 1 Then

            If DateType = 0 Then
                UpdateTime = Now()
            Else
                UpdateTime = GetBody(NewsCode, DsString, DoString, False, False)
                UpdateTime = FpHtmlEnCode(UpdateTime)
                UpdateTime = PE_CDate(Trim(Replace(UpdateTime, "&nbsp;", " ")))
            End If

        ElseIf UpDateType = 2 Then
        Else
            UpdateTime = Now()
        End If

        '………………………………………………
        '作者获取过滤
        If AuthorType = 1 Then
            Author = GetBody(NewsCode, AsString, AoString, False, False) '获得当前作者字符
        ElseIf AuthorType = 2 Then '指定作者
            Author = AuthorStr
        Else '为0时
            Author = "佚名"
        End If

        '作者过滤
        Author = FpHtmlEnCode(Author)

        If Author = "" Or Author = "$False$" Then
            Author = "佚名"
        Else

            '只左边30个字符（没有人会叫很长的名字）
            If Len(Author) > 30 Then
                Author = Left(Author, 30)
            End If
        End If

        '………………………………………………
        '来源获取过滤
        If CopyFromType = 1 Then
            CopyFrom = GetBody(NewsCode, FsString, FoString, False, False)
        ElseIf CopyFromType = 2 Then
            CopyFrom = CopyFromStr
        Else
            CopyFrom = "不详"
        End If

        CopyFrom = FpHtmlEnCode(CopyFrom)

        If CopyFrom = "" Or CopyFrom = "$False$" Then
            CopyFrom = "不详"
        Else

            If Len(CopyFrom) > 30 Then
                CopyFrom = Left(CopyFrom, 30)
            End If
        End If

        '………………………………………………
        '关键字获取过滤
        If KeyType = 0 Then
            Key = Title
            Key = CreateKeyWord(Key, KeyScatterNum)
        ElseIf KeyType = 1 Then
            Key = GetBody(NewsCode, KsString, KoString, False, False)
            Key = FpHtmlEnCode(Key)

            Key = Replace(Key, ",", "|")
            Key = Replace(Key, "&nbsp;", "|")
            Key = Replace(Key, " ", "|")

            Dim arrKey, KeyString, j
            arrKey = Split(Key, "|")
            For j = 0 To UBound(arrKey)
                If arrKey(j) <> "" Then
                    If KeyString = "" Then
                        KeyString = arrKey(j)
                    Else
                        KeyString = KeyString & "|" & arrKey(j)
                    End If
                End If
            Next
            Key = KeyString
            'Key = CreateKeyWord(Key, KeyScatterNum)
            If Len(Key) > 253 Then
                Key = "|" & Left(Key, 253) & "|"
            Else
                Key = "|" & Key & "|"
            End If
        ElseIf KeyType = 2 Then
            Key = KeyStr
            Key = FpHtmlEnCode(Key)

            If Len(Key) > 253 Then
                Key = "|" & Left(Key, 253) & "|"
            Else
                Key = "|" & Key & "|"
            End If
        End If

        If Key = "" Or Key = "$False$" Then
            Key = "|" & Title & "|"
        End If

        '过滤非法字符
        Key = ReplaceBadChar(Key)
        
        '保存采集文件绝对路径地址
        If SaveFlashUrlToFile = True Then
            Content = CollectionFilePath(Content, NewsUrl)
        End If

        '保存远程图片
        If CollecTest = False And SaveFiles = True Then
            Content = ReplaceSaveRemoteFile(Content, FilesOverStr, True, FilesPath, NewsUrl)
        Else
            Content = ReplaceSaveRemoteFile(Content, FilesOverStr, False, FilesPath, NewsUrl)
        End If

        
        FilterProperty = Script_Iframe & "|" & Script_Object & "|" & Script_Script & "|" & Script_Class & "|" & Script_Div & "|" & Script_Table & "|" & Script_Tr & "|" & Script_Td & "|" & Script_Span & "|" & Script_Img & "|" & Script_Font & "|" & Script_A & "|" & Script_Html
        Content = FilterScript(Content, FilterProperty) '脚本过滤

        If IntroType = 0 Then
        ElseIf IntroType = 1 Then
            Intro = GetBody(NewsCode, IsString, IoString, False, False)
            Intro = Trim(nohtml(Intro))
        ElseIf IntroType = 2 Then
            Intro = nohtml(IntroStr)
        ElseIf IntroType = 3 Then
            Intro = Left(Replace(Replace(Replace(Replace(Trim(nohtml(Content)), vbCrLf, ""), " ", ""), "&nbsp;", ""), "　", ""), IntroNum)
        End If

        If Intro = "$False$" Then
            Intro = ""
        End If
    End If

    If FoundErr <> True Then
        '………………………………………………
        ' 保存文章显示
        If CollecTest = False Then
            Call SaveArticle
            ' 保存历史记录
            sql = "INSERT INTO PE_HistrolyNews(ItemID,ChannelID,ClassID,ArticleID,Title,NewsCollecDate,NewsUrl,Result) VALUES ('" & ItemID & "','" & ChannelID & "','" & ClassID & "','" & ArticleID & " ','" & Title & "','" & Now() & "','" & NewsUrl & "'," & PE_True & ")"
            Conn.Execute (sql)
            CollecNewsi = CollecNewsi + 1           '成功采集的新闻数量+1
            ItemSucceedNum = ItemSucceedNum + 1     '完成采集项目+1
            ItemSucceedNum2 = ItemSucceedNum2 + 1   '成功采集项目数量+1
        End If

        ErrMsg = ErrMsg & "新闻标题："

        If CollecTest = False Then
            ErrMsg = ErrMsg & "<a href='Admin_Article.asp?ChannelID=" & ChannelID & "&Action=Show&ArticleID=" & ArticleID & "' target=_blank><font color=red>" & Title & "</font></a>"
        Else
            ErrMsg = ErrMsg & "<font color=red>" & Title & "</font>"
        End If

        ErrMsg = ErrMsg & "<br>"
        ErrMsg = ErrMsg & "新闻作者：" & Author & "<br>"
        ErrMsg = ErrMsg & "新闻来源：" & CopyFrom & "<br>"
        ErrMsg = ErrMsg & "关 键 字："
        If Len(Key) > 4 Then
            ErrMsg = ErrMsg & Mid(Key, 2, Len(Key) - 2) & "<br>"
        Else
             ErrMsg = ErrMsg & Key & "<br>"
        End If
        ErrMsg = ErrMsg & "采集页面：<a href=" & NewsUrl & " target=_blank>" & NewsUrl & "</a><br>"
        ErrMsg = ErrMsg & "其它信息：分页--" & PaingNum & " 页,图片--" & ImagesNum & " 张<br>"

        If Content_view = True And CollecType <> 2 Then
            ErrMsg = ErrMsg & "正文预览："
            ErrMsg = ErrMsg & Left(Content, 250) & "......"
        End If

    Else
        CollecNewsj = CollecNewsj + 1 '失败采集的新闻数量+1 添加历史记录 His_Repeat是否添加过了历史记录

        If His_Repeat = True Or IsTitle = True Then
            If CollecType = 0 Then
                TimeNum = 1           '采集的时间间隔
            Else
                TimeNum = 0
            End If

            ErrMsg = ErrMsg & "<li>目标新闻：<font color=red>"

            If His_Title = "" Then
                ErrMsg = ErrMsg & NewsUrl '如果为空显示当前链接
            Else
                ErrMsg = ErrMsg & His_Title
            End If

            ErrMsg = ErrMsg & "</font></a>  已存在,不予采集。"
            ErrMsg = ErrMsg & "<li>采集时间：" & His_NewsCollecDate & "</li>"
            ErrMsg = ErrMsg & "<li>新闻来源：<a href='" & NewsUrl & "' target=_blank>" & NewsUrl & "</a>"
            ErrMsg = ErrMsg & "<li>提示信息：如想再次采集,请先将该新闻的历史记录<font color=red>删除</font></li>"

            If His_Result = True Then
                ErrMsg = ErrMsg & "<li>以及主数据库中的新闻删除</li>"
            End If
        End If

        If CollecTest = False And His_Repeat = False And IsTitle = False Then
            sql = "INSERT INTO PE_HistrolyNews(ItemID,ChannelID,ClassID,NewsUrl,Result) VALUES ('" & ItemID & "','" & ChannelID & "','" & ClassID & "','" & NewsUrl & "'," & PE_False & ")"
            Conn.Execute (sql)
        End If
    End If

    Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"" class='border'>"
    Response.Write "   <tr>"
    Response.Write "      <td height=""22"" colspan=""2"" align=""left"" class=""title"">No:<font color=red>" & CollecNewsA & "</font></td>"
    Response.Write "   </tr>"
    Response.Write "   <tr>"
    Response.Write "      <td colspan=""2"" align=""left"" class=""tdbg"">" & ErrMsg & "</td>"
    Response.Write "   </tr>"
    Response.Write "  <tr>"
    Response.Write "</table>"
    Response.Write "<br>"
End Sub
'=================================================
'过程名：CheckItem
'作  用：批量检测项目是否可采集
'=================================================
Sub CheckItem()
        
    Dim ItemID, ItemName, WebUrl, ListStr, LsString, LoString
    Dim ListPaingType, LPsString, LPoString, ListPaingStr1, ListPaingStr2
    Dim ListPaingID1, ListPaingID2, ListPaingStr3
    Dim HsString, HoString, HttpUrlType, HttpUrlStr
    Dim TsString, ToString, CsString, CoString
    Dim rsItem, sql, FoundErr

    sql = "Select * from PE_Item Where Flag=" & PE_True

    Set rsItem = Server.CreateObject("adodb.recordset")
    rsItem.Open sql, Conn, 1, 1

    If rsItem.EOF And rsItem.BOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>您还没有,可采集的项目,请添加审核。</li>"
    Else

        Do While Not rsItem.EOF

            ItemID = rsItem("ItemID")
            ItemName = rsItem("ItemName")
            WebUrl = rsItem("WebUrl")
            ListStr = rsItem("ListStr")
            LsString = rsItem("LsString")
            LoString = rsItem("LoString")
            ListPaingType = rsItem("ListPaingType")
            LPsString = rsItem("LPsString")
            LPoString = rsItem("LPoString")
            ListPaingStr1 = rsItem("ListPaingStr1")
            ListPaingStr2 = rsItem("ListPaingStr2")
            ListPaingID1 = rsItem("ListPaingID1")
            ListPaingID2 = rsItem("ListPaingID2")
            ListPaingStr3 = rsItem("ListPaingStr3")

            HsString = rsItem("HsString")
            HoString = rsItem("HoString")
            HttpUrlType = rsItem("HttpUrlType")
            HttpUrlStr = rsItem("HttpUrlStr")

            TsString = rsItem("TsString")
            ToString = rsItem("ToString")
            CsString = rsItem("CsString")
            CoString = rsItem("CoString")
                        
            If LsString = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>列表开始标记不能为空！无法继续,请返回上一步进行设置！</li>"
            End If

            If LoString = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>列表结束标记不能为空！无法继续,请返回上一步进行设置！</li>"
            End If

            If ListPaingType = 0 Or ListPaingType = 1 Then
                If ListStr = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>列表索引页不能为空！无法继续,请返回上一步进行设置！</li>"
                End If

                If ListPaingType = 1 Then
                    If LPsString = "" Or LPoString = "" Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>索引分页开始、结束标记不能为空！无法继续,请返回上一步进行设置！</li>"
                    End If
                End If

                If ListPaingStr1 <> "" And Len(ListPaingStr1) < 15 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>索引分页重定向设置不正确！无法继续,请返回上一步进行设置！</li>"
                End If

            ElseIf ListPaingType = 2 Then

                If ListPaingStr2 = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>批量生成原字符串不能为空！无法继续,请返回上一步进行设置</li>"
                End If

                If IsNumeric(ListPaingID1) = False Or IsNumeric(ListPaingID2) = False Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>批量生成的范围只能是数字！无法继续,请返回上一步进行设置</li>"
                Else

                    If ListPaingID1 = 0 And ListPaingID2 = 0 Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>批量生成的范围不正确！无法继续,请返回上一步进行设置</li>"
                    End If
                End If

            ElseIf ListPaingType = 3 Then

                If ListPaingStr3 = "" Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>索引分页不能为空！无法继续,请返回上一步进行设置</li>"
                End If

            Else
                FoundErr = True
                ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>请选择返回上一步设置索引分页类型</li>"
            End If

            If HsString = "" Or HoString = "" Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>链接开始/结束标记不能为空！无法继续,请返回上一步进行设置</li>"
            End If
                        
            If LoginType = 1 Then
                LoginData = UrlEncoding(LoginUser & "&" & LoginPass)
                LoginResult = PostHttpPage(LoginUrl, LoginPostUrl, LoginData, PE_CLng(WebUrl))

                If InStr(LoginResult, LoginFalse) > 0 Then
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>登录网站时发生错误,请确认登录信息的正确性！</li>"
                End If
            End If

            If FoundErr <> True Then

                Select Case ListPaingType

                    Case 0, 1
                        ListUrl = ListStr

                    Case 2
                        ListUrl = Replace(ListPaingStr2, "{$ID}", CStr(ListPaingID1))

                    Case 3

                        If InStr(ListPaingStr3, vbCrLf) > 0 Then
                            ListUrl = Left(ListPaingStr3, InStr(ListPaingStr3, vbCrLf))
                        Else
                            ListUrl = ListPaingStr3
                        End If

                End Select

            End If

            If FoundErr <> True Then
                ListCode = GetHttpPage(ListUrl, PE_CLng(WebUrl)) '获取网页源代码

                If ListCode <> "$False$" Then
                    ListCode = GetBody(ListCode, LsString, LoString, False, False) '获取列表页

                    If ListCode <> "$False$" Then
                        NewsArrayCode = GetArray(ListCode, HsString, HoString, False, False) '获取链接地址

                        If NewsArrayCode <> "$False$" Then
                            If InStr(NewsArrayCode, "$Array$") > 0 Then
                                NewsArray = Split(NewsArrayCode, "$Array$") '分割得到地址

                                If HttpUrlType = 1 Then
                                    NewsUrl = Trim(Replace(HttpUrlStr, "{$ID}", NewsArray(0)))
                                Else
                                    NewsUrl = Trim(DefiniteUrl(NewsArray(0), ListUrl)) '转为绝对路径
                                End If

                                NewsPaingNextCode = GetHttpPage(NewsUrl, PE_CLng(WebUrl)) '获取网页源代码

                                If NewsPaingType = 1 Then '当是设置代码分页时
                                    If NewsPaingStr1 <> "" And Len(NewsPaingStr1) > 15 Then
                                        '获取分页地址
                                        ListPaingNext = Replace(NewsPaingStr1, "{$ID}", GetPaing(NewsPaingNextCode, NPsString, NPoString, False, False))
                                    Else
                                        ListPaingNext = GetPaing(NewsPaingNextCode, NPsString, NPoString, False, False) '获取分页地址

                                        If ListPaingNext <> "$False$" Then
                                            ListPaingNext = DefiniteUrl(ListPaingNext, NewsUrl) '将相对路径转绝对路径
                                        End If
                                    End If

                                ElseIf NewsPaingType = 2 Then
                                    PageListCode = GetBody(NewsPaingNextCode, PsString, PoString, False, False) '获取列表页

                                    If PageListCode <> "$False$" Then
                                        PageArrayCode = GetArray(PageListCode, PhsString, PhoString, False, False) '获取链接地址

                                        If PageArrayCode <> "$False$" Then
                                            If InStr(PageArrayCode, "$Array$") > 0 Then
                                                PageArray = Split(PageArrayCode, "$Array$") '分割得到地址

                                                For i = 0 To UBound(PageArray)

                                                    If ListPaingNext = "" Then
                                                        ListPaingNext = DefiniteUrl(PageArray(i), NewsUrl) '将相对路径转绝对路径
                                                    Else
                                                        ListPaingNext = ListPaingNext & "$Array$" & DefiniteUrl(PageArray(i), NewsUrl) '将相对路径转绝对路径
                                                    End If

                                                    '去掉地址开始
                                                    Dim TempPaingArray, tempj
                                                    TempPaingArray = Split(ListPaingNext, "$Array$")
                                                    ListPaingNext = ""
                                                    For tempj = 0 To UBound(TempPaingArray)
                                                        If InStr(LCase(ListPaingNext), LCase(TempPaingArray(tempj))) < 1 Then
                                                            ListPaingNext = ListPaingNext & "$Array$" & TempPaingArray(tempj)
                                                        End If
                                                    Next
                                                    ListPaingNext = Right(ListPaingNext, Len(ListPaingNext) - 7)
                                                    '去掉地址结束

                                                Next

                                            Else
                                                ListPaingNext = DefiniteUrl(PageArrayCode, NewsUrl) '将相对路径转绝对路径
                                            End If

                                        Else
                                            FoundErr = True
                                            ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>在获取分页链接列表时出错。</li>"
                                        End If

                                    Else
                                        FoundErr = True
                                        ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>在截取分页代码发生错误。</li>"
                                    End If
                                End If

                            Else
                                FoundErr = True
                                ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>只发现一个有效链接？：" & NewsArrayCode & "</li>"
                            End If

                        Else
                            FoundErr = True
                            ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>在获取链接列表时出错。</li>"
                        End If

                    Else
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>在截取列表时发生错误。</li>"
                    End If

                Else
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li><font color=red>" & ItemName & "</font>在获取:" & ListUrl & "网页源码时发生错误。</li>"
                End If
            End If

            If FoundErr <> True Then
                NewsCode = GetHttpPage(NewsUrl, PE_CLng(WebUrl))

                If NewsCode <> "$False$" Then
                    Title = GetBody(NewsCode, TsString, ToString, False, False)
                    Content = GetBody(NewsCode, CsString, CoString, False, False)

                    If Title = "$False$" Or Content = "$False$" Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "<li>在截取标题/正文的时候发生错误：" & NewsUrl & "</li>"
                    End If

                Else
                    FoundErr = True
                    ErrMsg = ErrMsg & "<li>在获取源码时发生错误：" & NewsUrl & "</li>"
                End If
            End If

            If FoundErr = True Then
                Conn.Execute ("update PE_Item set Flag=" & PE_False & " where ItemID=" & ItemID)
            Else
                ErrMsg = "项目：<font color=red>" & ItemName & "</font>&nbsp;检测成功,没有发现任何问题"
            End If

            Response.Write "<br><br>" & vbCrLf
            Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
            Response.Write "  <tr align='center' class='title'><td height='22'><strong>" & ItemName & "项目检测如下</strong></td></tr>" & vbCrLf
            Response.Write "  <tr class='tdbg'><td height='100' valign='top' align='center'><br>" & ErrMsg & "</td></tr>" & vbCrLf
            Response.Write "  <tr align='center' class='tdbg'><td>"
            Response.Write "</td></tr>" & vbCrLf
            Response.Write "</table>" & vbCrLf
            FoundErr = False
            ErrMsg = ""

            Response.Flush
            rsItem.MoveNext
        Loop

    End If

    Response.Write "<br><center><a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a></center>"

    rsItem.Close
    Set rsItem = Nothing

    Call CloseConn
End Sub

'=================================================
'过程名：CreateItemHtml
'作  用：采集后自动生成HTML
'=================================================
Sub CreateItemHtml()
    Dim CollectionCreateHTML, CreateNum, CreateCount, TimingCreateUrl, TimingCreate, ArticleNum
    Dim arrContent, arrContent2

    CollectionCreateHTML = Trim(Request("CollectionCreateHTML"))
    TimingCreate = Trim(Request("TimingCreate"))
    CreateNum = PE_CLng(Trim(Request("CreateNum")))
    CreateCount = UBound(Split(CollectionCreateHTML, "|"))

    If CreateNum <= CreateCount Then
        If InStr(CollectionCreateHTML, "|") > 0 Then
            arrContent = Split(CollectionCreateHTML, "|")
            arrContent2 = Split(arrContent(CreateNum), "$")
            ChannelID = PE_CLng(arrContent2(0))
            ClassID = PE_CLng(arrContent2(1))
            SpecialID = ReplaceBadChar(arrContent2(2))
            ArticleNum = PE_CLng(arrContent2(3))
        Else
            arrContent2 = Split(CollectionCreateHTML, "$")
            ChannelID = PE_CLng(arrContent2(0))
            ClassID = PE_CLng(arrContent2(1))
            SpecialID = ReplaceBadChar(arrContent2(2))
            ArticleNum = PE_CLng(arrContent2(3))
        End If
        TimingCreateUrl = "Admin_CreateArticle.asp?Action=CreateArticle&CreateType=7&ChannelID=" & ChannelID & "&ClassID=" & ClassID & "&SpecialID=" & SpecialID & "&ArticleNum=" & ArticleNum & "&CreateNum=" & CreateNum & "&ShowBack=No&CollectionCreateHTML=" & CollectionCreateHTML & "&TimingCreate=" & TimingCreate
    Else
        If TimingCreate <> "" Then
            TimingCreateUrl = "Admin_Timing.asp?Action=DoTiming&TimingCreate=" & TimingCreate
        Else
            Response.Write "<html><head><title>成功信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
            Response.Write "<link href='images/Style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbCrLf
            Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
            Response.Write "  <tr align='center' class='title'><td height='22'><strong>恭喜您！</strong></td></tr>" & vbCrLf
            Response.Write "  <tr class='tdbg'><td height='100' valign='top' align='center'><br>&nbsp;采集所有项目生成完成!</td></tr>" & vbCrLf
            Response.Write "  <tr align='center' class='tdbg'><td>"
            Response.Write "</td></tr>" & vbCrLf
            Response.Write "</table>" & vbCrLf
            Exit Sub
        End If
    End If

    Response.Write "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    Response.Write "function aaa(){window.location.href='" & TimingCreateUrl & "';}" & vbCrLf
    Response.Write "    setTimeout('aaa()',5000);" & vbCrLf
    Response.Write "</script>" & vbCrLf
End Sub


'**************************************************
'函数名：GetManagePath
'作  用：采集项目导航
'参  数：iChannelID ----频道ID
'返回值：采集导航栏
'**************************************************
Function GetManagePath(ByVal iChannelID)
    Dim strPath, sqlPath, rsPath
    iChannelID = PE_CLng(iChannelID)
    strPath = "<IMG SRC='images/img_u.gif' height='12'>您现在的位置：采集" & ItemName & "管理&nbsp;&gt;&gt;&nbsp;"

    If iChannelID = 0 Then
        strPath = strPath & "<a href='" & strFileName & "&ChannelID=0'>所有频道项目</a>"
    Else
        sqlPath = "select ChannelID,ChannelName from PE_Channel where ChannelID=" & iChannelID
        Set rsPath = Conn.Execute(sqlPath)

        If rsPath.BOF And rsPath.EOF Then
            strPath = strPath & "错误的频道参数"
        Else
            strPath = strPath & "<a href='" & strFileName & "&ChannelID=" & rsPath(0) & "'>" & rsPath(1) & "项目</a>"
        End If

        rsPath.Close
        Set rsPath = Nothing
    End If

    GetManagePath = strPath
End Function
'=================================================
'过程名：SetCache
'作  用：加载项目缓存
'=================================================
Sub SetCache()
    rnd_temp = CStr(rnd_num(5))
    '获取项目信息
    sql = "select I.ItemID,I.ItemName,I.ChannelID,I.ClassID,I.SpecialID,I.WebName,I.WebUrl,"
    sql = sql & "I.LoginType,I.LoginUrl,I.LoginPostUrl,I.LoginUser,I.LoginPass,I.LoginFalse,I.ItemDoem,"
    sql = sql & "I.ListStr,I.LsString,I.LoString,I.ListPaingType,I.LPsString,I.LPoString,I.ListPaingStr1,"
    sql = sql & "I.ListPaingStr2,I.ListPaingID1,I.ListPaingID2,I.ListPaingStr3,"
    sql = sql & "I.HsString,I.HoString,I.HttpUrlType,I.HttpUrlStr,"
    sql = sql & "I.TsString,I.ToString,I.CsString,I.CoString,I.AuthorType,"
    sql = sql & "I.AuthorStr,I.AsString,I.AoString,I.CopyFromType,"
    sql = sql & "I.FsString,I.FoString,I.CopyFromStr,I.KeyType,I.KsString,I.KoString,I.KeyStr,I.KeyScatterNum,"
    sql = sql & "I.NewsPaingType,I.NPsString,I.NPoString,I.NewsPaingStr1,I.NewsPaingStr2,"
    sql = sql & "I.PaginationType,I.MaxCharPerPage,I.InfoPoint,"
    sql = sql & "I.OnTop,I.Hot,I.Elite,I.Hits,I.Stars,I.UpdateTime,"
    sql = sql & "I.SkinID,I.TemplateID,I.Script_Iframe,I.Script_Object,I.Script_Script,"
    sql = sql & "I.Script_Class,I.Script_Div,I.Script_Span,I.Script_Img,I.Script_Font,"
    sql = sql & "I.Script_A,I.Script_Html,I.SaveFiles,I.AddWatermark,I.AddThumb,I.CollecOrder,"
    sql = sql & "I.Status,I.CreateImmediate,I.IncludePicYn,I.DefaultPicYn,"
    sql = sql & "I.CollectionNum,I.CollectionType,I.UpDateType,I.DateType,I.DsString,"
    sql = sql & "I.DoString,I.ShowCommentLink,I.Script_Table,I.Script_Tr,I.Script_Td,"
    sql = sql & "I.PsString,I.PoString,I.PhsString,I.PhoString,"
    sql = sql & "I.IsString,I.IoString,I.IntroType,I.IntroStr,I.IntroNum,"
    sql = sql & "I.IsField,I.Field,"
    sql = sql & "I.InfoPurview,I.arrGroupID,I.ChargeType,"
    sql = sql & "I.PitchTime,I.ReadTimes,I.DividePercent,I.SaveFlashUrlToFile,ThumbnailType,ThsString,ThoString,"
    sql = sql & "C.ChannelDir,C.UploadDir,C.UpFileType,C.UseCreateHTML from PE_Item I left join PE_Channel C on I.ChannelID=C.ChannelID where I.ItemID in (" & ItemIDStr & ")"

    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 3
    If rs.EOF Then   '没有找到该项目
        ItemEnd = True
        ErrMsg = ErrMsg & "<li>参数错误,找不到该项目</li>"
    Else
        PE_Cache.SetValue "Collection" & rnd_temp, rs.GetRows()
    End If
    rs.Close
    Set rs = Nothing
    '加载过滤
    sql = "Select * from PE_Filters Where Flag=" & PE_True & " order by FilterID ASC"
    Set rs = Conn.Execute(sql)
    If rs.EOF And rs.BOF Then
    Else
        PE_Cache.SetValue "Arr_Filters" & rnd_temp, rs.GetRows()
    End If
    rs.Close
    Set rs = Nothing
    '加载历史记录
    sql = "select NewsUrl,Title,NewsCollecDate,Result from PE_HistrolyNews"
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open sql, Conn, 1, 1
    If Not rs.EOF Then
        PE_Cache.SetValue "Arr_Histrolys" & rnd_temp, rs.GetRows()
    End If
    rs.Close
    Set rs = Nothing
End Sub
'=================================================
'过程名：loadItem
'作  用：加载项目
'=================================================
Sub loadItem()
    Dim ItemNumTemp
    ItemNumTemp = ItemNum - 1

    ItemID = Arr_Item(0, ItemNumTemp)
    ItemName = Arr_Item(1, ItemNumTemp)
    ChannelID = Arr_Item(2, ItemNumTemp)
    ClassID = Arr_Item(3, ItemNumTemp)      '栏目
    SpecialID = Arr_Item(4, ItemNumTemp)    '专题
    WebUrl = Arr_Item(6, ItemNumTemp)
    LoginType = Arr_Item(7, ItemNumTemp)
    LoginUrl = Arr_Item(8, ItemNumTemp)
    LoginPostUrl = Arr_Item(9, ItemNumTemp)
    LoginUser = Arr_Item(10, ItemNumTemp)
    LoginPass = Arr_Item(11, ItemNumTemp)
    LoginFalse = Arr_Item(12, ItemNumTemp)
    ListStr = Arr_Item(14, ItemNumTemp)      '列表地址
    LsString = Arr_Item(15, ItemNumTemp)     '列表
    LoString = Arr_Item(16, ItemNumTemp)
    ListPaingType = Arr_Item(17, ItemNumTemp)
    LPsString = Arr_Item(18, ItemNumTemp)
    LPoString = Arr_Item(19, ItemNumTemp)
    ListPaingStr1 = Arr_Item(20, ItemNumTemp)
    ListPaingStr2 = Arr_Item(21, ItemNumTemp)
    ListPaingID1 = Arr_Item(22, ItemNumTemp)
    ListPaingID2 = Arr_Item(23, ItemNumTemp)
    ListPaingStr3 = Arr_Item(24, ItemNumTemp)
    HsString = Arr_Item(25, ItemNumTemp)
    HoString = Arr_Item(26, ItemNumTemp)
    HttpUrlType = Arr_Item(27, ItemNumTemp)
    HttpUrlStr = Arr_Item(28, ItemNumTemp)
    TsString = Arr_Item(29, ItemNumTemp)        '标题
    ToString = Arr_Item(30, ItemNumTemp)
    CsString = Arr_Item(31, ItemNumTemp)        '正文
    CoString = Arr_Item(32, ItemNumTemp)
    AuthorType = Arr_Item(33, ItemNumTemp)              '作者
    AuthorStr = Arr_Item(34, ItemNumTemp)
    AsString = Arr_Item(35, ItemNumTemp)
    AoString = Arr_Item(36, ItemNumTemp)
    CopyFromType = Arr_Item(37, ItemNumTemp)    '来源
    FsString = Arr_Item(38, ItemNumTemp)
    FoString = Arr_Item(39, ItemNumTemp)
    CopyFromStr = Arr_Item(40, ItemNumTemp)
    KeyType = Arr_Item(41, ItemNumTemp)         '关键词
    KsString = Arr_Item(42, ItemNumTemp)
    KoString = Arr_Item(43, ItemNumTemp)
    KeyStr = Arr_Item(44, ItemNumTemp)
    KeyScatterNum = Arr_Item(45, ItemNumTemp)
    NewsPaingType = Arr_Item(46, ItemNumTemp)
    NPsString = Arr_Item(47, ItemNumTemp)
    NPoString = Arr_Item(48, ItemNumTemp)
    NewsPaingStr1 = Arr_Item(49, ItemNumTemp)
    NewsPaingStr2 = Arr_Item(50, ItemNumTemp)
    PaginationType = Arr_Item(51, ItemNumTemp)
    MaxCharPerPage = Arr_Item(52, ItemNumTemp)
    InfoPoint = Arr_Item(53, ItemNumTemp)
    OnTop = Arr_Item(54, ItemNumTemp)
    Hot = Arr_Item(55, ItemNumTemp)
    Elite = Arr_Item(56, ItemNumTemp)
    Hits = Arr_Item(57, ItemNumTemp)
    Stars = Arr_Item(58, ItemNumTemp)
    UpdateTime = Arr_Item(59, ItemNumTemp)
    SkinID = Arr_Item(60, ItemNumTemp)
    TemplateID = Arr_Item(61, ItemNumTemp)
    Script_Iframe = Arr_Item(62, ItemNumTemp)
    Script_Object = Arr_Item(63, ItemNumTemp)
    Script_Script = Arr_Item(64, ItemNumTemp)
    Script_Class = Arr_Item(65, ItemNumTemp)
    Script_Div = Arr_Item(66, ItemNumTemp)
    Script_Span = Arr_Item(67, ItemNumTemp)
    Script_Img = Arr_Item(68, ItemNumTemp)
    Script_Font = Arr_Item(69, ItemNumTemp)
    Script_A = Arr_Item(70, ItemNumTemp)
    Script_Html = Arr_Item(71, ItemNumTemp)
    SaveFiles = Arr_Item(72, ItemNumTemp)
    AddWatermark = Arr_Item(73, ItemNumTemp)
    AddThumb = Arr_Item(74, ItemNumTemp)
    CollecOrder = Arr_Item(75, ItemNumTemp)
    Status = Arr_Item(76, ItemNumTemp)
    CreateImmediate = Arr_Item(77, ItemNumTemp)
    IncludePicYn = Arr_Item(78, ItemNumTemp)
    DefaultPicYn = Arr_Item(79, ItemNumTemp)
    CollectionNum = Arr_Item(80, ItemNumTemp)
    CollectionType = Arr_Item(81, ItemNumTemp)
    UpDateType = Arr_Item(82, ItemNumTemp)
    DateType = Arr_Item(83, ItemNumTemp)
    DsString = Arr_Item(84, ItemNumTemp)
    DoString = Arr_Item(85, ItemNumTemp)
    ShowCommentLink = Arr_Item(86, ItemNumTemp)
    Script_Table = Arr_Item(87, ItemNumTemp)
    Script_Tr = Arr_Item(88, ItemNumTemp)
    Script_Td = Arr_Item(89, ItemNumTemp)
    PsString = Arr_Item(90, ItemNumTemp)
    PoString = Arr_Item(91, ItemNumTemp)
    PhsString = Arr_Item(92, ItemNumTemp)
    PhoString = Arr_Item(93, ItemNumTemp)
    IsString = Arr_Item(94, ItemNumTemp)
    IoString = Arr_Item(95, ItemNumTemp)
    IntroType = Arr_Item(96, ItemNumTemp)
    IntroStr = Arr_Item(97, ItemNumTemp)
    IntroNum = Arr_Item(98, ItemNumTemp)
    IsField = Arr_Item(99, ItemNumTemp)
    Field = Arr_Item(100, ItemNumTemp)
    InfoPurview = Arr_Item(101, ItemNumTemp)
    arrGroupID = Arr_Item(102, ItemNumTemp)
    ChargeType = Arr_Item(103, ItemNumTemp)
    PitchTime = Arr_Item(104, ItemNumTemp)
    ReadTimes = Arr_Item(105, ItemNumTemp)
    DividePercent = Arr_Item(106, ItemNumTemp)
    SaveFlashUrlToFile = Arr_Item(107, ItemNumTemp)
    ThumbnailType = Arr_Item(108, ItemNumTemp)
    ThsString = Arr_Item(109, ItemNumTemp)
    ThoString = Arr_Item(110, ItemNumTemp)

    FilesPath = InstallDir & Arr_Item(111, ItemNumTemp) & "/" & Arr_Item(112, ItemNumTemp) & "/" & dirMonth
    FilesOverStr = Arr_Item(113, ItemNumTemp)
    UseCreateHTML = Arr_Item(114, ItemNumTemp)
End Sub
'==================================================
'过程名：SaveArticle
'作  用：保存文章
'参  数：无
'==================================================
Sub SaveArticle()

    Dim rsArticle, mrs, arrSpecialID
    Set mrs = Conn.Execute("select max(ArticleID) from PE_Article")

    If IsNull(mrs(0)) Then
        ArticleID = 1
    Else
        ArticleID = mrs(0) + 1
    End If

    Set mrs = Nothing
         
    Set rsArticle = Server.CreateObject("adodb.recordset")
    sql = "select top 1 * from PE_Article"
    rsArticle.Open sql, Conn, 1, 3
    rsArticle.addnew
    rsArticle("ArticleID") = ArticleID
    rsArticle("ChannelID") = ChannelID
    rsArticle("ClassID") = ClassID

    If SpecialID = "" Then
        arrSpecialID = Split("0", ",")
    Else
        arrSpecialID = Split(SpecialID, ",")
    End If
    For i = 0 To UBound(arrSpecialID)
        Conn.Execute ("insert into PE_InfoS (ModuleType,ItemID,SpecialID) values (1," & ArticleID & "," & PE_CLng(arrSpecialID(i)) & ")")
    Next
    rsArticle("Title") = Title
    rsArticle("TitleFontType") = 0
    rsArticle("Intro") = Intro
    If CollecType = 2 Then
        If IsLink = True Then
            rsArticle("Content") = NewsUrl
        Else
            rsArticle("LinkUrl") = NewsUrl
            rsArticle("Content") = ""
        End If
    Else
        rsArticle("Content") = Content
    End If
    rsArticle("Keyword") = Key
    rsArticle("Hits") = Hits
    rsArticle("Author") = Author
    rsArticle("CopyFrom") = CopyFrom

    If IncludePicYn = True Then
        If IncludePic = "" Or IsNull(IncludePic) = True Then
            rsArticle("IncludePic") = 0
        Else
            rsArticle("IncludePic") = IncludePic
        End If
    Else
        rsArticle("IncludePic") = 0
    End If

    rsArticle("OnTop") = OnTop
    rsArticle("Elite") = Elite
    rsArticle("Stars") = Stars
    rsArticle("UpdateTime") = UpdateTime
    rsArticle("CreateTime") = UpdateTime
    rsArticle("PaginationType") = PaginationType
    rsArticle("MaxCharPerPage") = MaxCharPerPage
    rsArticle("SkinID") = SkinID
    rsArticle("TemplateID") = TemplateID

    rsArticle("DefaultPicUrl") = DefaultPicUrl

    If SaveFiles = True Then
        rsArticle("UploadFiles") = UploadFiles
    Else
        rsArticle("UploadFiles") = ""
    End If

    rsArticle("Inputer") = UserName
    rsArticle("Editor") = AdminName
    rsArticle("ShowCommentLink") = ShowCommentLink
    rsArticle("Status") = Status
    rsArticle("Deleted") = False
    rsArticle("PresentExp") = 0
    rsArticle("Receive") = False
    rsArticle("ReceiveUser") = ""
    rsArticle("Received") = ""
    rsArticle("AutoReceiveTime") = 0
    rsArticle("ReceiveType") = 0
    rsArticle("InfoPoint") = InfoPoint
    
    Dim rsField, DefaultValue, ArticleField
    Set rsField = Conn.Execute("select * from PE_Field where (ChannelID=-1 or ChannelID=" & ChannelID & ") and (FieldType=1 or FieldType=2)")
    If Not (rsField.BOF And rsField.EOF) Then
        rsField.MoveFirst

        Do While Not rsField.EOF
            If Trim(rsField("DefaultValue")) = "" Then
                DefaultValue = " "
            End If
            If IsField > 0 Then
                If InStr(Field, "|||") > 0 Then
                    arrField = Split(Field, "|||")

                    For iField = 0 To UBound(arrField)
                        arrField2 = Split(arrField(iField), "{#F}")

                        If rsField("FieldID") = PE_CLng(arrField2(0)) Then
                            FieldType = arrField2(2)
                            FisSting = arrField2(3)
                            FioSting = arrField2(4)
                            FieldStr = arrField2(5)

                            If FieldType = 0 Then
                            ElseIf FieldType = 1 Then
                                ArticleField = GetBody(NewsCode, FisSting, FioSting, False, False)
                                ArticleField = Trim(ArticleField)
                            ElseIf FieldType = 2 Then
                                ArticleField = FieldStr
                            End If

                            Exit For
                        End If

                    Next
                    If rsField("FieldType") = 1 Then
                        If Len(ArticleField) > 253 Then
                            ArticleField = Left(ArticleField, 253)
                        End If
                    End If

                    If Field = "" Or Field = "$False$" Then
                        rsArticle(Trim(rsField("FieldName"))) = DefaultValue
                    Else
                        rsArticle(Trim(rsField("FieldName"))) = ArticleField
                    End If

                Else

                    If InStr(Field, "{#F}") > 0 Then
                        arrField2 = Split(Field, "{#F}")
                        FieldType = PE_CLng(arrField2(2))
                        FisSting = arrField2(3)
                        FioSting = arrField2(4)
                        FieldStr = arrField2(5)

                        If rsField("FieldID") = PE_CLng(arrField2(0)) Then
                            If FieldType = 0 Then
                            ElseIf FieldType = 1 Then
                                ArticleField = GetBody(NewsCode, FisSting, FioSting, False, False)
                                ArticleField = Trim(ArticleField)
                            ElseIf FieldType = 2 Then
                                ArticleField = FieldStr
                            End If
                        End If
                    End If

                    If Field = "" Or Field = "$False$" Then
                        rsArticle(Trim(rsField("FieldName"))) = DefaultValue
                    Else
                        rsArticle(Trim(rsField("FieldName"))) = ArticleField
                    End If
                End If

            Else
                rsArticle(Trim(rsField("FieldName"))) = DefaultValue
            End If

            rsField.MoveNext
        Loop

    End If

    rsField.Close
    Set rsField = Nothing

    rsArticle("InfoPurview") = InfoPurview
    rsArticle("arrGroupID") = arrGroupID
    rsArticle("ChargeType") = ChargeType
    rsArticle("PitchTime") = PitchTime
    rsArticle("ReadTimes") = ReadTimes
    rsArticle("DividePercent") = DividePercent

    rsArticle.Update
    rsArticle.Close
    Set rsArticle = Nothing
    
    If Status = 3 Then
        Conn.Execute ("update PE_Channel set ItemCount=ItemCount+1,ItemChecked=ItemChecked+1 where ChannelID=" & ChannelID & "")
        Conn.Execute ("update PE_Class set ItemCount=ItemCount+1 where ClassID=" & ClassID & "")
    Else
        Conn.Execute ("update PE_Channel set ItemCount=ItemCount+1 where ChannelID=" & ChannelID & "")
    End If

End Sub
'=================================================
'过程名：HistoryNum
'作  用：反馈项目新闻数
'参数    Itemid 所属采集项目
'        Result 采集项目成功或失败
'=================================================
Sub HistrolyNum(ByVal ItemID, ByVal Result)
    If IsNumeric(ItemID) = False Then
        Response.Write "采集项目不存在"
        Exit Sub
    End If
    sql = "select count(HistrolyNewsID) from PE_HistrolyNews where ItemID=" & ItemID & " and Result=" & Result
    Set rs = Conn.Execute(sql)
    If rs.BOF And rs.EOF Then
        Response.Write "<font color='green'>0</font>"
    Else
        If Result = PE_True Then
            Response.Write "<font color='blue'>" & rs(0) & "</font>"
        Else
            If rs(0) = 0 Then
                Response.Write "<font color='green'>" & rs(0) & "</font>"
            Else
                Response.Write "<font color='red'>" & rs(0) & "</font>"
            End If
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub
'=================================================
'过程名：GetArrOfCreateHTML
'作  用：生成HTML赋值
'=================================================
Sub GetArrOfCreateHTML()
    If CreateImmediate = True And UseCreateHTML <> 0 And ItemSucceedNum2 <> 0 Then
        If CollectionCreateHTML = "" Then
            CollectionCreateHTML = ChannelID & "$" & ClassID & "$" & SpecialID & "$" & ItemSucceedNum2
        Else
            CollectionCreateHTML = CollectionCreateHTML & "|" & ChannelID & "$" & ClassID & "$" & SpecialID & "$" & ItemSucceedNum2
        End If
    End If
End Sub
'=================================================
'过程名：WriteSuccessMsg2
'作  用：采集成功信息
'=================================================
Sub WriteSuccessMsg2(ErrMsg)
    ItemSucceedNum = 0
    ItemSucceedNum2 = 0
    ArticleList = ""
    ItemNum = ItemNum + 1
    ListNum = 1
    Response.Write "<html><head><title>成功信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    Response.Write "<link href='images/Style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbCrLf
    Response.Write "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    Response.Write "  <tr align='center' class='title'><td height='22'><strong>恭喜您！</strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='100' valign='top' align='center'><br>" & ErrMsg & "</td></tr>" & vbCrLf
    Response.Write "  <tr align='center' class='tdbg'><td>"
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<br>" & vbCrLf

    If ItemEnd = False Then
        Call Refresh("Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,TimeNum)	
        'Response.Write "<meta http-equiv=""refresh"" content=" & TimeNum & ";url=""Admin_Collection.asp?Action=Start&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&Arr_i=" & Arr_i & "&CollecNewsA=" & CollecNewsA & "&CollecNewsi=" & CollecNewsi & "&CollecNewsj=" & CollecNewsj & "&IsTitle=" & Trim(Request("IsTitle")) & "&IsLink=" & Trim(Request("IsLink")) & "&ItemIDtemp=" & ItemIDtemp & "&rnd_temp=" & rnd_temp & "&ArticleList=" & Replace(ArticleList,"/","|") & "&ItemSucceedNum=" & ItemSucceedNum & "&ItemSucceedNum2=" & ItemSucceedNum2 & "&ImagesNumAll=" & ImagesNumAll & "&ItemID=" & ItemIDStr & "&CollecType=" & CollecType & "&CollecTest=" & Trim(Request("CollecTest")) & "&Content_view=" & Trim(Request("Content_view")) & "&ListPaingNext=" & Replace(ListPaingNext,"/","|") & "&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
    Else
        PE_Cache.DelCache ("Collection" & Request("rnd_temp"))
        PE_Cache.DelCache ("ArticleList" & Request("rnd_temp"))
        PE_Cache.DelCache ("ThumbnailList" & Request("rnd_temp"))
        PE_Cache.DelCache ("Arr_Filters" & Request("rnd_temp"))
        PE_Cache.DelCache ("Arr_Histrolys" & Request("rnd_temp"))

        Dim arrContent, arrContent2, C_ChannelID, C_ClassID, C_SpecialID, C_ArticleNum
        If CollectionCreateHTML <> "" Then
            Response.Write "<center><FONT style='font-size:12px' color='red'>请稍等,5秒钟后系统开始生成采集后的文章。</FONT></center>"
            Call Refresh("Admin_Collection.asp?Action=CreateItemHtml&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,5)		
            'Response.Write "<meta http-equiv=""refresh"" content=5;url=""Admin_Collection.asp?Action=CreateItemHtml&CollectionCreateHTML=" & CollectionCreateHTML & "&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
        Else
            If Timing_AreaCollection = "1" Then
                Response.Write "<center><FONT style='font-size:12px' color='red'>请稍等,5秒钟后系统开始生成区域采集。</FONT></center>"
                Call Refresh("Admin_AreaCollection.asp?Action=AreaCollectionCreateFile&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate,5)				
                'Response.Write "<meta http-equiv=""refresh"" content=5;url=""Admin_AreaCollection.asp?Action=AreaCollectionCreateFile&Timing_AreaCollection=" & Timing_AreaCollection & "&TimingCreate=" & TimingCreate & """>"
            Else
                If TimingCreate <> "" Then
                    Response.Write "<center><FONT style='font-size:12px' color='red'>请稍等,5秒钟后系统开始定时生成。</FONT></center>"
                    Call Refresh("Admin_Timing.asp?Action=DoTiming&TimingCreate=" & TimingCreate,5)					
                    'Response.Write "<meta http-equiv=""refresh"" content=5;url=""Admin_Timing.asp?Action=DoTiming&TimingCreate=" & TimingCreate & """>"
                End If
            End If
        End If
    End If
End Sub




'==================================================
'函数名：UrlEncoding
'作  用：转换编码
'==================================================
Function UrlEncoding(DataStr)
    On Error Resume Next
    Dim StrReturn, Si, ThisChr, InnerCode, Hight8, Low8
    StrReturn = ""
    For Si = 1 To Len(DataStr)
        ThisChr = Mid(DataStr, Si, 1)
        If Abs(Asc(ThisChr)) < &HFF Then
            StrReturn = StrReturn & ThisChr
        Else
            InnerCode = Asc(ThisChr)
            If InnerCode < 0 Then
                InnerCode = InnerCode + &H10000
            End If
            Hight8 = (InnerCode And &HFF00) \ &HFF
            Low8 = InnerCode And &HFF
            StrReturn = StrReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
        End If
    Next
    UrlEncoding = StrReturn
End Function

'==================================================
'函数名：ReplaceBadChar2
'作  用：替换正则表达式特殊字符
'参  数：strChar-----要过滤的字符
'返回值：替换后的字符
'==================================================
Function ReplaceBadChar2(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceBadChar2 = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "^,(,),*,?,[,],$,+,|,{,}"
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "\" & arrBadChar(i))
    Next
    ReplaceBadChar2 = tempChar
End Function

'==================================================
'函数名：GetPaing
'作  用：获取分页
'参  数：ConStr   ------要找的内容
'参  数：StartStr ------链接网址头部
'参  数：OverStr  ------链接网址尾部
'参  数：IncluL   ------是否统计网址头部
'参  数：IncluR   ------是否统计网址尾部
'==================================================
Function GetPaing(ByVal ConStr, StartStr, OverStr, IncluL, IncluR)
    If ConStr = "$False$" Or ConStr = "" Or StartStr = "" Or OverStr = "" Or IsNull(ConStr) = True Or IsNull(StartStr) = True Or IsNull(OverStr) = True Then
        GetPaing = "$False$"
        Exit Function
    End If
    Dim Start, Over, ConTemp, tempStr
    tempStr = LCase(ConStr)
    StartStr = LCase(StartStr)
    OverStr = LCase(OverStr)
    Over = InStr(1, tempStr, OverStr)
    If Over <= 0 Then
        GetPaing = "$False$"
        Exit Function
    Else
        If IncluR = True Then
            Over = Over + Len(OverStr)
        End If
    End If
    tempStr = Mid(tempStr, 1, Over)
    Start = InStrRev(tempStr, StartStr)
    If IncluL = False Then
        Start = Start + Len(StartStr)
    End If
    
    If Start <= 0 Or Start >= Over Then
        GetPaing = "$False$"
        Exit Function
    End If
    ConTemp = Mid(ConStr, Start, Over - Start)
    ConTemp = Trim(ConTemp)
    ConTemp = Replace(ConTemp, " ", "%20")
    ConTemp = Replace(ConTemp, ",", "")
    ConTemp = Replace(ConTemp, "'", "")
    ConTemp = Replace(ConTemp, """", "")
    ConTemp = Replace(ConTemp, ">", "")
    ConTemp = Replace(ConTemp, "<", "")
    ConTemp = Replace(ConTemp, "&nbsp;", "")
    GetPaing = ConTemp
End Function
'==================================================
'过程名：Filters
'作  用：过滤
'==================================================
Sub Filters()
    If IsNull(Arr_Filters) = True Or IsArray(Arr_Filters) = False Then
        Exit Sub
    End If
    For Filteri = 0 To UBound(Arr_Filters, 2)
        FilterStr = ""
        If Arr_Filters(1, Filteri) = ItemID Or Arr_Filters(1, Filteri) = -1 And Arr_Filters(9, Filteri) = True Then
            If Arr_Filters(3, Filteri) = 1 Then '标题过滤
                If Arr_Filters(4, Filteri) = 1 Then
                    Title = Replace(Title, Arr_Filters(5, Filteri), Arr_Filters(8, Filteri))
                ElseIf Arr_Filters(4, Filteri) = 2 Then
                    FilterStr = GetBody(Title, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
                    Do While FilterStr <> "$False$"
                        Title = Replace(Title, FilterStr, Arr_Filters(8, Filteri))
                        FilterStr = GetBody(Title, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
                    Loop
                End If
            ElseIf Arr_Filters(3, Filteri) = 2 Then '正文过滤
                If Arr_Filters(4, Filteri) = 1 Then
                    Content = Replace(Content, Arr_Filters(5, Filteri), Arr_Filters(8, Filteri))
                ElseIf Arr_Filters(4, Filteri) = 2 Then
                    FilterStr = GetBody(Content, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
                    Do While FilterStr <> "$False$"
                        Content = Replace(Content, FilterStr, Arr_Filters(8, Filteri))
                        FilterStr = GetBody(Content, Arr_Filters(6, Filteri), Arr_Filters(7, Filteri), True, True)
                    Loop
                End If
            End If
        End If
    Next
End Sub

'**************************************************
'函数名：CreateKeyWord
'作  用：由给定的字符串生成关键字
'参  数：Constr---要生成关键字的原字符串
'返回值：生成的关键字
'**************************************************
Function CreateKeyWord(ByVal ConStr, ByVal Num)
    If ConStr = "" Or IsNull(ConStr) = True Or ConStr = "$False$" Or IsNumeric(Num) = False Then
        CreateKeyWord = "$False$"
        Exit Function
    End If
    If CLng(Num) < 2 Then
        Num = 2
    End If
    ConStr = Replace(ConStr, Chr(32), "")
    ConStr = Replace(ConStr, Chr(9), "")
    ConStr = Replace(ConStr, "&nbsp;", "")
    ConStr = Replace(ConStr, " ", "")
    ConStr = Replace(ConStr, "(", "")
    ConStr = Replace(ConStr, ")", "")
    ConStr = Replace(ConStr, "<", "")
    ConStr = Replace(ConStr, ">", "")
    Dim i, ConstrTemp
    If Num >= Len(ConStr) Then
        CreateKeyWord = "|" & Left(ConStr, 254) & "|"
        Exit Function
    Else
        For i = 1 To Len(ConStr)
            If i + Num > Len(ConStr) Then
                Exit For
            Else
                ConstrTemp = ConstrTemp & "|" & Mid(ConStr, i, Num)
            End If
        Next
    End If
    If Len(ConstrTemp) < 254 Then
        ConstrTemp = ConstrTemp & "|"
    Else
        ConstrTemp = Left(ConstrTemp, 254) & "|"
    End If
    CreateKeyWord = ConstrTemp
End Function

'==================================================
'函数名：ReplaceSaveRemoteFile
'作  用：替换、保存远程文件
'参  数：ConStr ------ 要替换的字符串
'参  数：OverStr ----- 保存的文件后缀名
'参  数：SaveFiles ------ 是否保存文件,False不保存,True保存
'参  数：SaveFilePath- 保存文件夹
'参  数: TistUrl------ 当前网页地址
'==================================================
Function ReplaceSaveRemoteFile(ConStr, OverStr, SaveFiles, SaveFilePath, TistUrl)
    'On Error Resume Next
    If IsObjInstalled("Microsoft.XMLHTTP") = False Then
        ReplaceSaveRemoteFile = ConStr
        Exit Function
    End If
    If ConStr = "$False$" Or ConStr = "" Then '内容为空或假退出
        ReplaceSaveRemoteFile = "$False$"
        Exit Function
    End If
    OverStr = Replace(Replace(OverStr, "$", ""), " ", "")
    Dim tempStr, TempStr2, tempi, TempArray, TempArray2
    Dim RemoteFileUrl, temptime, SaveFileName, ThumbFileName, StrFileType

    IncludePic = 0

    regEx.Pattern = "<img.+?[^\>]>" '查询内容中所有 <img..>
    Set Matches = regEx.Execute(ConStr)
    For Each Match In Matches
        If tempStr <> "" Then
            tempStr = tempStr & "$Array$" & Match.value '累计数组
        Else
            tempStr = Match.value
        End If
    Next
    If tempStr <> "" Then
        TempArray = Split(tempStr, "$Array$") '分割数组
        tempStr = ""
        For tempi = 0 To UBound(TempArray)
            regEx.Pattern = "src\s*=\s*.+?\.(" & OverStr & ")" '查询src =内的链接
            Set Matches = regEx.Execute(TempArray(tempi))
            For Each Match In Matches
                If tempStr <> "" Then
                    tempStr = tempStr & "$Array$" & Match.value '累加得到 链接加$Array$ 字符
                    IncludePic = 2  '组图
                Else
                    tempStr = Match.value
                    IncludePic = 1  '图文
                End If
            Next
        Next
    End If
    If tempStr <> "" Then
        regEx.Pattern = "src\s*=\s*" '过滤 src =
        tempStr = regEx.Replace(tempStr, "")
    End If

    If ThumbnailType = 1 And ThumbnailUrl <> "" Then '采集列表缩略图
        If InStr(tempStr, "$Array$") > 0 Then
            tempStr = ThumbnailUrl & "$Array$" & tempStr
        Else
            tempStr = ThumbnailUrl
        End If
    End If

    If tempStr = "" Or IsNull(tempStr) = True Then '如果处理后这里没有图片原值返回
        ReplaceSaveRemoteFile = ConStr
        Exit Function
    End If

    tempStr = Replace(tempStr, """", "")
    tempStr = Replace(tempStr, "'", "") '过滤图片组的'"
    If Right(SaveFilePath, 1) = "/" Then '保存文件夹右边不能有/
        SaveFilePath = Left(SaveFilePath, Len(SaveFilePath) - 1)
    End If
    If SaveFiles = True Then
        If Not fso.FolderExists(Server.MapPath(SaveFilePath)) Then
            fso.CreateFolder Server.MapPath(SaveFilePath)
            If fso.FolderExists(Server.MapPath(SaveFilePath)) Then
                SaveFiles = True
            Else
                SaveFiles = False
            End If
        End If
    End If
    SaveFilePath = SaveFilePath & "/"
    '去掉重复图片开始
    TempArray = Split(tempStr, "$Array$")
    tempStr = ""
    For tempi = 0 To UBound(TempArray)
        If InStr(LCase(tempStr), LCase(TempArray(tempi))) < 1 Then
            tempStr = tempStr & "$Array$" & TempArray(tempi)
        End If
    Next
    tempStr = Right(tempStr, Len(tempStr) - 7)
    TempArray = Split(tempStr, "$Array$")
    '去掉重复图片结束
    '转换相对图片地址开始
    tempStr = ""
    For tempi = 0 To UBound(TempArray)
        tempStr = tempStr & "$Array$" & DefiniteUrl(Trim(TempArray(tempi)), TistUrl)
    Next
    tempStr = Right(tempStr, Len(tempStr) - 7)
    tempStr = Replace(tempStr, Chr(0), "")

    TempArray2 = Split(tempStr, "$Array$")
    tempStr = ""

    Dim IsThumb

    '转换相对图片地址结束
    '图片转换/保存
    For tempi = 0 To UBound(TempArray2)
        IsThumb = False
        RemoteFileUrl = TempArray2(tempi)
        If RemoteFileUrl <> "$False$" And RemoteFileUrl <> "" And SaveFiles = True Then '保存图片
            StrFileType = GetFileExt(RemoteFileUrl)

             '如果文件类型 是动态型退出
            If StrFileType = "asp" Or StrFileType = "asa" Or StrFileType = "aspx" Or StrFileType = "php" Or StrFileType = "cer" Or StrFileType = "cdx" Or StrFileType = "exe" Then
                UploadFiles = ""
                ReplaceSaveRemoteFile = ConStr
                Exit Function
            End If
            temptime = GetNumString()
            SaveFileName = temptime & "." & StrFileType      '建立文件名
            ThumbFileName = temptime & "_S." & StrFileType  '建立文件名
            
            If SaveRemoteFile(RemoteFileUrl, SaveFilePath & SaveFileName) = True Then
                ConStr = Replace(ConStr, TempArray(tempi), "[InstallDir_ChannelDir]{$UploadDir}/" & dirMonth & "/" & SaveFileName) '替换原来的位置
                If PhotoObject = 1 Then
                    Dim PE_Thumb
                    Set PE_Thumb = New CreateThumb
                    If tempi = 0 And AddThumb = True Then
                        If PE_Thumb.CreateThumb(SaveFilePath & SaveFileName, SaveFilePath & ThumbFileName, 0, 0) = True Then
                            IsThumb = True
                        End If
                    End If
                    If AddWatermark = True Then
                        Call PE_Thumb.AddWatermark(SaveFilePath & SaveFileName)
                    End If
                    Set PE_Thumb = Nothing
                End If

                If IsThumb = True Then
                    UploadFiles = dirMonth & "/" & ThumbFileName & "|" & dirMonth & "/" & SaveFileName
                Else
                    If UploadFiles = "" Then
                        UploadFiles = dirMonth & "/" & SaveFileName
                    Else
                        UploadFiles = UploadFiles & "|" & dirMonth & "/" & SaveFileName
                    End If
                End If
                If PE_CLng(Trim(Request.Form("IncludePic"))) = 0 Then
                    If tempi > 0 Then
                        IncludePic = 2
                    Else
                        IncludePic = 1
                    End If
                Else
                    IncludePic = PE_CLng(Trim(Request.Form("IncludePic")))
                End If

                If InStr(UploadFiles, "|") = 0 Then
                    DefaultPicUrl = UploadFiles
                    ImagesNum = 1
                Else
                    FilesArray = Split(UploadFiles, "|")
                    DefaultPicUrl = FilesArray(0)
                    ImagesNum = UBound(FilesArray) + 1
                End If
                ImagesNumAll = ImagesNumAll + ImagesNum '采集图片统计
            End If            
        ElseIf RemoteFileUrl <> "$False$" And SaveFiles = False Then '不保存图片
            SaveFileName = RemoteFileUrl '文件名等于原来的图片字符
            ConStr = Replace(ConStr, TempArray(tempi), SaveFileName)
            UploadFiles = ""
            ImagesNum = 0
            If tempi =0 Then
                DefaultPicUrl = RemoteFileUrl
            End If
        End If
    Next
    ReplaceSaveRemoteFile = ConStr
End Function
'==================================================
'函数名：CollectionFilePath
'作  用：保存采集 图片 附件 控件 的绝对路径到文本
'参  数：ConStr ------ 要替换的字符串
'参  数: TistUrl------ 当前网页地址
'==================================================
Function CollectionFilePath(ConStr, TistUrl)
    On Error Resume Next
    If ConStr = "$False$" Or ConStr = "" Or TistUrl = "" Or TistUrl = "$False$" Then
        CollectionFilePath = ConStr
        Exit Function
    End If
    Dim tempStr, TempStr2, TempStr3, tempi, TempArray, TempArray2, RemoteFileUrl
    Dim f, IsSwftxt, FileName
    IsSwftxt = True
    FileName = InstallDir & "CollectionFilePath.txt"
       
    regEx.Pattern = "(<)(img|param|embed|flash8)(.[^\<]*)(value|src|href)(\s*=)(.[^\<]*)(\.)(gif|jpg|jpeg|jpe|bmp|png|swf|mid|mp3|wmv|asf|avi|mpg|ram|rm|ra|rmvb)"
    Set Matches = regEx.Execute(ConStr)
    For Each Match In Matches
        If tempStr <> "" Then
            tempStr = tempStr & "$Array$" & Match.value
        Else
            tempStr = Match.value
        End If
    Next
    If tempStr <> "" Then
        regEx.Pattern = "(<)(img|param|embed|flash8)(.[^\<]*)(value|src|href)(\s*=)"
        tempStr = regEx.Replace(tempStr, "")
    End If
        
    If tempStr = "" Or IsNull(tempStr) = True Then
        CollectionFilePath = ConStr
        Exit Function
    End If
    
    tempStr = Replace(tempStr, """", "")
    tempStr = Replace(tempStr, "'", "")

    '去掉重复文件开始
    TempArray = Split(tempStr, "$Array$")
    tempStr = ""

    For tempi = 0 To UBound(TempArray)
        If InStr(LCase(tempStr), LCase(TempArray(tempi))) < 1 Then
            tempStr = tempStr & "$Array$" & TempArray(tempi)
        End If
    Next
    tempStr = Right(tempStr, Len(tempStr) - 7)
    TempArray = Split(tempStr, "$Array$")
    '去掉重复文件结束
    '转换相对地址开始
    tempStr = ""
    For tempi = 0 To UBound(TempArray)
        tempStr = tempStr & "$Array$" & DefiniteUrl(TempArray(tempi), TistUrl)
    Next
    tempStr = Right(tempStr, Len(tempStr) - 7)
    tempStr = Replace(tempStr, Chr(0), "")
    TempArray2 = Split(tempStr, "$Array$")
    tempStr = ""
    '转换相对地址结束
    '替换

    If Not fso.FileExists(Server.MapPath(FileName)) Then
        fso.CreateTextFile (Server.MapPath(FileName))
        If fso.FileExists(Server.MapPath(FileName)) Then
            IsSwftxt = False
        End If
    End If

    If IsSwftxt = True Then
        For tempi = 0 To UBound(TempArray2)
            RemoteFileUrl = TempArray2(tempi)
            regEx.Pattern = TempArray(tempi)
            ConStr = regEx.Replace(ConStr, RemoteFileUrl)
            
            Set f = fso.OpenTextFile(Server.MapPath(FileName), 8, 0)
            If SwfTime = "" Then
                f.Write Chr(13) & Chr(10) & " " & ItemName & Now
                SwfTime = Now
            End If
            f.Write Chr(13) & Chr(10) & RemoteFileUrl
        Next
    End If

    f.Close
    Set f = Nothing
    CollectionFilePath = ConStr
End Function


'**************************************************
'函数名：CheckRepeat
'作  用：检测历史记录有无重复
'参  数：strUrl---网站Url
'返回值：True --- 有
'**************************************************
Function CheckRepeat(strUrl)
    CheckRepeat = False
    If IsArray(Arr_Histrolys) = True Then
        For His_i = 0 To UBound(Arr_Histrolys, 2)
            If Arr_Histrolys(0, His_i) = strUrl Then
                CheckRepeat = True
                His_Title = Arr_Histrolys(1, His_i)
                His_NewsCollecDate = Arr_Histrolys(2, His_i)
                His_Result = Arr_Histrolys(3, His_i)
                Exit For
            End If
        Next
    End If
End Function
'=================================================
'过程名：rnd_num
'作  用：产生指定位置的随机数
'参  数：产生的随机数  ----内容
'=================================================
Function rnd_num(rLen)
    Dim ri, rmax, rmin
    rmax = 1
    rmin = 1
    For ri = 1 To rLen + 1
        rmax = rmax * 10
    Next
    rmax = rmax - 1
    For ri = 1 To rLen
        rmin = rmin * 10
    Next
    Randomize
    rnd_num = Int((rnd_num - rmin + 1) * Rnd) + rmin
End Function
%>
