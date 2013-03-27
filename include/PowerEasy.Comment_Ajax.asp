<!--#include file="../Start.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************
Response.Charset="gb2312"

Dim ModuleName, InfoID, Titlelen, Tablelen, ModuleType, rsComment,sqlComment
ModuleName = PE_HTMLEncode(Request.QueryString("ModuleName"))
InfoID = PE_HTMLEncode(Request.QueryString("InfoID"))
Titlelen = PE_HTMLEncode(Request.QueryString("Titlelen"))
Tablelen = PE_HTMLEncode(Request.QueryString("Tablelen"))

Dim int_Start,NumberLink,NoActiveLinkColor,toF_,toP10_,toP1_,toN1_,toN10_,toL_,PageGoType,cPageNo,jsFun
MaxPerPage = PE_HTMLEncode(Request.QueryString("MaxPerPage")) '设置每页显示数目
NumberLink=8 '数字导航显示数目
PageGoType = 0 '是下拉菜单还是输入值跳转，当多次调用时只能选1
NoActiveLinkColor="#999999" '非热链接颜色
toF_="<font face=webdings>9</font>"  			'首页 
toP10_=" <font face=webdings>7</font>"			'上十
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"				'尾页

If ModuleName = "" Then
    Response.Write "请指定ModuleName"
    Response.End
End If
If Titlelen = "" Then
	Titlelen = 20
Else
	Titlelen = PE_CLng(Titlelen)
End If
If Tablelen = "" Then Tablelen = "100%"
If Tablelen <> "100%" Then Tablelen = Tablelen & "px"

sqlComment = "select ModuleType from PE_Channel where ChannelDir = '"&ModuleName&"'"
Set rsComment = Conn.Execute(sqlComment)
If rsComment.Eof or rsComment.Bof Then
	Response.Write "请指定正确的ModuleName"
	Response.End
End If
ModuleType = rsComment("ModuleType")
rsComment.Close:Set rsComment = Nothing

If InfoID = "" Then
    Response.Write "请指定ID"
    Call CloseConn
    Response.End
Else
    InfoID = PE_CLng(InfoID)
End If

Response.Write "<table width="""&Tablelen&""" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""commentTable"">"
Response.Write "  <tr class=""commentTitle"">"
Response.Write "    <th width=""5%"">序号</th>"
Response.Write "    <th>摘要</th>"
Response.Write "    <th width=""15%"">评论人</th>"
Response.Write "    <th width=""5%"">评分</th>"
Response.Write "    <th width=""15%"">日期</th>"
Response.Write "    <th width=""5%"">状态</th>"
Response.Write "    <th width=""10%"">回复人</th>"
Response.Write "    <th width=""5%"" class=""commentTdEnd"">展开</th>"
Response.Write "  </tr>"
If SystemDatabaseType = "SQL" Then
sqlComment = "select CommentID,ModuleType,InfoID,UserType,UserName,Email,IP,WriteTime,Score,Content,IsNull(ReplyName,'&nbsp;') as ReplyName,ReplyContent,ReplyTime from PE_Comment where ModuleType = "&ModuleType&" and InfoID = "&InfoID&" and Passed = 1 Order By CommentID desc"
Else
sqlComment = "select CommentID,ModuleType,InfoID,UserType,UserName,Email,IP,WriteTime,Score,Content,iif( IsNull(ReplyName),'&nbsp;', ReplyName) as ReplyName,ReplyContent,ReplyTime from PE_Comment where ModuleType = "&ModuleType&" and InfoID = "&InfoID&" and Passed = 1 Order By CommentID desc"
End If
Set rsComment = CreateObject("Adodb.RecordSet")
rsComment.Open sqlComment,Conn,1,1
If rsComment.EOF or rsComment.BOF Then
	Response.Write "  <tr>"
	Response.Write "    <td colspan=""8"">暂时没有评论</td>"
	Response.Write "  </tr>"
Else
	Dim ReplyStatus
	rsComment.PageSize = MaxPerPage
	cPageNo = PE_HTMLEncode(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = PE_CLng(cPageNo)
	If cPageNo <= 0 Then cPageNo = 1
	If cPageNo > rsComment.PageCount Then cPageNo = rsComment.PageCount 
	rsComment.AbsolutePage = cPageNo
	For int_Start = 1 To MaxPerPage
		If rsComment("ReplyName") = "&nbsp;" Then
			ReplyStatus = "×"
		Else
			ReplyStatus = "√"
		End If
		Response.Write "  <tr onmouseover= ""cc(this);""   onmouseout= ""bc(this);"">"
		Response.Write "    <td>"&rsComment("CommentID")&"</td>"
		Response.Write "    <td>"&GetSubStr(rsComment("Content"),Titlelen,True)&"</td>"
		Response.Write "    <td>"&rsComment("UserName")&"</td>"
		Response.Write "    <td>"&rsComment("Score")&"</td>"
		Response.Write "    <td>"&FormatDateTime(rsComment("WriteTime"),2)&"</td>"
		Response.Write "    <td>"&ReplyStatus&"</td>"
		Response.Write "    <td>"&rsComment("ReplyName")&"</td>"
		Response.Write "    <td class=""commentTdEnd""><img id=""commentImg"&rsComment("CommentID")&""" onclick=""showComment('"&rsComment("CommentID")&"')"" src="""&strInstallDir&"images/open.gif"" alt=""展开或隐藏评论""></td>"
		Response.Write "  </tr>"
		Response.Write "  <tr id=""commentTr"&rsComment("CommentID")&""" style=""display:none;"">"
		Response.Write "    <td class=""commentTdEnd"">&nbsp;</td>"
		Response.Write "    <td colspan=""7"">"
		Response.Write "    <div style=""float:left;text-align:left;width:100%;border-bottom:#ccc 1px solid;"">"
		Response.Write "    <div style=""float:left;width:300px;color:red;"">【"&rsComment("UserName")&"】&nbsp;"&rsComment("IP")&"</div>"
		Response.Write "    <div style=""float:right;text-align:right;width:180px;color:#063;"">"&rsComment("WriteTime")&"&nbsp;发表</div>"
		Response.Write "    </div>"
		Response.Write "    <div style=""float:left;text-align:left;padding:10px 0px 10px 0px;width:100%;border-bottom:#ccc 1px solid;color:#063;""><b>评论:</b>"&rsComment("Content")&"</div>"
		If rsComment("ReplyName") <> "&nbsp;" Then
			Response.Write "    <div style=""float:left;text-align:left;margin-top:5px;width:100%;border:#97d2df 1px dashed;background:#e8f5f8;"">"
			Response.Write "    <div style=""float:left;width:100%;color:red;"">管理员["&rsComment("ReplyName")&"]于"&rsComment("ReplyTime")&"回复:</div>"
			Response.Write "    <div style=""float:left;width:100%;"">"&rsComment("ReplyContent")&"</div>"
			Response.Write "    </div>"
		End If
		Response.Write "    </td>"
		Response.Write "  </tr>"
		rsComment.MoveNext
		if rsComment.EOF or rsComment.BOF then Exit For
	Next
	Response.Write "  <tr>"
	Response.Write "    <td colspan=""8"" class=""commentPager"">"&fPageCount(rsComment,NumberLink,NoActiveLinkColor,toF_,toP10_,toP1_,toN1_,toN10_,toL_,PageGoType,cPageNo)&"</td>"
	Response.Write "  </tr>"
End If
Response.Write "</table>"

'*********************************************************
' 目的：分页的页面参数保持
'          提交查询的一致性
' 输入：moveParam：分页参数
'         removeList：要移除的参数
' 返回：分页Url
'*********************************************************
Function PageUrl(moveParam,removeList)
	dim strName
	dim KeepUrl,KeepForm,KeepMove
	removeList=removeList&","&moveParam
	KeepForm=""
	For Each strName in Request.Form 
		'判断form参数中的submit、空值
		if not InstrRev(","&removeList&",",","&strName&",", -1, 1)>0 and Request.Form(strName)<>"" then
			KeepForm=KeepForm&"&"&strName&"="&Request.Form(strName)
		end if
		removeList=removeList&","&strName
	Next
	
	KeepUrl=""
	For Each strName In Request.QueryString
		If not (InstrRev(","&removeList&",",","&strName&",", -1, 1)>0) Then
			KeepUrl = KeepUrl & "&" & strName & "=" & Request.QueryString(strName)
		End If
	Next
	
	KeepMove=KeepForm&KeepUrl
	
	If (KeepMove <> "") Then 
	  KeepMove = Right(KeepMove, Len(KeepMove) - 1)
	  KeepMove = Server.HTMLEncode(KeepMove) & "&"
	End If

	PageUrl =  KeepMove & moveParam & "="
End Function

Function fPageCount(Page_Rs,showNumberLink_,nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,PageGoType,Page)

Dim This_Func_Get_Html_,toPage_,p_,sp2_,I,tpagecount
Dim NaviLength,StartPage,EndPage

This_Func_Get_Html_ = ""  : I = 1   
NaviLength=showNumberLink_ 

if IsEmpty(PageGoType) then PageGoType = 1
tpagecount=Page_Rs.pagecount
If tPageCount<1 Then tPageCount=1 

if not Page_Rs.eof or not Page_Rs.bof then

toPage_ = PageUrl("Page","submit,GetType,no-cache,_")

if Page=1 then 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""首页"">"&toF_&"</font> " &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href=javascript:ajaxPager('"&toPage_&"1') title=""首页"">"&toF_&"</a> " &vbNewLine
End If 
if Page<NaviLength then
	StartPage = 1
else
	StartPage = fix(Page / NaviLength) * NaviLength	
end if	
EndPage=StartPage+NaviLength-1 
If EndPage>tPageCount Then EndPage=tPageCount 

If StartPage>1 Then 
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href=javascript:ajaxPager('"&toPage_& Page - NaviLength &"') title=""上"&NumberLink&"页"">"&toP10_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""上"&NumberLink&"页"">"&toP10_&"</font> "  &vbNewLine
End If 

If Page <> 1 and Page <>0 Then 
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href=javascript:ajaxPager('"&toPage_&(Page-1)&"')  title=""上一页"">"&toP1_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""上一页"">"&toP1_&"</font> "  &vbNewLine
End If 

For I=StartPage To EndPage 
	If I=Page Then 
		This_Func_Get_Html_=This_Func_Get_Html_& "<b>"&I&"</b>"  &vbNewLine
	Else 
		This_Func_Get_Html_=This_Func_Get_Html_& "<a href=javascript:ajaxPager('"&toPage_&I&"')>" &I& "</a>"  &vbNewLine
	End If 
	If I<>tPageCount Then This_Func_Get_Html_=This_Func_Get_Html_& vbNewLine
Next 

If Page <> Page_Rs.PageCount and Page <>0 Then 
	This_Func_Get_Html_=This_Func_Get_Html_& " <a href=javascript:ajaxPager('"&toPage_&(Page+1)&"') title=""下一页"">"&toN1_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""下一页"">"&toN1_&"</font> "  &vbNewLine
End If 

If EndPage<tpagecount Then  
	This_Func_Get_Html_=This_Func_Get_Html_& " <a href=javascript:ajaxPager('"&toPage_& Page + NaviLength &"')  title=""下"&NumberLink&"页"">"&toN10_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& " <font color="&nonLinkColor_&"  title=""下"&NumberLink&"页"">"&toN10_&"</font> "  &vbNewLine
End If 

if Page_Rs.PageCount<>Page then  
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href=javascript:ajaxPager('"&toPage_&Page_Rs.PageCount&"') title=""尾页"">"&toL_&"</a>"  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""尾页"">"&toL_&"</font>"  &vbNewLine
End If 

If PageGoType = 1 then 
	Dim Show_Page_i
	Show_Page_i = Page + 1
	if Show_Page_i > tPageCount then Show_Page_i = 1
	This_Func_Get_Html_=This_Func_Get_Html_& "<input type=""text"" size=""4"" maxlength=""10"" name=""Func_Input_Page"" onmouseover=""this.focus();"" onfocus=""this.value='"&Show_Page_i&"';"" onKeyUp=""value=value.replace(/[^1-9]/g,'')"" onbeforepaste=""clipboardData.setData('text',clipboardData.getData('text').replace(/[^1-9]/g,''))"">" &vbNewLine _
		&"<input type=""button"" value=""Go"" onmouseover=""Func_Input_Page.focus();"" onclick=""javascript:var Js_JumpValue;Js_JumpValue=document.all.Func_Input_Page.value;if(Js_JumpValue=='' || !isNaN(Js_JumpValue)) ajaxPager('"&topage_&"'+Js_JumpValue); else ajaxPager('"&topage_&"1');"">"  &vbNewLine

Else 

	This_Func_Get_Html_=This_Func_Get_Html_& " 跳转:<select NAME=menu1 onChange=""var Js_JumpValue;Js_JumpValue=this.options[this.selectedIndex].value;if(Js_JumpValue!='') ajaxPager(Js_JumpValue);"">"
	for i=1 to tPageCount
		This_Func_Get_Html_=This_Func_Get_Html_& "<option value="&topage_&i
		if Page=i then This_Func_Get_Html_=This_Func_Get_Html_& " selected style='color:#0000FF'"
		This_Func_Get_Html_=This_Func_Get_Html_& ">第"&cstr(i)&"页</option>" &vbNewLine
	next
	This_Func_Get_Html_=This_Func_Get_Html_& "</select>" &vbNewLine

End if

This_Func_Get_Html_=This_Func_Get_Html_& p_&sp2_&" &nbsp;每页<b>"&Page_Rs.PageSize&"</b>个记录，现在是:<b><span class=""tx"">"&sp2_&Page&"</span>/"&tPageCount&"</b>页，共<b><span id='recordcount'>"&sp2_&Page_Rs.recordCount&"</span></b>个记录。"

else
	'没有记录
end if
fPageCount = This_Func_Get_Html_
End Function
%>
<script> 
cc = function(obj)
{ 
	obj.className   =   "commentListOver";
} 
bc = function(obj)
{ 
	obj.className   =   "commentListOut";
}
showComment = function(obj)
{
	var imgobj = document.getElementById("commentImg" + obj);
	var trobj = document.getElementById("commentTr" + obj);
	if(trobj.style.display == "none")
	{
		trobj.style.display = "block";
		imgobj.src="<%=strInstallDIr%>images/close.gif";
	}
	else
	{
		trobj.style.display = "none";
		imgobj.src="<%=strInstallDir%>images/open.gif";
	}
} 
</script>