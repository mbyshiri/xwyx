<!--#include file="CommonCode.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim TopicID
TopicID = PE_CLng(Request("TopicID"))
Response.Write "<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312""></head><body>"
If Request("action") = "show" Then
    Dim Star, PageTree, MaxPerPageTree
    Star = Request("Star")
    If Star = "" Or Not IsNumeric(Star) Then Star = 1
    Star = CLng(Star)
    PageTree = Star

    Response.Write "<script language=""javascript"">" & vbCrLf
    Response.Write "function showpage(PageTree,RecordCount,PageSize,PageCount)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    var arrstr='<div style=""width:100%;height:20;"">　　　&nbsp;&nbsp;共 <Strong>'+RecordCount+'</Strong> 条回复  &nbsp;页次：<Strong>'+PageTree+'/'+PageCount+'</Strong>页 &nbsp;<Strong>'+PageSize+'</Strong>条回复/页  &nbsp;分页：'" & vbCrLf
    Response.Write "    if (PageTree=='1')" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        arrstr+='<font face=webdings color=""#FF0000"">9</font>';" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        arrstr+='<a href=""ListingTree.asp?TopicID=" & TopicID & "&action=show&star=1"" title=""第一页"" target=""hiddeniframe""><font face=webdings>9</font></a>';" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    var p;" & vbCrLf
    Response.Write "    if ((PageTree-1)%10==0) " & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        p=(PageTree-1) /10" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    else" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        p=(((PageTree-1)-(PageTree-1)%10)/10)" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (p*10 > 0)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        arrstr+='<a href=""ListingTree.asp?TopicID=" & TopicID & "&action=show&star='+p*10+'"" title=""上十页"" target=""hiddeniframe"" ><font face=webdings>7</font></a> ';" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    arrstr+='<b>';" & vbCrLf
    Response.Write "    for (var i=p*10+1;i<p*10+11;i++)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        if (i==PageTree)" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            arrstr+=' <font color=""#FF0000"">'+i+'</font> ';" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        else" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            arrstr+=' <a href=""ListingTree.asp?TopicID=" & TopicID & "&action=show&star='+i+'"" target=""hiddeniframe"">'+i+'</a> ';" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        if (i==PageCount) break;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    arrstr+='</b>';" & vbCrLf
    Response.Write "    if (i<PageCount)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        arrstr+='<a href=""ListingTree.asp?TopicID=" & TopicID & "&action=show&star='+i+'"" title=""下十页"" target=""hiddeniframe""><font face=webdings>8</font></a>   ';" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if (PageTree==PageCount)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        arrstr+='<Font face=webdings color=""#FF0000"">:</font>';" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    else" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        arrstr+='<a href=""ListingTree.asp?TopicID=" & TopicID & "&action=show&star='+PageCount+'"" title=""最尾页"" target=""hiddeniframe""><font face=webdings>:</font></a>  ';" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    arrstr+='</div>';" & vbCrLf
    Response.Write "    return(arrstr)" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "var parentfollow=parent.document.getElementById(""FollowTr" & TopicID & """)" & vbCrLf
    Response.Write "var parentfollowTd=parent.document.getElementById(""FollowTd" & TopicID & """)" & vbCrLf
    Response.Write "var parentfollowImg=parent.document.getElementById(""FollowImg" & TopicID & """)" & vbCrLf
    Response.Write "if(parentfollow)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    parentfollow.style.display="""";" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "parentfollowTd.style.display="""";" & vbCrLf
    Response.Write "parentfollowImg.innerHTML='<a href=""ListingTree.asp?TopicID=" & TopicID & "&Action=hidden"" target=""hiddeniframe""  title=""关闭主题回复的列表"" ><img src=""Images/no.gif"" border=""0"" ></a>';" & vbCrLf
    Response.Write "parentfollowTd.innerHTML='<div style=""width:240px;margin-left:18px;border:1px solid black;background-color:lightyellow;color:black;padding:2px"">正在读取关于本主题的跟贴，请稍侯……</div>'" & vbCrLf
    Response.Write "</Script>" & vbCrLf
    Dim temporaryStr, i, rsGuestBook, sql, TotalPage
    temporaryStr = ""
    i = 0
    If TreeMaxPerPage = "" Or Not IsNumeric(TreeMaxPerPage) Then
        MaxPerPageTree = 5
    Else
        MaxPerPageTree = CLng(Trim(TreeMaxPerPage))
    End If
    sql = "select GuestContent,GuestTitle,GuestName,GuestDatetime,GuestContentLength from PE_GuestBook where TopicID=" & TopicID & " and GuestID<>TopicID"
    Set rsGuestBook = Server.CreateObject("adodb.recordset")
    rsGuestBook.Open sql, Conn, 1, 1
    If Star > 1 Then
        rsGuestBook.Move (Star - 1) * MaxPerPageTree
    End If
    Do While Not rsGuestBook.EOF
        temporaryStr = temporaryStr & "<div style='width:100%;height:20'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src='Images/no.gif'><a href='Guest_Reply.asp?TopicID=" & TopicID & "'>" & rsGuestBook("GuestTitle") & "</a><I><font color=gray>(" & rsGuestBook("GuestContentLength") & "字) － " & rsGuestBook("GuestName") & "，" & FormatDateTime(rsGuestBook("GuestDatetime"), 0) & "</font></I></div>"
        i = i + 1
        If i >= MaxPerPageTree Then Exit Do
        rsGuestBook.MoveNext
    Loop
    If rsGuestBook.RecordCount Mod MaxPerPageTree = 0 Then
        TotalPage = rsGuestBook.RecordCount \ MaxPerPageTree
    Else
        TotalPage = rsGuestBook.RecordCount \ MaxPerPageTree + 1
    End If
    temporaryStr = Replace(Replace(Replace(Replace(Replace(Replace(temporaryStr, "\", "\\"), "'", "\'"), vbCrLf, ""), Chr(13), ""), "<BR>", ""), "</P><P>", "")
    Response.Write "<Script Language=JavaScript>"
    Response.Write "var arrstr='';" & vbCrLf
    Response.Write "arrstr='" & temporaryStr & "';" & vbCrLf
    Response.Write "arrstr+=showpage(" & PageTree & "," & rsGuestBook.RecordCount & "," & MaxPerPageTree & "," & TotalPage & ");" & vbCrLf
    Response.Write "parent.document.getElementById(""FollowTd" & TopicID & """).innerHTML=arrstr;" & vbCrLf
    Response.Write "</Script>" & vbCrLf
    Set rsGuestBook = Nothing
    Call CloseConn
Else
    Response.Write "<script language=""javascript"">" & vbCrLf
    Response.Write "var parentfollow=parent.document.getElementById(""follow" & TopicID & """)" & vbCrLf
    Response.Write "var parentfollowTd=parent.document.getElementById(""followTd" & TopicID & """)" & vbCrLf
    Response.Write "var parentfollowImg=parent.document.getElementById(""followImg" & TopicID & """)" & vbCrLf
    Response.Write "if(parentfollow)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    parentfollow.style.display=""none"";    " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "parentfollowTd.style.display=""none"";" & vbCrLf
    Response.Write "parentfollowImg.innerHTML='<a href=""ListingTree.asp?TopicID=" & TopicID & "&Action=show"" target=""hiddeniframe""  title=""打开回复的主题列表"" ><img src=""Images/yes.gif"" border=""0"" ></a>';" & vbCrLf
    Response.Write "</script>" & vbCrLf
End If

Response.Write "</body></html>"
Call CloseConn
%>
