<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

'评论所用变量
Dim InfoID, rsCommentUser, sql, sqlout, InfoTitle, CommentNum, ShowVote
Dim InfoUrl


If ChannelID > 0 Then
    Call GetChannel(ChannelID)
Else
    Response.Write "频道参数丢失！"
    FoundErr = True
    Response.End
End If

UserLogined = CheckUserLogined()
If UserLogined = True Then GetUser (UserName)

CommentNum = Trim(Request("CommentNum"))
ShowVote = PE_CLng(Trim(Request("ShowVote")))
If Action = "ShowAll" Then
    PageTitle = XmlText("Site", "Comment/ShowDetal", "显示评论详细内容")
Else
    PageTitle = XmlText("Site", "Comment/Send", "发表评论")
End If

Select Case ModuleType
Case 1
    InfoID = Trim(Request("ArticleID"))
    strFileName = "Comment.asp?ArticleID=" & InfoID & "&Action=ShowAll"
    sql = "select A.Title,A.UpdateTime,A.InfoPurview,A.InfoPoint,C.EnableComment,C.CheckComment,C.ClassID,C.ParentDir,C.ClassDir,C.ClassPurview from PE_Article A left join PE_Class C on A.ClassID=C.ClassID where A.ArticleID="
    sqlout = "select ClassID, ChannelID from PE_Article where ArticleID="
Case 2
    InfoID = Trim(Request("SoftID"))
    strFileName = "Comment.asp?SoftID=" & InfoID & "&Action=ShowAll"
    sql = "select S.SoftName,S.UpdateTime,C.EnableComment,C.CheckComment,C.ClassID,C.ParentDir,C.ClassDir from PE_Soft S inner join PE_Class C on S.ClassID=C.ClassID where S.SoftID="
    sqlout = "select ClassID,ChannelID from PE_Soft where SoftID="
Case 3
    InfoID = Trim(Request("PhotoID"))
    strFileName = "Comment.asp?PhotoID=" & InfoID & "&Action=ShowAll"
    sql = "select P.PhotoName,P.UpdateTime,P.InfoPurview,P.InfoPoint,C.EnableComment,C.CheckComment,C.ClassID,C.ParentDir,C.ClassDir,C.ClassPurview from PE_Photo P inner join PE_Class C on P.ClassID=C.ClassID where P.PhotoID="
    sqlout = "select ClassID,ChannelID from PE_Photo where PhotoID="
Case 5
    InfoID = Trim(Request("ProductID"))
    strFileName = "Comment.asp?ProductID=" & InfoID & "&Action=ShowAll"
    sql = "select P.ProductName,P.UpdateTime,C.EnableComment,C.CheckComment,C.ClassID,C.ParentDir,C.ClassDir from PE_Product P left join PE_Class C on P.ClassID=C.ClassID where P.ProductID="
    sqlout = "select ClassID,ChannelID from PE_Product where ProductID="
Case 6
    InfoID = Trim(Request("SupplyID"))
    strFileName = "Comment.asp?SupplyID=" & InfoID & "&Action=ShowAll"
    sql = "select P.SupplyTitle,C.EnableComment,C.CheckComment,C.ClassID from PE_Supply P left join PE_Class C on P.ClassID=C.ClassID where P.SupplyId="
    sqlout = "select ClassID,ChannelID from PE_Supply where SupplyId="
End Select
If InfoID = ""  and Action<>"UpdateVote" Then
    FoundErr = True
    ErrMsg = ErrMsg & Replace(XmlText("Site", "Comment/Err1", "<li>请指定{$ChannelShortName}ID</li>"), "{$ChannelShortName}", ChannelShortName)
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Call CloseConn
    Response.End
Else
    InfoID = PE_CLng(InfoID)
End If

sql = sql & InfoID
sqlout = sqlout & InfoID
If CommentNum = "" Then
    CommentNum = 10
Else
    CommentNum = PE_CLng(CommentNum)
End If

Select Case Action
Case "JS"
    Call GetCommentJS(CommentNum)
Case "Save"
    Call SaveComment
Case "UpdateVote"
    Call UpdateVote 	
Case Else
    Call ShowComment
End Select

If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If


Sub ShowComment()
    Dim CommentedID, arrCommentedID, i, trs
    Set trs = Conn.Execute(sqlout)
    If trs.BOF And trs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & Replace(XmlText("Site", "Comment/Err2", "<br>找不到指定的{$ChannelShortName}"), "{$ChannelShortName}", ChannelShortName)
        Exit Sub
    End If
    'If trs(0) = -1 Then
      '  FoundErr = True
        'ErrMsg = ErrMsg & Replace(XmlText("Site", "Comment/Err5", "<li>对不起，未指定栏目的{$ChannelShortName}暂不开放发表评论！</li>"), "{$ChannelShortName}", ChannelShortName)
       ' Exit Sub
    'End If
    Set trs = nothing
    Set trs = Conn.Execute(sql)
    If trs.BOF And trs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & Replace(XmlText("Site", "Comment/Err2", "<br>找不到指定的{$ChannelShortName}"), "{$ChannelShortName}", ChannelShortName)
        Exit Sub
    End If
    InfoTitle = trs(0)
    ChannelUrl = UrlPrefix(SiteUrlType, ChannelUrl) & ChannelUrl
    ChannelUrl_ASPFile = UrlPrefix(SiteUrlType, ChannelUrl_ASPFile) & ChannelUrl_ASPFile
    Select Case ModuleType
    Case 1
        InfoUrl = GetInfoUrl(trs("ParentDir"), trs("ClassDir"), trs("UpdateTime"), InfoID, trs("ClassPurview"), trs("InfoPurview"), trs("InfoPoint"))
    Case 2
        InfoUrl = GetInfoUrl(trs("ParentDir"), trs("ClassDir"), trs("UpdateTime"), InfoID, "", "", "")
    Case 3
        InfoUrl = GetInfoUrl(trs("ParentDir"), trs("ClassDir"), trs("UpdateTime"), InfoID, trs("ClassPurview"), trs("InfoPurview"), trs("InfoPoint"))
    Case 5
        InfoUrl = GetInfoUrl(trs("ParentDir"), trs("ClassDir"), trs("UpdateTime"), InfoID, "", "", "")
    Case 6
        InfoUrl = GetInfoUrl("", "", "", InfoID, "", "", "")
    End Select

    Set trs = Nothing

    Dim sqlComment, rsComment
    
    CurrentPage = Trim(Request("page"))
    If CurrentPage = "" Then
        CurrentPage = 1
    Else
        CurrentPage = PE_CLng(CurrentPage)
    End If
    SkinID = DefaultSkinID
    strHtml = GetTemplate(ChannelID, 16, 0)
    strHtml = Replace(strHtml, "{$ArticleID}", InfoID)
    strHtml = Replace(strHtml, "{$SupplyID}", InfoID)
    strHtml = Replace(strHtml, "{$SoftID}", InfoID)
    strHtml = Replace(strHtml, "{$PhotoID}", InfoID)
    strHtml = Replace(strHtml, "{$ProductID}", InfoID)
    Call ReplaceCommonLabel

    strHtml = PE_Replace(strHtml, "{$Meta_Keywords_Channel}", Meta_Keywords_Channel)
    strHtml = PE_Replace(strHtml, "{$Meta_Description_Channel}", Meta_Description_Channel)
    strHtml = PE_Replace(strHtml, "{$ChannelID}", ChannelID)
    strHtml = PE_Replace(strHtml, "{$ChannelDir}", ChannelDir)
    strHtml = PE_Replace(strHtml, "{$ChannelName}", ChannelName)
    strHtml = PE_Replace(strHtml, "{$ChannelShortName}", ChannelShortName)
    strHtml = PE_Replace(strHtml, "{$UploadDir}", UploadDir)
    strHtml = PE_Replace(strHtml, "{$MenuJS}", GetMenuJS(ChannelDir, False))
    strHtml = PE_Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(SkinID))

    strNavPath = XmlText("BaseText", "Nav", "您现在的位置：") & "&nbsp;<a class='LinkPath' href='" & SiteUrl & "'>" & SiteName & "</a>"
    If Trim(ChannelName) <> "" And ShowNameOnPath <> False Then
        strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;<a class='LinkPath' href='"
        If UseCreateHTML > 0 Then
            strNavPath = strNavPath & ChannelUrl & "/Index" & FileExt_Index
        Else
            strNavPath = strNavPath & ChannelUrl & "/Index.asp"
        End If
        strNavPath = strNavPath & "'>" & ChannelName & "</a>"
    End If
    strNavPath = strNavPath & "&nbsp;" & strNavLink & "&nbsp;" & XmlText("Site", "Comment/Send", "发表评论")

    strHtml = Replace(strHtml, "{$PageTitle}", PageTitle)
    strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)
    strHtml = Replace(strHtml, "{$ShowJS_Comment}", ShowJS_Comment(UserLogined))
    
    strHtml = Replace(strHtml, "{$ArticleTitle}", InfoTitle)
    strHtml = Replace(strHtml, "{$SupplyTitle}", InfoTitle)
    strHtml = Replace(strHtml, "{$SoftTitle}", InfoTitle)
    strHtml = Replace(strHtml, "{$PhotoTitle}", InfoTitle)
    strHtml = Replace(strHtml, "{$ProductTitle}", InfoTitle)
    strHtml = Replace(strHtml, "{$InfoUrl}", InfoUrl)
    Dim CommentIsShow
    regEx.Pattern = "【CommentIsShow】([\s\S]*?)【\/CommentIsShow】"
    Set Matches = regEx.Execute(strHtml)
    For Each Match In Matches
        CommentIsShow = Match.value
    Next
    If UserLogined = True Then
         strHtml = Replace(strHtml, CommentIsShow, "")
    End If
    strHtml = Replace(strHtml, "{$UserName}", UserName)
    strHtml = Replace(strHtml, "{$UserEmail}", email)
    strHtml = Replace(strHtml, "【CommentIsShow】", "")
    strHtml = Replace(strHtml, "【/CommentIsShow】", "")
    
    Dim strCommentList, arrTemp
    If Action = "ShowAll" Then
        regEx.Pattern = "\{\$ShowCommentList\((.*?)\)\}"
        Set Matches = regEx.Execute(strHtml)
		
        For Each Match In Matches
        arrTemp = Split(Match.SubMatches(0), ",")
        If UBound(arrTemp) = 0 Then				
            strCommentList = ShowCommentList(PE_CLng(Match.SubMatches(0)),0,1,False,True)
        ElseIf  UBound(arrTemp) = 4 Then		
            strCommentList = ShowCommentList(PE_CLng(arrTemp(0)),PE_CLng(arrTemp(1)),PE_CLng(arrTemp(2)),PE_CBool(arrTemp(3)),PE_CBool(arrTemp(4)))		
        End If		
            strHtml = Replace(strHtml, Match.value, strCommentList)
        Next
    Else
        regEx.Pattern = "\{\$ShowCommentList\((.*?)\)\}"
        Set Matches = regEx.Execute(strHtml)
        For Each Match In Matches
            strHtml = Replace(strHtml, Match.value, "")
        Next
    End If
    strHtml = Replace(strHtml, "value= ", "value='' ")
    strHtml = Replace(strHtml, "Value= ", "value='' ")
    Response.Write strHtml
End Sub
'=================================================
'过程名：ShowCommentList()
'作  用：显示评论列表
'参  数：CommentShowType  1。表格显示
'                         2。分条显示
'                         3。DIV输出
'参  数：CommentNum       评论数量,0为不限制
'参  数：OrderType        排序方式,1为按照评论ID排序,2为按照热门评论排列
'参  数：ShowVote          是否显示投票的支持，反对选项
'参  数：UsePage          是否显示分页

'=================================================
Function ShowCommentList(CommentShowType, CommentNum, OrderType, ShowVote, UsePage)
    Dim rsComment, sqlComment, iCount, strHTM, strUserName
    ShowVote = PE_CBool(ShowVote)	
    If PE_CLng(CommentNum) = 0 Then 
        sqlComment = "select * "
    Else
        sqlComment = "select top "& PE_CLng(CommentNum) &" * "               
    End If
    sqlComment = sqlComment &" from PE_Comment where ModuleType=" & ModuleType & " and InfoID=" & InfoID & " and Passed=" & PE_True & " order by "
    If PE_CLng(OrderType) = 0 Or OrderType = 1 Then 
         sqlComment = sqlComment &" CommentID desc "
    Else
         sqlComment = sqlComment &" Support desc "             
    End If    	
    Set rsComment = Server.CreateObject("ADODB.Recordset")
    rsComment.Open sqlComment, Conn, 1, 1
    If rsComment.BOF And rsComment.EOF Then
        strHTM = strHTM & XmlText("Site", "Comment/Err3", "&nbsp;&nbsp;&nbsp;&nbsp;没有任何评论")
    Else
        totalPut = rsComment.RecordCount
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
                rsComment.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        Dim strCommentList1, strCommentList2, strCommentReply1, strCommentReply2
        If ShowVote = True Then
            strCommentList1 = "<tr class='Comment_tdbg1'><td width='100'>{$CommentName}</td><td style='width:480; word-wrap:break-word;'>{$Content}</td><td align='center' width='120'>{$WriteTime}</td><td align='center' width='40'>{$Score}分</td><td  width='150'><span id=Support{$CommentID} Name=Support{$CommentID}><a href=""javascript:Support({$CommentID});"">支持</a></span>[<span id=SupportCount{$CommentID} Name=SupportCount{$CommentID}>{$SupportCount}</span>] <span id=Opposed{$CommentID} Name=Opposed{$CommentID}><a href=""javascript:Opposed({$CommentID});"">反对</a></span>[<span id=OpposedCount{$CommentID} Name=OpposedCount{$CommentID}>{$OpposedCount}</span>]</span></a></td></tr>"
            strCommentList2 ="<tr class='Comment_title'><td height=22 colspan='3'>&nbsp;&nbsp评论人：{$CommentName}&nbsp;&nbsp;评论时间：{$WriteTime}&nbsp;&nbsp;打分：{$Score}分</td></tr><tr class='Comment_tdbg1'><td>{$Content}</td><td width=5>&nbsp;</td></tr><tr><td colspan='3' ><span id=Support{$CommentID} Name=Support{$CommentID}><a href=""javascript:Support({$CommentID});"">支持</a></span>[<span id=SupportCount{$CommentID} Name=SupportCount{$CommentID}>{$SupportCount}</span>] <span id=Opposed{$CommentID} Name=Opposed{$CommentID}><a href=""javascript:Opposed({$CommentID});"">反对</a></span>[<span id=OpposedCount{$CommentID} Name=OpposedCount{$CommentID}>{$OpposedCount}</span>]</span></a></td></tr>"
        Else
            strCommentList1 = XmlText("Site", "Comment/ShowComment1", "<tr class='Comment_tdbg1'><td width='100'>{$CommentName}</td><td style='width:480; word-wrap:break-word;'>{$Content}</td><td align='center' width='120'>{$WriteTime}</td><td align='center' width='40'>{$Score}分</td></tr>")
            strCommentList2 = XmlText("Site", "Comment/ShowComment2", "<tr class='Comment_title'><td height=22 colspan='3'>&nbsp;&nbsp评论人：{$CommentName}&nbsp;&nbsp;评论时间：{$WriteTime}&nbsp;&nbsp;打分：{$Score}分</td></tr><tr class='Comment_tdbg1'><td width=5>&nbsp;</td><td  class=>{$Content}</td><td width=5 class=>&nbsp;</td></tr>")
        End If		
        strCommentReply1 = XmlText("Site", "Comment/AdminReplayType1", "<tr class='Comment_tdbg2'><td>&nbsp;</td><td colspan='5'>★&nbsp;管理员『{$AdminName}』于{$ReplyTime}回复道：&nbsp;&nbsp;&nbsp;&nbsp;{$ReplyContent}</td></tr>")
        strCommentReply2 = XmlText("Site", "Comment/AdminReplayType2", "<tr class='Comment_tdbg2'><td width=5>&nbsp;</td><td>★&nbsp;管理员『{$AdminName}』于 {$ReplyTime}回复道：&nbsp;&nbsp;{$ReplyContent}</td><td width=5 class=>&nbsp;</td></tr>")

		
        Select Case CommentShowType
        Case 1
            strHTM = "<table width='100%' align='center' border='0' cellspacing='1' cellpadding='2' class='Comment_border'>"
            If ShowVote = True Then			
                strHTM = strHTM & "<tr class='Comment_title' align='center'><td>评论人</td><td>评论内容</td><td>评论时间</td><td>打分</td><td>投票</td></tr>"
            Else
                strHTM = strHTM & XmlText("Site", "Comment/ShowCommentList", "<tr class='Comment_title' align='center'><td>评论人</td><td>评论内容</td><td>评论时间</td><td>打分</td></tr>")
            End If							
            Do While Not rsComment.EOF
                If rsComment("UserType") = 1 Then
                    strUserName = "【<a href='" & strInstallDir & "ShowUser.asp?UserName=" & rsComment("UserName") & "'>" & rsComment("UserName") & "</a>】"
                Else
                    strUserName = "【<span title='类别：游客" & vbCrLf & "姓名：" & rsComment("UserName") & vbCrLf & "信箱：" & rsComment("Email") & vbCrLf & "Oicq：" & rsComment("Oicq") & vbCrLf & "主页：" & rsComment("Homepage") & "' style='cursor:hand'>" & rsComment("UserName") & "</span>】"
                End If
                strHTM = strHTM & Replace(Replace(Replace(Replace(Replace(Replace(Replace(strCommentList1, "{$CommentName}", strUserName), "{$Content}", ReplaceText(rsComment("Content"), 3)), "{$WriteTime}", rsComment("WriteTime")), "{$Score}", rsComment("Score")),"{$CommentID}",rsComment("CommentID")),"{$SupportCount}",PE_CLng(rsComment("Support"))),"{$OpposedCount}",PE_CLng(rsComment("Opposed")))
                If rsComment("ReplyContent") <> "" Then
                    strHTM = strHTM & Replace(Replace(Replace(strCommentReply1, "{$AdminName}", rsComment("ReplyName")), "{$ReplyTime}", rsComment("ReplyTime")), "{$ReplyContent}", rsComment("ReplyContent"))
                End If
                rsComment.MoveNext
                iCount = iCount + 1
                If iCount >= MaxPerPage Then Exit Do
            Loop
            strHTM = strHTM & "</table><br>"
        Case 2
            Do While Not rsComment.EOF
                If rsComment("UserType") = 1 Then
                    strUserName = "【<a href='" & strInstallDir & "ShowUser.asp?UserName=" & rsComment("UserName") & "'>" & rsComment("UserName") & "</a>】"
                Else
                    strUserName = "【<span title='类别：游客" & vbCrLf & "姓名：" & rsComment("UserName") & vbCrLf & "信箱：" & rsComment("Email") & vbCrLf & "Oicq：" & rsComment("Oicq") & vbCrLf & "主页：" & rsComment("Homepage") & "' style='cursor:hand'>" & rsComment("UserName") & "</span>】"
                End If
                
                strHTM = strHTM & "     <table width='100%' align='center' border='0' cellspacing='1' cellpadding='2' class='Comment_border'>" & vbCrLf
                
                strHTM = strHTM & Replace(Replace(Replace(Replace(Replace(Replace(Replace(strCommentList2, "{$CommentName}", strUserName), "{$Content}", ReplaceText(rsComment("Content"), 3)), "{$WriteTime}", rsComment("WriteTime")), "{$Score}", rsComment("Score")),"{$CommentID}",rsComment("CommentID")),"{$SupportCount}",PE_CLng(rsComment("Support"))),"{$OpposedCount}",PE_CLng(rsComment("Opposed")))        
                
                If rsComment("ReplyContent") <> "" Then
                    strHTM = strHTM & Replace(Replace(Replace(strCommentReply2, "{$AdminName}", rsComment("ReplyName")), "{$ReplyTime}", rsComment("ReplyTime")), "{$ReplyContent}", rsComment("ReplyContent"))
                End If

                strHTM = strHTM & "     </table>" & vbCrLf
                strHTM = strHTM & "      <table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
                strHTM = strHTM & "        <tr>" & vbCrLf
                strHTM = strHTM & "          <td class='main_shadow'>" & vbCrLf
                strHTM = strHTM & "          </td>" & vbCrLf
                strHTM = strHTM & "        </tr>" & vbCrLf
                strHTM = strHTM & "      </table>" & vbCrLf
                rsComment.MoveNext
                iCount = iCount + 1
                If iCount >= MaxPerPage Then Exit Do
            Loop
            strHTM = strHTM & "<br>"
        Case 3
            Do While Not rsComment.EOF
                If rsComment("UserType") = 1 Then
                    strUserName = "【<a href='" & strInstallDir & "ShowUser.asp?UserName=" & rsComment("UserName") & "'>" & rsComment("UserName") & "</a>】"
                Else
                    strUserName = "【<span title=""类别：游客" & vbCrLf & "姓名：" & rsComment("UserName") & vbCrLf & "信箱：" & rsComment("Email") & vbCrLf & "Oicq：" & rsComment("Oicq") & vbCrLf & "主页：" & rsComment("Homepage") & """ style=""cursor:hand"">" & rsComment("UserName") & "</span>】"
                End If
                
                strHTM = strHTM & ("<div class=""comment_body"">" & vbCrLf)
                strHTM = strHTM & ("<div class=""comment_user"">评论人：" & strUserName & "</div>" & vbCrLf)
                strHTM = strHTM & ("<div class=""comment_time"">评论时间：" & rsComment("WriteTime") & "</div>" & vbCrLf)
                strHTM = strHTM & ("<div class=""comment_score"">打分：" & rsComment("Score") & "分</div>" & vbCrLf)
                strHTM = strHTM & ("<div class=""comment_content"">" & ReplaceText(rsComment("Content"), 3) & "</div>" & vbCrLf)
                If ShowVote = True Then			
                    strHTM = strHTM & ("<div class=""comment_vote""><span id=Support{$CommentID} Name=Support{$CommentID}><a href=""javascript:Support({$CommentID});"">支持</a></span>[<span id=SupportCount{$CommentID}  Name=SupportCount{$CommentID}>{$SupportCount}</span>] <span id=Opposed{$CommentID}  Name=Opposed{$CommentID}><a href=""javascript:Opposed({$CommentID});"">反对</a></span>[<span id=OpposedCount{$CommentID}  Name=OpposedCount{$CommentID}>{$OpposedCount}</span>]</span></a></div>" & vbCrLf)
                    strHTM = Replace(Replace(Replace(strHTM ,"{$CommentID}",rsComment("CommentID")),"{$SupportCount}",PE_CLng(rsComment("Support"))),"{$OpposedCount}",PE_CLng(rsComment("Opposed"))) 					
                End If								
                If rsComment("ReplyContent") <> "" Then
                    strHTM = strHTM & ("<div class=""comment_adminreply"">★&nbsp;管理员『" & rsComment("ReplyName") & "』于" & rsComment("ReplyTime") & "回复道：&nbsp;&nbsp;&nbsp;&nbsp;" & rsComment("ReplyContent") & "</div>" & vbCrLf)
                End If
                strHTM = strHTM & "</div>" & vbCrLf
                rsComment.MoveNext
                iCount = iCount + 1
                If iCount >= MaxPerPage Then Exit Do
            Loop
        End Select
    End If
    rsComment.Close
    Set rsComment = Nothing
    If PE_CBool(UsePage) = True Then
        If XmlText("Site", "Comment/ShowPageType", "Chinese") = "English" Then
            strHTM = strHTM & ShowPage_en(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("Site", "Comment/ShowPageChar", "Comment"), False)
        Else
            strHTM = strHTM & ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("Site", "Comment/ShowPageChar", "条评论"), False)
        End If
        If CommentShowType < 3 Then
           strHTM = strHTM & ("</td></tr><tr height='40'><td align='center' colspan='2'>【<a href='" & InfoUrl & "'>返回" & ChannelShortName & "内容页</a>】</td></tr>")
        Else
           strHTM = strHTM & ("<div class=""comment_backurl"">【<a href='" & InfoUrl & "'>返回" & ChannelShortName & "内容页</a>】")
    End If
    End If
    If ShowVote = True Then
        strHTM = strHTM & ShowVoteJS_Comment
    End If			
    ShowCommentList = strHTM
End Function
'=================================================
'过程名：ShowJS_Comment()
'作  用：评论输入判断,快捷键提交评论相关js
'参  数：无
'=================================================
Function ShowJS_Comment(IsLogin)
    Dim strJS
    strJS = "<script language='JavaScript' type='text/JavaScript'>" & vbCrLf
    strJS = strJS & "function Check(){" & vbCrLf
    If IsLogin = False Then
        strJS = strJS & "  if (document.form1.Name.value==''){" & vbCrLf
        strJS = strJS & "    alert('请输入姓名！');" & vbCrLf
        strJS = strJS & "    document.form1.Name.focus();" & vbCrLf
        strJS = strJS & "    return false;" & vbCrLf
        strJS = strJS & "  }" & vbCrLf
    End If
    strJS = strJS & "  if (document.form1.Content.value==''){" & vbCrLf
    strJS = strJS & "    alert('请输入内容！');" & vbCrLf
    strJS = strJS & "    document.form1.Content.focus();" & vbCrLf
    strJS = strJS & "    return false;" & vbCrLf
    strJS = strJS & "  }" & vbCrLf
    strJS = strJS & "  return true;  " & vbCrLf
    strJS = strJS & "}" & vbCrLf
    strJS = strJS & "function KeyDown(obj){" & vbCrLf
    strJS = strJS & "  var isIE = navigator.userAgent.indexOf(""MSIE"")>0;" & vbCrLf
    strJS = strJS & "  obj.onkeydown = function(e){" & vbCrLf
    strJS = strJS & "    if(isIE){" & vbCrLf
    strJS = strJS & "      if(event.ctrlKey && event.keyCode == 13)" & vbCrLf
    strJS = strJS & "      {" & vbCrLf
    strJS = strJS & "         document.form1.submit();" & vbCrLf
    strJS = strJS & "       }" & vbCrLf
    strJS = strJS & "    }" & vbCrLf
    strJS = strJS & "    else{" & vbCrLf
    strJS = strJS & "      if(e.ctrlKey && e.which == 13)" & vbCrLf
    strJS = strJS & "      {" & vbCrLf
    strJS = strJS & "       document.form1.submit();" & vbCrLf
    strJS = strJS & "      }" & vbCrLf
    strJS = strJS & "    }" & vbCrLf
    strJS = strJS & "  } " & vbCrLf
    strJS = strJS & "}	" & vbCrLf	
    strJS = strJS & "</script>" & vbCrLf
    ShowJS_Comment = strJS
End Function



'=================================================
'函数名：SaveComment()
'作  用：保存评论
'参  数：无
'=================================================
Sub SaveComment()

    Dim trs, NeedCheck
    Set trs = Conn.Execute(sql)
    If trs.BOF And trs.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & Replace(XmlText("Site", "Comment/Err2", "<br>找不到指定的{$ChannelShortName}"), "{$ChannelShortName}", ChannelShortName)
        Exit Sub
    End If
    'InfoTitle = trs(0)
    EnableComment = trs("EnableComment") Or UserEnableComment 
    NeedCheck = trs("CheckComment") And (Not UserCheckComment)
    ClassID = trs(6)
    If ClassID = "" Or IsNull(ClassID) Then
        Dim rstemp,rsTempClass
        Set rstemp = Conn.Execute(sqlout)
        If rstemp(0) = -1 Then
            Set rsTempClass = Conn.Execute("Select * from PE_Channel Where ChannelID="&rstemp(1))
            EnableComment = PE_Cbool(rsTempClass("EnableComment")) Or UserEnableComment
            NeedCheck = PE_Cbool(rsTempClass("CheckComment")) And (Not UserCheckComment)		
        Else
            EnableComment = False
            NeedCheck = True		
        End If
    End If
    Set trs = Nothing
    If EnableComment <> True Then
        FoundErr = True
        ErrMsg = ErrMsg & XmlText("Site", "Comment/Err4", "<br><li>对不起，您没有本栏目发表评论的权限！</li>")
        Exit Sub
    End If

    Dim rsComment, tClass
    Dim CommentUserType, CommentUserName, CommentUserSex, CommentUserEmail, CommentUserOicq
    Dim CommentUserIcq, CommentUserMsn, CommentUserHomepage, CommentUserScore, CommentUserContent
    If UserLogined = False Then
        CommentUserType = 0
        CommentUserName = PE_HTMLEncode(Trim(Request("Name")))
        If CommentUserName = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>输入姓名有误！</li>"
            Exit Sub
        End If
        CommentUserSex = PE_HTMLEncode(Trim(Request("Sex")))
        CommentUserOicq = PE_HTMLEncode(Trim(Request("Oicq")))
        CommentUserIcq = PE_HTMLEncode(Trim(Request("Icq")))
        CommentUserMsn = PE_HTMLEncode(Trim(Request("Msn")))
        CommentUserEmail = PE_HTMLEncode(Trim(Request("Email")))
        CommentUserHomepage = ReplaceUrlBadChar(Trim(Request("Homepage")))
        If CommentUserHomepage = "http://" Or IsNull(CommentUserHomepage) Then CommentUserHomepage = ""
    Else
        CommentUserType = 1
        CommentUserName = UserName
    End If

    CommentUserScore = PE_CLng(Request.Form("Score"))
    CommentUserContent = Trim(Request.Form("Content"))
    If CommentUserContent = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请输入内容</li>"
        Exit Sub
    End If

    'CommentUserContent = ReplaceText(ReplaceBadChar(FilterJS(CommentUserContent)), 3)
    CommentUserContent = PE_HTMLEncode(CommentUserContent)
    Set rsComment = Server.CreateObject("adodb.recordset")
    sql = "select top 1 * from PE_Comment"
    rsComment.Open sql, Conn, 1, 3
    rsComment.addnew
    rsComment("ModuleType") = ModuleType
    rsComment("InfoID") = InfoID
    rsComment("UserType") = CommentUserType
    rsComment("UserName") = CommentUserName
    rsComment("Sex") = CommentUserSex
    rsComment("Oicq") = CommentUserOicq
    rsComment("Icq") = CommentUserIcq
    rsComment("Msn") = CommentUserMsn
    rsComment("Email") = CommentUserEmail
    rsComment("Homepage") = CommentUserHomepage
    rsComment("IP") = UserTrueIP
    rsComment("Score") = CommentUserScore
    rsComment("Content") = ReplaceBadUrl(CommentUserContent) '过滤非法系统URL
    rsComment("WriteTime") = Now()
    rsComment("Passed") = PE_CBool((Not NeedCheck))
	rsComment.Update
    rsComment.Close
    Set rsComment = Nothing
    Conn.Execute ("update PE_Channel set CommentCount=CommentCount+1 where ChannelID=" & ChannelID & "")
    Select Case ModuleType
    Case 1
        Conn.Execute ("update PE_Article set CommentCount=CommentCount+1 where ArticleID=" & InfoID & "")
    Case 2
        Conn.Execute ("update PE_Soft set CommentCount=CommentCount+1 where SoftID=" & InfoID & "")
    Case 3
        Conn.Execute ("update PE_Photo set CommentCount=CommentCount+1 where PhotoID=" & InfoID & "")
    Case 5
        Conn.Execute ("update PE_Product set CommentCount=CommentCount+1 where ProductID=" & InfoID & "")
    End Select
    If NeedCheck = False Then
       ' Response.Redirect ComeUrl
        Call WriteSuccessMsg(XmlText("Site", "Comment/SusMsg1", "发表评论成功！"), ComeUrl)
    Else
        Call WriteSuccessMsg(XmlText("Site", "Comment/SusMsg2", "发表评论成功！请等候管理员的审核！审核后才会显示"), ComeUrl)
    End If
'	Response.write "<meta http-equiv=""refresh"" content=""1; url="& ComeUrl &""" />"
End Sub


'=================================================
'过程名：ShowVoteJS_Comment()
'作  用：评论投票相关js
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
    strJS = strJS & "                XMLHttp = false;" & vbCrLf    			    
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
    strJS = strJS & "        var content = document.getElementsByName(""Support""+id);" & vbCrLf      
    strJS = strJS & "        var Support = document.getElementsByName(""SupportCount""+id);" & vbCrLf					
    strJS = strJS & "        _xmlhttp.open('GET',url,true);" & vbCrLf
    strJS = strJS & "        _xmlhttp.onreadystatechange=function()" & vbCrLf
    strJS = strJS & "        {" & vbCrLf
    strJS = strJS & "            if(_xmlhttp.readyState == 4)" & vbCrLf
    strJS = strJS & "            {" & vbCrLf
    strJS = strJS & "                if(_xmlhttp.status == 200)" & vbCrLf      
    strJS = strJS & "               {" & vbCrLf
    strJS = strJS & "                   var ResponseText = unescape(_xmlhttp.responseText);	" & vbCrLf
    strJS = strJS & "                    for(var i=0;i<Support.length;i++){	" & vbCrLf				
    strJS = strJS & "                      Support[i].innerHTML=ResponseText;	" & vbCrLf
    strJS = strJS & "                    }" & vbCrLf		
    strJS = strJS & "                    for(i=0;i<content.length;i++){	" & vbCrLf								
    strJS = strJS & "                      content[i].innerHTML='已支持';" & vbCrLf
    strJS = strJS & "                    }	" & vbCrLf			
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
    strJS = strJS & "        var content = document.getElementsByName(""Opposed""+id);  " & vbCrLf    
    strJS = strJS & "        var Opposed = document.getElementsByName(""OpposedCount""+id);" & vbCrLf					
    strJS = strJS & "        _xmlhttp.open('GET',url,true);" & vbCrLf
    strJS = strJS & "        _xmlhttp.onreadystatechange=function()" & vbCrLf
    strJS = strJS & "        {" & vbCrLf
    strJS = strJS & "            if(_xmlhttp.readyState == 4)" & vbCrLf  
    strJS = strJS & "            {" & vbCrLf
    strJS = strJS & "                if(_xmlhttp.status == 200)     " & vbCrLf 
    strJS = strJS & "                {" & vbCrLf
    strJS = strJS & "                    var ResponseText = unescape(_xmlhttp.responseText);	" & vbCrLf
    strJS = strJS & "                    for(var i=0;i<Opposed.length;i++){	" & vbCrLf	
    strJS = strJS & "                      Opposed[i].innerHTML=ResponseText;	" & vbCrLf	
    strJS = strJS & "                    }" & vbCrLf			
    strJS = strJS & "                    for(i=0;i<content.length;i++){	" & vbCrLf			
    strJS = strJS & "                     content[i].innerHTML='已反对';" & vbCrLf
    strJS = strJS & "                    }" & vbCrLf	
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


'=================================================
'过程名：GetCommentJS
'作  用：显示相关评论的JS代码
'参  数：CommentNum  ----最多显示多少个评论
'=================================================
Sub GetCommentJS(CommentNum)
    Dim rsComment, sqlComment, strComment, strUserName
    If CommentNum > 0 And CommentNum <= 100 Then
        sqlComment = "select top " & CommentNum
    Else
        sqlComment = "select top 10 "
    End If
    sqlComment = sqlComment & " * from PE_Comment where ModuleType=" & ModuleType & " and InfoID=" & InfoID & " and Passed=" & PE_True & " order by CommentID desc"
    
    Set rsComment = Server.CreateObject("ADODB.Recordset")
    rsComment.Open sqlComment, Conn, 1, 1
    If rsComment.BOF And rsComment.EOF Then
        strComment = XmlText("Site", "Comment/Err3", "&nbsp;&nbsp;&nbsp;&nbsp;没有任何评论")
    Else
        strComment = strComment & "<table width='100%' align='center' border='0' cellspacing='1' cellpadding='2' class='comment_border'>"
        If ShowVote = 1 Then 	
            strComment = strComment & "<tr class='comment_title' align='center'><td>评论人</td><td>评论内容</td><td>评论时间</td><td>打分</td><td width=150>投票</td></tr>"
        Else
            strComment = strComment & XmlText("Site", "Comment/ShowCommentListJs", "<tr class='comment_title' align='center'><td>评论人</td><td>评论内容</td><td>评论时间</td><td>打分</td></tr>")		       
        End If			    
		Dim strCommentList, strCommentReply1, strCommentReply2
        If ShowVote = 1 Then 			
            strCommentList = "<tr bgcolor='white'><td align='center' width='80'>{$CommentName}</td><td >{$Content}</td><td align='center' width='120'>{$WriteTime}</td><td align='center' width='30'>{$Score}分</td>"
            strCommentList = strCommentList &"<td width='150' align='center'><span id=Support{$CommentID}><a href=""javascript:Support({$CommentID});"">支持</a></span>[<span id=SupportCount{$CommentID}>{$SupportCount}</span>] <span id=Opposed{$CommentID}><a href=""javascript:Opposed({$CommentID});"">反对</a></span>[<span id=OpposedCount{$CommentID}>{$OpposedCount}</span>]</span></a></td>"
            strCommentList = strCommentList & "</tr>"
        Else
            strCommentList = XmlText("Site", "Comment/ShowCommentJs", "<tr bgcolor='white'><td align='center' width='80'>{$CommentName}</td><td>{$Content}</td><td align='center' width='120'>{$WriteTime}</td><td align='center' width='40'>{$Score}分</td></tr>")        
        End If							
        strCommentReply1 = "<tr><td>&nbsp;</td><td colspan='6'><font color='009900'>★</font>&nbsp;发布人『<font color='blue'>{$AdminName}</font>』于 {$ReplyTime} 回复道：&nbsp;&nbsp;&nbsp;&nbsp;{$ReplyContent}<br></td></tr>"
        strCommentReply2 = XmlText("Site", "Comment/AdminReplayTypeJs", "<tr><td>&nbsp;</td><td colspan='5'><font color='009900'>★</font>&nbsp;管理员『<font color='blue'>{$AdminName}</font>』于 {$ReplyTime} 回复道：&nbsp;&nbsp;&nbsp;&nbsp;{$ReplyContent}<br></td></tr>")
        Do While Not rsComment.EOF
            If rsComment("UserType") = 1 Then
                strUserName = "【<a href='" & strInstallDir & "ShowUser.asp?UserName=" & rsComment("UserName") & "'><font color='green'>" & rsComment("UserName") & "</font></a>】"
            Else
                strUserName = "【<span title='类别：游客" & "\n" & "姓名：" & rsComment("UserName") & "\n" & "信箱：" & rsComment("Email") & "\n" & "Oicq：" & rsComment("Oicq") & "\n" & "主页：" & rsComment("Homepage") & "' style='cursor:hand'>" & rsComment("UserName") & "</span>】"
            End If
            
            strComment = strComment &Replace(Replace(Replace(Replace(Replace(Replace(Replace(strCommentList, "{$CommentName}", strUserName), "{$Content}", FilterJS(Replace(ReplaceText(rsComment("Content"), 3), vbCrLf, "\n"))), "{$WriteTime}", rsComment("WriteTime")), "{$Score}", rsComment("Score")), "{$CommentID}",rsComment("CommentID")),"{$SupportCount}",PE_CLng(rsComment("Support"))),"{$OpposedCount}",PE_CLng(rsComment("Opposed"))) 
            If rsComment("ReplyContent") <> "" Then
                If ModuleType = 6 Then
                    strComment = strComment & Replace(Replace(Replace(strCommentReply1, "{$AdminName}", rsComment("ReplyName")), "{$ReplyTime}", rsComment("ReplyTime")), "{$ReplyContent}", Replace(rsComment("ReplyContent"), vbCrLf, "\n"))
                Else
                    strComment = strComment & Replace(Replace(Replace(strCommentReply2, "{$AdminName}", rsComment("ReplyName")), "{$ReplyTime}", rsComment("ReplyTime")), "{$ReplyContent}", Replace(rsComment("ReplyContent"), vbCrLf, "\n"))
                End If
            End If
            rsComment.MoveNext
        Loop
        rsComment.Close
        Set rsComment = Nothing
        strComment = strComment & "</table>"
        strComment = strComment & Replace(XmlText("Site", "Comment/ShowMore", "<div align='center'><a href='{$strFileName}'>查看评论详细内容及更多评论</a></div>"), "{$strFileName}", ChannelUrl & "/" & strFileName)
    End If
    Response.Write "document.write(""" & Replace(strComment, """", "\""") & """);"
End Sub

Sub UpdateVote()
    Dim id
    Dim Rs,Sql ,Votetype,num
    id = Replace(Trim(Request.QueryString("id")),"'","")
    Votetype = PE_Clng(Request("votetype"))
    Set Rs = Server.CreateObject("ADODB.Recordset")
    Sql = "Select * From PE_Comment Where Commentid="&PE_Clng(id)
    Rs.Open Sql,Conn,3,3
    If Rs.Eof And Rs.Bof Then
        Response.Write("nodate")
    Else
        if Votetype = 1 Then
            num =PE_CLng(Rs("Support"))
            num = num + 1
            Rs("Support") = num
        Else
            num = PE_Clng(Rs("Opposed"))
            num = num + 1
            Rs("Opposed") = num	
        End If	
        Rs.Update
        Rs.Close
        Set Rs = Nothing
        Response.Write PE_Clng(num)
    End If
End Sub

'**************************************************
'函数名：GetInfoUrl
'作  用：得到文章、下载、图片、商品的Url地址
'参  数：
'返回值：替换后字符串
'**************************************************
Function GetInfoUrl(ByVal tParentDir, ByVal tClassDir, ByVal tUpdateTime, ByVal tInfoID, ByVal tClassPurview, ByVal tInfoPurview, ByVal tInfoPoint)
    If IsNull(tParentDir) Then tParentDir = ""
    If IsNull(tClassDir) Then tClassDir = ""
    If IsNull(tClassPurview) Then tClassPurview = 0
    If IsNull(tInfoPurview) Then tInfoPurview = 0

    Select Case ModuleType
    Case 1
        If UseCreateHTML > 0 And tClassPurview = 0 And tInfoPoint = 0 And tInfoPurview = 0 Then
            GetInfoUrl = ChannelUrl & GetItemPath(StructureType, tParentDir, tClassDir, tUpdateTime) & GetItemFileName(FileNameType, ChannelDir, tUpdateTime, tInfoID) & FileExt_Item
        Else
            GetInfoUrl = ChannelUrl_ASPFile & "/ShowArticle.asp?ArticleID=" & tInfoID
        End If
    Case 2
        If UseCreateHTML > 0 Then
            GetInfoUrl = ChannelUrl & GetItemPath(StructureType, tParentDir, tClassDir, tUpdateTime) & GetItemFileName(FileNameType, ChannelDir, tUpdateTime, tInfoID) & FileExt_Item
        Else
            GetInfoUrl = ChannelUrl_ASPFile & "/ShowSoft.asp?SoftID=" & tInfoID
        End If
    Case 3
        If UseCreateHTML > 0 And tClassPurview = 0 And tInfoPoint = 0 And tInfoPurview = 0 Then
            GetInfoUrl = ChannelUrl & GetItemPath(StructureType, tParentDir, tClassDir, tUpdateTime) & GetItemFileName(FileNameType, ChannelDir, tUpdateTime, tInfoID) & FileExt_Item
        Else
            GetInfoUrl = ChannelUrl_ASPFile & "/ShowPhoto.asp?PhotoID=" & tInfoID
        End If
    Case 5
        If UseCreateHTML > 0 Then
            GetInfoUrl = ChannelUrl & GetItemPath(StructureType, tParentDir, tClassDir, tUpdateTime) & GetItemFileName(FileNameType, ChannelDir, tUpdateTime, tInfoID) & FileExt_Item
        Else
            GetInfoUrl = ChannelUrl_ASPFile & "/ShowProduct.asp?ProductID=" & tInfoID
        End If
    Case 6
        GetInfoUrl = strInstallDir & ChannelDir & "/ShowSupply.asp?SupplyID=" & tInfoID
    End Select
End Function



%>
