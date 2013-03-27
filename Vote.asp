<!--#include file="Start.asp"-->
<!--#include file="Include/PowerEasy.Cache.asp"-->
<!--#include file="Include/PowerEasy.Common.Front.asp"-->
<!--#include file="Include/PowerEasy.Channel.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

ChannelID = 0
Dim ID, VoteType, VoteOption, arrOptions, sqlVote, rsVote, IsItem
Dim Voted, VotedID, arrVotedID, i
Dim strPath

ID = PE_CLng(Trim(Request("ID")))
VoteType = Trim(Request("VoteType"))
VoteOption = ReplaceBadChar(Trim(Request("VoteOption")))
VotedID = ReplaceBadChar(Trim(Request.Cookies("VotedID")))
If IsValidID(VotedID) = False Then
    VotedID = ""
End If
If ID = 0 Then
    FoundErr = True
    ErrMsg = ErrMsg + "<li>不能确定调查ID</li>"
    Call WriteErrMsg(ErrMsg, ComeUrl)
    Response.End
End If
Voted = FoundInArr(VotedID, ID, ",")

If Action = "" Or VoteOption = "" Then
    Action = "Show"
End If
If Action = "Vote" And VoteOption <> "" And Voted = False Then
    If VoteType = "Single" Then
        VoteOption = PE_CLng(VoteOption)
        Conn.Execute "Update PE_Vote set answer" & VoteOption & "= answer" & VoteOption & "+1 where ID=" & ID
    Else
        If InStr(VoteOption, ",") > 0 Then
            arrOptions = Split(VoteOption, ",")
            For i = 0 To UBound(arrOptions)
                Conn.Execute "Update PE_Vote set answer" & PE_CLng(Trim(arrOptions(i))) & "= answer" & PE_CLng(Trim(arrOptions(i))) & "+1 where ID=" & ID
            Next
        Else
            Conn.Execute "Update PE_Vote set answer" & PE_CLng(VoteOption) & "= answer" & PE_CLng(VoteOption) & "+1 where ID=" & ID
        End If
    End If
    If VotedID = "" Then
        VotedID = Trim(ID)
    Else
        VotedID = VotedID & "," & ID
    End If
    Response.Cookies("VotedID") = VotedID
End If

Set rsVote = Conn.Execute("Select * from PE_Vote Where ID=" & ID)
If rsVote.BOF Or rsVote.EOF Then
    FoundErr = True
    ErrMsg = ErrMsg & ("<li>" & XmlText("Site", "ShowVote/NoVote", "找不到相应的调查") & "</li>")
    Call WriteErrMsg(ErrMsg, ComeUrl)
    rsVote.Close
    Set rsVote = Nothing
    Response.End
End If

Dim VoteTips
If Action = "Vote" And VoteOption <> "" Then
    VoteTips = "<br><font color='#FF0000' size='3'>"
    If Voted = True Then
        VoteTips = VoteTips & XmlText("Site", "ShowVote/VoteED", "==　你已经投过票了，请勿重复投票！　==")
    Else
        VoteTips = VoteTips & XmlText("Site", "ShowVote/PreVote", "==　非常感谢您的投票！　==")
    End If
    VoteTips = VoteTips & "</font><br><br>"
Else
    VoteTips = ""
End If

Dim TotalVote
TotalVote = 0
For i = 1 To 20
    If IsNull(rsVote("Select" & i)) Or rsVote("Select" & i) = "" Then Exit For
    TotalVote = TotalVote + PE_CLng(rsVote("answer" & i))
Next

PageTitle = "网站调查"
strHtml = GetTemplate(ChannelID, 6, 0)

Call ReplaceCommonLabel

strNavPath = strNavPath & strNavLink & "&nbsp;" & PageTitle

strHtml = Replace(strHtml, "{$PageTitle}", SiteTitle & " >> " & PageTitle)
strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)

strHtml = Replace(strHtml, "{$MenuJS}", GetMenuJS("", False))
strHtml = Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))

Dim strVoteTitle
strVoteTitle = Replace(Replace(Replace(Replace(Replace(rsVote("Title"), "{$Date}", FormatDateTime(Now(), 2)), "{$Year}", Year(Now())), "{$Month}", Month(Now())), "{$Day}", Day(Now())), "{$Weekday}", WeekDayName(Weekday(Now())))
strHtml = Replace(strHtml, "{$VoteTitle}", strVoteTitle)
strHtml = Replace(strHtml, "{$VoteTips}", VoteTips)
strHtml = Replace(strHtml, "{$TotalVote}", TotalVote)

If TotalVote = 0 Then TotalVote = 1

Dim strVoteItems, strVoteItem, strTemp, perVote, lngTemp

regEx.Pattern = "\[VoteItem\]([\s\S]*?)\[\/VoteItem\]"
Set Matches = regEx.Execute(strHtml)
For Each Match In Matches
    strTemp = Match.value
Next

For i = 1 To 20
    If Trim(rsVote("Select" & i) & "") = "" Then Exit For
    lngTemp = PE_CLng(rsVote("answer" & i))
    perVote = Round(lngTemp / TotalVote, 4)
    
    strVoteItem = Replace(Replace(strTemp, "[VoteItem]", ""), "[/VoteItem]", "")
    strVoteItem = Replace(strVoteItem, "{$ItemNum}", i)
    strVoteItem = Replace(strVoteItem, "{$ItemSelect}", rsVote("Select" & i))
    strVoteItem = Replace(strVoteItem, "{$ItemAnswer}", lngTemp)
    strVoteItem = Replace(strVoteItem, "{$ItemPer}", perVote * 100)
    strVoteItem = Replace(strVoteItem, "{$ItemWidth}", CLng(500 * perVote))
    strVoteItem = Replace(strVoteItem, "{$ItemWidth2}", CLng(500 * (1 - perVote)))
    
    strVoteItems = strVoteItems & strVoteItem
Next
strHtml = Replace(strHtml, strTemp, strVoteItems)

Dim strVoteForm
If Action = "Show" And Voted = False Then
    strVoteForm = "<br><b>&nbsp;&middot;" & XmlText("Site", "ShowVote/Vote1", "您还没有投票，请您在此投下您宝贵的一票！") & "</b>"
    strVoteForm = strVoteForm & "<form name='VoteForm' method='post' action='vote.asp'>"

    strVoteForm = strVoteForm & "&nbsp;" & strVoteTitle & "<br>"
    If rsVote("VoteType") = "Single" Then
        For i = 1 To 20
            If Trim(rsVote("Select" & i) & "") = "" Then Exit For
            strVoteForm = strVoteForm & "<input type='radio' name='VoteOption' style='border:0' value='" & i & "'>" & rsVote("Select" & i) & "<br>"
        Next
    Else
        For i = 1 To 20
            If Trim(rsVote("Select" & i) & "") = "" Then Exit For
            strVoteForm = strVoteForm & "<input type='checkbox' name='VoteOption' style='border:0' value='" & i & "'>" & rsVote("Select" & i) & "<br>"
        Next
    End If
    strVoteForm = strVoteForm & "<br><input name='VoteType' type='hidden'value='" & rsVote("VoteType") & "'>"
    strVoteForm = strVoteForm & "<input name='Action' type='hidden' value='Vote' style='border:0'>"
    strVoteForm = strVoteForm & "<input name='ID' type='hidden' value='" & rsVote("ID") & "' style='border:0'>"
    strVoteForm = strVoteForm & "&nbsp;&nbsp;&nbsp;&nbsp;<a href='javascript:VoteForm.submit();'><img src='images/voteSubmit.gif' width='52' height='18' border='0'></a>&nbsp;&nbsp;"
    strVoteForm = strVoteForm & "<a href='Vote.asp?ID=" & rsVote("ID") & "&Action=Show' target='_blank'><img src='images/voteView.gif' width='52' height='18' border='0'></a>"
    strVoteForm = strVoteForm & "</form>"
Else
    strVoteForm = ""
End If
strHtml = Replace(strHtml, "{$VoteForm}", strVoteForm)
IsItem = rsVote("IsItem")
rsVote.Close
Set rsVote = Nothing

If IsItem = False Then
    Dim sqlOtherVote, rsOtherVote, strOtherVote
    If VotedID = "" Then
        sqlOtherVote = "Select * from PE_Vote Where ID <>" & ID & " order by ID desc"
    Else
        sqlOtherVote = "Select * from PE_Vote Where ID Not In (" & VotedID & ") order by ID desc"
    End If
    Set rsOtherVote = Conn.Execute(sqlOtherVote)
    If rsOtherVote.BOF And rsOtherVote.EOF Then
        If Action = "Vote" Then
            strOtherVote = "<br>" & XmlText("Site", "ShowVote/Vote2", "感谢您参加了本站的所有调查！！！")
        Else
            strOtherVote = ""
        End If
    Else
        strOtherVote = "<br>" & XmlText("Site", "ShowVote/Vote3", "欢迎你继续参加本站的其他调查：") & "<br><br>"
        Do While Not rsOtherVote.EOF
            strVoteTitle = Replace(Replace(Replace(Replace(Replace(rsOtherVote("Title"), "{$Date}", FormatDateTime(Now(), 2)), "{$Year}", Year(Now())), "{$Month}", Month(Now())), "{$Day}", Day(Now())), "{$Weekday}", WeekDayName(Weekday(Now())))
            
            strOtherVote = strOtherVote & "<li><a class='LinkVote' href='Vote.asp?ID=" & rsOtherVote("ID") & "'>" & strVoteTitle & "</a></li>"
            rsOtherVote.MoveNext
        Loop
    End If
    rsOtherVote.Close
    Set rsOtherVote = Nothing
    strHtml = Replace(strHtml, "{$OtherVote}", strOtherVote)
Else
    strHtml = Replace(strHtml, "{$OtherVote}", "")
End If

Response.Write strHtml
%>
