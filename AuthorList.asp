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

PageTitle = "作者列表"
strFileName = "AuthorList.asp"
Dim strtmp, HotNum
Dim tempChannelID
Dim iUsername, i, rsAuthor, sqlAuthor
ChannelID = PE_Clng(Trim(Request("ChannelID")))
If ChannelID > 0 Then
    strFileName = "AuthorList.asp?ChannelID=" & ChannelID
    Call GetChannel(ChannelID)
Else
    ChannelName = "全部频道"
End If
iUsername = Trim(Request("Username"))
If iUsername <> "" Then
    iUsername = ReplaceBadChar(iUsername)
    strFileName = strFileName & "&Username=" & iUsername
End If
MaxPerPage = 20
strNavPath = strNavPath & strNavLink & "&nbsp;" & PageTitle
HotNum = 100

strHTML = GetTemplate(0, 11, 0)
Call ReplaceCommonLabel

strHTML = Replace(strHTML, "{$PageTitle}", SiteTitle & " >> " & PageTitle)
strHTML = Replace(strHTML, "{$ShowPath}", strNavPath)
strHTML = Replace(strHTML, "{$MenuJS}", GetMenuJS("", False))
strHTML = Replace(strHTML, "{$Skin_CSS}", GetSkin_CSS(0))

strHTML = Replace(strHTML, "{$GetAuthorList}", ShowAuthorList())
Response.Write strHTML
Call CloseConn

Function ShowAuthorList()
    strtmp = "<table width=100% cellspacing='1' cellpadding='0' border='0'>"
    strtmp = strtmp & "    <tr class='main_title_760' align='center'><td style='word-break:break-all;Width:fixed'><b>" & ChannelName & "作者列表" & "</b></td></tr>"
    strtmp = strtmp & "<tr><td><table width=100% cellspacing='1' cellpadding='0' border='0' class='Channel_border'><tr class='Channel_title' align='center'><td width='150'>" & XmlText("ShowSource", "ShowAuthorList/t1", "姓名") & "</td><td width='60'>" & XmlText("ShowSource", "ShowAuthorList/t2", "性别") & "</td><td width='80'>" & XmlText("ShowSource", "ShowAuthorList/t3", "所属频道") & "</td><td>" & XmlText("ShowSource", "ShowAuthorList/t4", "作者简介") & "</td></tr>"
    sqlAuthor = "select * from PE_Author Where Passed=" & PE_True
    If ChannelID > 0 Then sqlAuthor = sqlAuthor & " and (ChannelID=" & ChannelID &" Or ChannelID = 0)"
    If iUsername <> "" Then sqlAuthor = sqlAuthor & " and UserName='" & iUsername & "'"
    If Action = "ListElite" Then sqlAuthor = sqlAuthor & " and isElite=" & PE_True
    If Action = "ListHot" Then sqlAuthor = sqlAuthor & " and Hits>" & HotNum
    sqlAuthor = sqlAuthor & " order by onTop " & PE_OrderType & ",LastUseTime Desc"
    Set rsAuthor = Server.CreateObject("ADODB.Recordset")
    rsAuthor.Open sqlAuthor, Conn, 1, 1
    If rsAuthor.BOF And rsAuthor.EOF Then
        totalPut = 0
        strtmp = strtmp & "<tr class='Channel_tdbg'><td align='center'colspan='4'>" & XmlText("ShowSource", "ShowAuthorList/NoFoundAuthor", "本频道尚未添加作者") & "</td></tr></table>"
    Else
        totalPut = rsAuthor.RecordCount
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                rsAuthor.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
        i = 0
        Do While Not rsAuthor.EOF
            If PE_CLng(rsAuthor("ChannelID")) = 0 Then
                tempChannelID = ChannelID
            Else
                tempChannelID = PE_CLng(rsAuthor("ChannelID"))                
            End If		
            If tempChannelID > 0 Then
                strtmp = strtmp & "<tr class='Channel_tdbg'><td align='center'><a href='ShowAuthor.asp?ChannelID=" & tempChannelID & "&AuthorName=" & rsAuthor("AuthorName") & "'>" & rsAuthor("AuthorName") & "</a></td><td align='center'>"
            Else
                strtmp = strtmp & "<tr class='Channel_tdbg'><td align='center'>" & rsAuthor("AuthorName") & "</td><td align='center'>"
            End If
            If rsAuthor("Sex") = 1 Then
                strtmp = strtmp & strMan
            Else
                strtmp = strtmp & strGirl
            End If
            
            If rsAuthor("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsAuthor("ChannelID"))
                PrevChannelID = rsAuthor("ChannelID")
            End If
            strtmp = strtmp & "</td><td>&nbsp;" & ChannelName & "</td><td>&nbsp;" & Left(nohtml(rsAuthor("Intro")), 200) & "</td></tr>"
            rsAuthor.MoveNext
            i = i + 1
            If i >= MaxPerPage Then Exit Do
        Loop
        strtmp = strtmp & "</table></td></tr><tr class='Channel_pager'><td align='center'>" & ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("ShowSource", "ShowAuthorList/PageChar", "个作者"), False) & "</td></tr></Table>"
    End If
    rsAuthor.Close
    Set rsAuthor = Nothing
    ShowAuthorList = strtmp
End Function
%>
