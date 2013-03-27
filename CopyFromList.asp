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

PageTitle = "来源列表"

Dim strtmp, HotNum
Dim SourceName, i, rsCopyFrom, sqlCopyFrom
ChannelID = PE_Clng(Trim(Request("ChannelID")))

If ChannelID > 0 Then
    strFileName = "CopyfromList.asp?ChannelID=" & ChannelID
    Call GetChannel(ChannelID)
Else
    ChannelName = "全部频道"
End If
MaxPerPage = 20
strNavPath = strNavPath & strNavLink & "&nbsp;" & PageTitle
HotNum = 100

strHTML = GetTemplate(0, 13, 0)
Call ReplaceCommonLabel

strHTML = Replace(strHTML, "{$PageTitle}", SiteTitle & " >> " & PageTitle)
strHTML = Replace(strHTML, "{$ShowPath}", strNavPath)
strHTML = Replace(strHTML, "{$MenuJS}", GetMenuJS("", False))
strHTML = Replace(strHTML, "{$Skin_CSS}", GetSkin_CSS(0))

strHTML = Replace(strHTML, "{$ShowCopyFromList}", GetCopyfromList())
Response.Write strHTML
Call CloseConn

Function GetCopyfromList()
    strtmp = "<table width=100% cellspacing='1' cellpadding='0' border='0'>"
    strtmp = strtmp & "    <tr class='main_title_760' align='center'><td style='word-break:break-all;Width:fixed'><b>" & ChannelName & "来源列表" & "</b></td></tr>"
    strtmp = strtmp & "<tr><td><table width=100% cellspacing='1' cellpadding='0' border='0' class='Channel_border'><tr class='Channel_title' align='center'><td width='150'>" & XmlText("ShowSource", "ShowCopyFromList/t1", "名称") & "</td><td width='60'>" & XmlText("ShowSource", "ShowCopyFromList/t2", "主页") & "</td><td width='80'>" & XmlText("ShowSource", "ShowCopyFromList/t3", "所属频道") & "</td><td width='470'>" & XmlText("ShowSource", "ShowCopyFromList/t4", "来源简介") & "</td></tr>"
    sqlCopyFrom = "select * from PE_CopyFrom Where Passed=" & PE_True
    If ChannelID > 0 Then sqlCopyFrom = sqlCopyFrom & " and ChannelID=" & ChannelID
    If Action = "ListElite" Then sqlCopyFrom = sqlCopyFrom & " and isElite=" & PE_True
    If Action = "ListHot" Then sqlCopyFrom = sqlCopyFrom & " and Hits>" & HotNum
    sqlCopyFrom = sqlCopyFrom & " order by onTop " & PE_OrderType & ",LastUseTime Desc"
    Set rsCopyFrom = Server.CreateObject("ADODB.Recordset")
    rsCopyFrom.Open sqlCopyFrom, Conn, 1, 1
    If rsCopyFrom.BOF And rsCopyFrom.EOF Then
        totalPut = 0
        strtmp = strtmp & "<tr class='Channel_tdbg'><td align='center'colspan='4'>" & XmlText("ShowSource", "ShowCopyFromList/NoFoundCopyFrom", "本频道尚未添加来源") & "</td></tr></table></td></tr></table>"
    Else
        totalPut = rsCopyFrom.RecordCount
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                        rsCopyFrom.Move (CurrentPage - 1) * MaxPerPage
                Else
                    CurrentPage = 1
                End If
        End If
        i = 0
        Do While Not rsCopyFrom.EOF
            If rsCopyFrom("ChannelID") <> PrevChannelID Then
                Call GetChannel(rsCopyFrom("ChannelID"))
                PrevChannelID = rsCopyFrom("ChannelID")
            End If
            strtmp = strtmp & ("<tr class='Channel_tdbg'><td align='center'><a href='ShowCopyFrom.asp?ChannelID=" & rsCopyFrom("ChannelID") & "&SourceName=" & rsCopyFrom("SourceName") & "'>" & rsCopyFrom("SourceName") & "</a></td><td align='center'>")
            strtmp = strtmp & rsCopyFrom("HomePage")
            strtmp = strtmp & ("</td><td>&nbsp;" & ChannelName & "</td><td>&nbsp;" & Left(nohtml(rsCopyFrom("Intro")), 200) & "</td></tr>")
            rsCopyFrom.MoveNext
            i = i + 1
            If i >= MaxPerPage Then Exit Do
        Loop
        strtmp = strtmp & ("</table></td></tr><tr Class='Channel_pager'><td align='center' colspan='4'>" & ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("ShowSource", "ShowCopyFromList/PageChar", "个来源"), False) & "</td></tr></Table>")
    End If
    rsCopyFrom.Close
    Set rsCopyFrom = Nothing
    GetCopyfromList = strtmp
End Function
%>
