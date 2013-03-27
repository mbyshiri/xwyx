<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Function GetChannelName(iChannelID)
    Dim strChannelName, rsChannel, sqlChannel
    If iChannelID = -1 Then
        strChannelName = "网站通用"
    ElseIf iChannelID = 0 Then
        strChannelName = "网站首页"
    Else
        sqlChannel = "select ChannelName from PE_Channel where ChannelID=" & iChannelID
        Set rsChannel = Conn.Execute(sqlChannel)
        If rsChannel.BOF And rsChannel.EOF Then
            strChannelName = ""
        Else
            strChannelName = rsChannel(0)
        End If
        rsChannel.Close
        Set rsChannel = Nothing
    End If
    GetChannelName = strChannelName
End Function

Function GetChannelList(iChannelID)
    Dim rsChannel, sqlChannel, strChannel, i
    strChannel = "|&nbsp;"
    If iChannelID = -1 Then
        strChannel = strChannel & "<a href='" & strFileName & "&ChannelID=-1'><font color=red>网站通用</font></a> | "
    Else
        strChannel = strChannel & "<a href='" & strFileName & "&ChannelID=-1'>网站通用</a> | "
    End If
    If iChannelID = 0 Then
        strChannel = strChannel & "<a href='" & strFileName & "&ChannelID=0'><font color=red>网站首页</font></a> | "
    Else
        strChannel = strChannel & "<a href='" & strFileName & "&ChannelID=0'>网站首页</a> | "
    End If
    sqlChannel = "select * from PE_Channel where ChannelType<=1 and Disabled=" & PE_False & " order by OrderID"
    Set rsChannel = Conn.Execute(sqlChannel)
    If rsChannel.BOF And rsChannel.EOF Then
        strChannel = strChannel & "没有任何频道"
    Else
        i = 1
        Do While Not rsChannel.EOF
            If rsChannel("ChannelID") = iChannelID Then
                strChannel = strChannel & "<a href='" & strFileName & "&ChannelID=" & iChannelID & "'><font color=red>" & rsChannel("ChannelName") & "</font></a>"
            Else
                strChannel = strChannel & "<a href='" & strFileName & "&ChannelID=" & rsChannel("ChannelID") & "'>" & rsChannel("ChannelName") & "</a>"
            End If
            strChannel = strChannel & " | "
            i = i + 1
            If i Mod 10 = 0 Then
                strChannel = strChannel & "<br>"
            End If
            rsChannel.MoveNext
        Loop
    End If
    rsChannel.Close
    Set rsChannel = Nothing
    GetChannelList = strChannel
End Function

Function GetChannel_Option(iChannelID)
    Dim strTemp, rsChannel, sqlChannel
    strTemp = strTemp & "<option value='-1'"
    If iChannelID = -1 Then strTemp = strTemp & " selected"
    strTemp = strTemp & ">网站通用</option>"
    strTemp = strTemp & "<option value='0'"
    If iChannelID = 0 Then strTemp = strTemp & " selected"
    strTemp = strTemp & ">网站首页</option>"
    sqlChannel = "select * from PE_Channel where ChannelType<=1 and Disabled=" & PE_False & " order by OrderID"
    Set rsChannel = Conn.Execute(sqlChannel)
    Do While Not rsChannel.EOF
        strTemp = strTemp & "<option value='" & rsChannel("ChannelID") & "'"
        If iChannelID = rsChannel("ChannelID") Then strTemp = strTemp & " selected"
        strTemp = strTemp & ">" & rsChannel("ChannelName") & "</option>"
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    GetChannel_Option = strTemp
End Function

Sub ShowManagePath(iChannelID)
    Response.Write "您现在的位置：网站" & ItemName & "管理&nbsp;&gt;&gt;&nbsp;"
    If iChannelID = -1 Then
        Response.Write "频道共用" & ItemName & ""
    ElseIf iChannelID = 0 Then
        Response.Write "网站首页" & ItemName & ""
    Else
        Dim rsPath
        Set rsPath = Conn.Execute("select ChannelID,ChannelName from PE_Channel where ChannelID=" & iChannelID)
        If rsPath.BOF And rsPath.EOF Then
            Response.Write "错误的频道参数"
        Else
            Response.Write "<a href='" & strFileName & "&ChannelID=" & rsPath(0) & "'>" & rsPath(1) & ItemName & "</a>"
        End If
        rsPath.Close
        Set rsPath = Nothing
    End If
End Sub

%>
