<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Function GetChannelName(iChannelID)
    Dim strChannelName, rsChannel, sqlChannel
    If iChannelID = -1 Then
        strChannelName = "��վͨ��"
    ElseIf iChannelID = 0 Then
        strChannelName = "��վ��ҳ"
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
        strChannel = strChannel & "<a href='" & strFileName & "&ChannelID=-1'><font color=red>��վͨ��</font></a> | "
    Else
        strChannel = strChannel & "<a href='" & strFileName & "&ChannelID=-1'>��վͨ��</a> | "
    End If
    If iChannelID = 0 Then
        strChannel = strChannel & "<a href='" & strFileName & "&ChannelID=0'><font color=red>��վ��ҳ</font></a> | "
    Else
        strChannel = strChannel & "<a href='" & strFileName & "&ChannelID=0'>��վ��ҳ</a> | "
    End If
    sqlChannel = "select * from PE_Channel where ChannelType<=1 and Disabled=" & PE_False & " order by OrderID"
    Set rsChannel = Conn.Execute(sqlChannel)
    If rsChannel.BOF And rsChannel.EOF Then
        strChannel = strChannel & "û���κ�Ƶ��"
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
    strTemp = strTemp & ">��վͨ��</option>"
    strTemp = strTemp & "<option value='0'"
    If iChannelID = 0 Then strTemp = strTemp & " selected"
    strTemp = strTemp & ">��վ��ҳ</option>"
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
    Response.Write "�����ڵ�λ�ã���վ" & ItemName & "����&nbsp;&gt;&gt;&nbsp;"
    If iChannelID = -1 Then
        Response.Write "Ƶ������" & ItemName & ""
    ElseIf iChannelID = 0 Then
        Response.Write "��վ��ҳ" & ItemName & ""
    Else
        Dim rsPath
        Set rsPath = Conn.Execute("select ChannelID,ChannelName from PE_Channel where ChannelID=" & iChannelID)
        If rsPath.BOF And rsPath.EOF Then
            Response.Write "�����Ƶ������"
        Else
            Response.Write "<a href='" & strFileName & "&ChannelID=" & rsPath(0) & "'>" & rsPath(1) & ItemName & "</a>"
        End If
        rsPath.Close
        Set rsPath = Nothing
    End If
End Sub

%>
