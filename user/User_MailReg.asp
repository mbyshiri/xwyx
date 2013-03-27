<!--#include file="CommonCode.asp"-->
<%
Select Case Action
Case "SaveAdd"
    Call Savemail
Case "Modify"
    Call Modify
Case Else
    Call main
End Select

Function GetChannels(IsOrder)
    Dim rsChannel, strChannels, sqlChannel, rsGetChannel, arrGetChannel
    Set rsGetChannel = Conn.Execute("select C.arrChannelID from PE_Contacter C inner join PE_User U on U.ContacterID=C.ContacterID where U.UserID = " & UserID)
    arrGetChannel = rsGetChannel("arrChannelID")
    If IsValidID(arrGetChannel) = False Then
        arrGetChannel = ""
    End If
    sqlChannel = "select * from PE_MailChannel M inner join PE_Channel C on M.ChannelID = C.ChannelID where C.ModuleType=1 and M.IsUse = " & PE_True & " "
    If arrGetChannel = "" Or IsNull(arrGetChannel) Then
        If IsOrder = True Then
            strChannels = "暂无内容"
            sqlChannel = sqlChannel & "and 1=2"
        Else
        End If
    Else
        If IsOrder = True Then
            sqlChannel = sqlChannel & " and M.ChannelID in(" & arrGetChannel & ")"
        Else
            sqlChannel = sqlChannel & " and M.ChannelID not in(" & arrGetChannel & ")"
        End If
    End If
    Set rsChannel = Conn.Execute(sqlChannel)
    If rsChannel.bof And rsChannel.EOF Then
        strChannels = "暂无内容"
        Exit Function
    End If
    Do While Not rsChannel.EOF
        If IsOrder = True Then
            strChannels = strChannels & "<dl><dt><input type='checkbox' name='ChannelID' value='" & rsChannel(ChannelID) & "' checked >" & rsChannel("ChannelName") & "</dt><dd></dd></dl>" & vbCrLf
        Else
            strChannels = strChannels & "<dl><dt><input type='checkbox' name='ChannelID' value='" & rsChannel(ChannelID) & "'>" & rsChannel("ChannelName") & "</dt><dd></dd></dl>" & vbCrLf
        End If
    rsChannel.movenext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    GetChannels = strChannels
End Function

Sub main()
    Dim rsGetChannel
    Set rsGetChannel = Conn.Execute("select C.arrChannelID from PE_Contacter C inner join PE_User U on U.ContacterID=C.ContacterID where U.UserID = " & UserID)
    If rsGetChannel.bof And rsGetChannel.EOF Then
        Response.Write "<script   language='javascript'>alert(""您的资料不完善,请先去会员中心完善您的资料！"");location.href=""User_Info.asp?Action=Modify"";</script>"
        rsGetChannel.Close
        Set rsGetChannel = Nothing
        Exit Sub
    End If
    
    Dim msql, rsUser
    msql = "select Email,UserID,UserName from PE_User where UserID=" & UserID & ""
    Set rsUser = Server.CreateObject("adodb.recordset")
    rsUser.Open msql, Conn, 1, 1
    Response.Write "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'> " & vbCrLf
    Response.Write "<html xmlns='http://www.w3.org/1999/xhtml'> " & vbCrLf
    Response.Write "<head> " & vbCrLf
    Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'> " & vbCrLf
    Response.Write "<title>邮件订阅服务</title> " & vbCrLf
    Response.Write "<link rel='stylesheet' type='text/css' href='images/Common.css' /> " & vbCrLf
    Response.Write "</head> " & vbCrLf
    Response.Write "<body> " & vbCrLf
    Response.Write "<div class='mailListContainer'> " & vbCrLf
    Response.Write "<div class='top'> " & vbCrLf
    Response.Write "<img src='images/maillogo.gif' alt='邮件订阅服务' /> " & vbCrLf
    Response.Write "</div> " & vbCrLf
    Response.Write "<div id='mailList_set'> " & vbCrLf
    Response.Write "<div class='title'> " & vbCrLf
    Response.Write "<p>欢迎使用本站邮件订阅服务。您可以通过邮件的形式获取感兴趣的最新信息。<p> " & vbCrLf
    Response.Write "<p>您接收订阅的邮箱是：<span class='mailAccount'>" & rsUser("Email") & "<br /></p> " & vbCrLf
    Response.Write "</div> " & vbCrLf
    Response.Write "<form name='myform' method='post' action='User_MailReg.asp'  target='_top'>" & vbCrLf
    Response.Write "<div class='listRow' id='mailListColumn_2'> " & vbCrLf
    Response.Write "<div class='column'> " & vbCrLf
    Response.Write "<h3>您已经订阅的频道：</h3> " & vbCrLf
    Response.Write "<span class='descRecommand'></span> " & vbCrLf
    Response.Write "</div> " & vbCrLf
    Response.Write "   " & GetChannels(True) & " " & vbCrLf
    Response.Write "<div class='listRow'> " & vbCrLf
    Response.Write "<h3 style='color:red'>说明：如果想退订请在频道前面的勾选框取消勾选</h3> " & vbCrLf
    Response.Write "</div> " & vbCrLf
    Response.Write "</div> " & vbCrLf
    Response.Write "<div class='listRow' id='mailListColumn_2'> " & vbCrLf
    Response.Write "<div class='column'> " & vbCrLf
    Response.Write "<h3>你要订阅的频道：</h3> " & vbCrLf
    Response.Write "<span class='descRecommand'></span> " & vbCrLf
    Response.Write "</div> " & vbCrLf
    Response.Write "   " & GetChannels(False) & " " & vbCrLf
    Response.Write "</div> " & vbCrLf
    Response.Write "<div class='listRow'> " & vbCrLf
    Response.Write "<h3>订阅说明：</h3> " & vbCrLf
    Response.Write "<ul> " & vbCrLf
    Response.Write "<li>普通栏目一般每日发送一期邮件，周刊一般每周发送一期邮件。</li> " & vbCrLf
    Response.Write "<li>订阅服务是免费的，其内容由<a href='../' target='_blank'>本站</a>为您提供。</li> " & vbCrLf
    Response.Write "</ul></div> " & vbCrLf
    Response.Write "<div class='operation'> " & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveAdd'>" & vbCrLf
    Response.Write "<input name='Input' type='submit'  title='' value='确定' align='middle' width='110' height='45' > " & vbCrLf
    Response.Write "<input name='Input' type='button'  title='' value='取消' align='middle' onclick='window.close();' width='110' height='45' > " & vbCrLf
    Response.Write "</div></div> " & vbCrLf
    Response.Write "</form>" & vbCrLf
    Response.Write "</div></body></html> " & vbCrLf
    rsUser.Close
    Set rsUser = Nothing
End Sub

Sub Savemail()
    Dim arrOrder, arrNoOrder, intTemp, rsAdd, rsDel, sqlContacter, rsContacter, arrUserAdd, arrUserDel, arrAddChannel, arrDelChannel, arrAdd, arrDel
    
    If IsValidID(Trim(Request("ChannelID"))) = True Then
        ChannelID = Trim(Request("ChannelID"))
    Else
        ChannelID = ""
    End If
    
    arrOrder = getOrderID(ChannelID)
    arrNoOrder = getNoOrderID(ChannelID)
    arrAdd = Split(arrOrder, ",")
    arrDel = Split(arrNoOrder, ",")

    sqlContacter = "select C.arrChannelID from PE_Contacter C inner join PE_User U on U.ContacterID=C.ContacterID where U.UserID = " & UserID
    Set rsContacter = Server.CreateObject("adodb.recordset")
    rsContacter.Open sqlContacter, Conn, 1, 3
    If rsContacter.bof And rsContacter.EOF Then

        rsContacter.Close
        Set rsContacter = Nothing
        Response.Write "<script   language='javascript'>alert(""找不到指定用户！"");location.href=""../"";</script>"
        Exit Sub
    Else
        rsContacter("arrChannelID") = ChannelID
        rsContacter.Update
    End If
    rsContacter.Close
    Set rsContacter = Nothing
    
    For intTemp = 0 To UBound(arrAdd)
        arrAddChannel = ""
        Set rsAdd = Server.CreateObject("adodb.recordset")
        rsAdd.Open "select * from PE_MailChannel Where ChannelID = " & PE_CLng(arrAdd(intTemp)), Conn, 1, 3
        If rsAdd.bof And rsAdd.EOF Then
            Response.Write "找不到指定的频道"
            Exit Sub
        ElseIf rsAdd("UserID") = "" Or IsNull(rsAdd("UserID")) Then
            arrAddChannel = PE_CLng(UserID)
        Else
            Dim i
            i = 0
            arrUserAdd = Split(rsAdd("UserID"), ",")
            Do While i <> UBound(arrUserAdd) + 1
                If PE_CLng(arrUserAdd(i)) <> PE_CLng(UserID) Then
                    arrAddChannel = arrAddChannel & arrUserAdd(i) & ","
                End If
                i = i + 1
            Loop
            arrAddChannel = arrAddChannel & PE_CLng(UserID)
        End If
        rsAdd("UserID") = arrAddChannel
        rsAdd.Update
        rsAdd.Close
        Set rsAdd = Nothing
    Next

    For intTemp = 0 To UBound(arrDel)
        Set rsDel = Server.CreateObject("adodb.recordset")
        rsDel.Open "select * from PE_MailChannel Where ChannelID = " & PE_CLng(arrDel(intTemp)), Conn, 1, 3
        If rsDel.bof And rsDel.EOF Then
    Response.Write "<script   language='javascript'>alert(""找不到指定频道！"");location.href=""../"";</script>"
            Exit Sub
        ElseIf IsValidID(rsDel("UserID")) = False Then
            Response.Write "<script   language='javascript'>alert(""找不到指定用户！"");location.href=""../"";</script>"
             Exit Sub
        Else
            i = 0
            arrDelChannel = ""
            arrUserDel = Split(rsDel("UserID"), ",")
                Do While i <> UBound(arrUserDel) + 1
                    If arrDelChannel = "" And PE_CLng(arrUserDel(i)) <> PE_CLng(UserID) Then
                        arrDelChannel = arrUserDel(i)
                   ElseIf arrDelChannel <> "" And PE_CLng(arrUserDel(i)) <> PE_CLng(UserID) Then
                        arrDelChannel = arrDelChannel & "," & arrUserDel(i)
                    End If
                    i = i + 1
                Loop
        End If
        rsDel("UserID") = arrDelChannel
        rsDel.Update
        rsDel.Close
        Set rsDel = Nothing
    Next

    Response.Write "<script   language='javascript'>alert(""订阅成功！"");location.href=""../user"";</script>"
End Sub

    
Function getOrderID(ChannelID)
    Dim rsOrdered, arrChannelID, intTemp, arrOrderID
    Set rsOrdered = Conn.Execute("select C.arrChannelID from PE_Contacter C inner join PE_User U on U.ContacterID=C.ContacterID where U.UserID = " & UserID)
    If rsOrdered.bof And rsOrdered.EOF Then

        rsOrdered.Close
        Set rsOrdered = Nothing
            Response.Write "<script   language='javascript'>alert(""找不到指定用户！"");location.href=""../"";</script>"
        Exit Function
    ElseIf rsOrdered(0) = "" Or IsNull(rsOrdered(0)) Then
        arrOrderID = ChannelID
    Else
        arrChannelID = Split(ChannelID, ",")
        For intTemp = 0 To UBound(arrChannelID)
            If FoundInArr(rsOrdered(0), arrChannelID(intTemp), ",") = False Then
                If arrOrderID = "" Then
                    arrOrderID = arrChannelID(intTemp)
                Else
                    arrOrderID = arrOrderID & "," & arrChannelID(intTemp)
                End If
            Else
            
            End If
        Next
    End If
    rsOrdered.Close
    Set rsOrdered = Nothing
    getOrderID = arrOrderID
End Function

Function getNoOrderID(ChannelID)
    Dim rsOrdered, arrOrderedID, intTemp, arrNoOrderID
    Set rsOrdered = Conn.Execute("select C.arrChannelID from PE_Contacter C inner join PE_User U on U.ContacterID=C.ContacterID where U.UserID = " & UserID)
    If rsOrdered.bof And rsOrdered.EOF Then
            Response.Write "<script   language='javascript'>alert(""找不到指定用户！"");location.href=""../"";</script>"
        rsOrdered.Close
        Set rsOrdered = Nothing
        Exit Function
    ElseIf rsOrdered(0) = "" Or IsNull(rsOrdered(0)) Then
        arrNoOrderID = ""
    Else
        arrOrderedID = Split(rsOrdered(0), ",")
        For intTemp = 0 To UBound(arrOrderedID)
            If FoundInArr(ChannelID, arrOrderedID(intTemp), ",") = False Then
                If arrNoOrderID = "" Then
                    arrNoOrderID = arrOrderedID(intTemp)
                Else
                    arrNoOrderID = arrNoOrderID & "," & arrOrderedID(intTemp)
                End If
            Else
            
            End If
        Next
    End If
    rsOrdered.Close
    Set rsOrdered = Nothing
    getNoOrderID = arrNoOrderID
End Function
%>

