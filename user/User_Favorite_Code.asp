<!--#include file="CommonCode.asp"-->
<!--#include file="../Include/PowerEasy.Article.asp"-->
<!--#include file="../Include/PowerEasy.Soft.asp"-->
<!--#include file="../Include/PowerEasy.Photo.asp"-->
<!--#include file="../Include/PowerEasy.Product.asp"-->
<!--#include file="../Include/PowerEasy.Supply.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Sub Execute()
    strFileName = "User_Favorite.asp"
    Dim InfoID, trs
    ChannelID = PE_CLng(Trim(Request("ChannelID")))
    If ChannelID > 0 Then
        Call GetChannel(ChannelID)
    'Else
    '   FoundErr = True
    '   ErrMsg = ErrMsg & "<li>请指定要查看的频道ID！</li>"
    '   Response.Write ErrMsg
    '   Exit Sub
    End If
    InfoID = Trim(Request("InfoID"))

    If IsValidID(InfoID) = False Then
        InfoID = ""
    End If

    Select Case Action
    Case "Add"
        InfoID = PE_CLng(Trim(Request("InfoID")))
        If MaxFavorite <= 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>对不起，您没有收藏夹的权限！</li>"
        Else
            Set trs = Conn.Execute("select count(InfoID) from PE_Favorite where UserID=" & UserID & "")
            If trs(0) >= MaxFavorite Then
                FoundErr = True
                ErrMsg = ErrMsg & "<li>您的收藏夹已经到达上限！</li>"
            End If
            Set trs = Nothing
        End If
        If FoundErr = True Then
            Call WriteErrMsg(ErrMsg, "")
            Exit Sub
        End If
        If InfoID > 0 Then
            Set trs = Conn.Execute("select InfoID from PE_Favorite where ChannelID=" & ChannelID & " and UserID=" & UserID & " and InfoID=" & InfoID & "")
            If trs.BOF And trs.EOF Then
                Conn.Execute ("insert into PE_Favorite (ChannelID,UserID,InfoID,DateAndTime) values (" & ChannelID & "," & UserID & "," & InfoID & "," & PE_Now & ")")
            End If
        End If
    Case "Remove"
        If InfoID = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>请指定要批量操作的ID</li>"
            Call WriteErrMsg(ErrMsg, "")
            Exit Sub
        End If
        If InStr(InfoID, ",") > 0 Then
            Conn.Execute ("delete from PE_Favorite where ChannelID=" & ChannelID & " and UserID=" & UserID & " and InfoID in (" & InfoID & ")")
        Else
            Conn.Execute ("delete from PE_Favorite where ChannelID=" & ChannelID & " and UserID=" & UserID & " and InfoID=" & InfoID & "")
        End If
    Case "Clear"
        Conn.Execute ("delete from PE_Favorite where ChannelID=" & ChannelID & " and UserID=" & UserID & "")
    End Select

    Response.Write "<SCRIPT language=javascript>" & vbCrLf
    Response.Write "function unselectall(){" & vbCrLf
    Response.Write "    if(document.myform.chkAll.checked){" & vbCrLf
    Response.Write " document.myform.chkAll.checked = document.myform.chkAll.checked&0;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function CheckAll(form){" & vbCrLf
    Response.Write "  for (var i=0;i<form.elements.length;i++){" & vbCrLf
    Response.Write "    var e = form.elements[i];" & vbCrLf
    Response.Write "    if (e.Name != 'chkAll'&&e.disabled!=true)" & vbCrLf
    Response.Write "       e.checked = form.chkAll.checked;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf

    Response.Write "function SelectAll(form){" & vbCrLf
    Response.Write "  for (var i=0;i<form.AdminPurview_Others.length;i++){" & vbCrLf
    Response.Write "    var e = form.AdminPurview_Others[i];" & vbCrLf
    Response.Write "    if (e.disabled==false)" & vbCrLf
    Response.Write "       e.checked = form.chkAll.checked;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

    Dim rsChannel
    Set rsChannel = Conn.Execute("select * from PE_Channel where ChannelType<=1 and ModuleType<>4 and ModuleType<>8 and Disabled=" & PE_False & " order by OrderID")
	Do While Not rsChannel.EOF
        If rsChannel("ChannelID") = ChannelID Then
            ChannelName = rsChannel("ChannelName")
            ChannelShortName = rsChannel("ChannelShortName")
            ModuleType = rsChannel("ModuleType")
            Response.Write "<a href='User_Favorite.asp?ChannelID=" & ChannelID & "'><font color='red'>" & ChannelName & "</font></a> | "
        Else
            Response.Write "<a href='User_Favorite.asp?ChannelID=" & rsChannel("ChannelID") & "'>" & rsChannel("ChannelName") & "</a> | "
        End If
        rsChannel.MoveNext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing

    Response.Write "  <form name='myform' method='Post' action='User_Favorite.asp'>"

    Dim PE_Content
    Select Case ModuleType
    Case 1
        Set PE_Content = New Article
    Case 2
        Set PE_Content = New Soft
    Case 3
        Set PE_Content = New Photo
    Case 5
        Set PE_Content = New Product
    Case 6
        Set PE_Content = New Supply
    Case Else
        Set PE_Content = New Article
    End Select
    Call PE_Content.Init
    Call PE_Content.ShowFavorite
    Set PE_Content = Nothing

    Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td width='200' height='30'><input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'> 选中本页显示的所有" & ChannelShortName & "</td>"
    Response.Write "    <td><input name='Action' type='hidden' id='Action' value='Remove'><input name='ChannelID' type='hidden' value='" & ChannelID & "'>"
    Response.Write "<input name='Submit' type='submit' id='Submit' value='取消收藏' onclick=""document.myform.Action.value='Remove';return confirm('确定不再收藏选中的信息吗？');"">&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "<input name='Submit' type='submit' id='Submit' value='清空收藏夹' onclick=""document.myform.Action.value='Clear';return confirm('确定要清空收藏夹吗？');"">"
    Response.Write "</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
End Sub




Sub ShowFavorite_Product()
End Sub


%>
