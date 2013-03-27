<!-- #include File="../Start.asp" -->
<!--#include file="../Include/PowerEasy.Common.Security.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Response.ContentType = "text/Html; charset=gb2312"
Response.Expires = -1
Dim Received, rootNode, itext, ilnum, ichannelid, itype, iname
Set Received = CreateObject("Microsoft.XMLDOM")
Received.async = False
Received.Load Request

Set rootNode = Received.getElementsByTagName("root")
If rootNode.length > 0 Then
    itext = rootNode(0).selectSingleNode("text").text
    If itext <> "" Then
        itext = ReplaceBadChar(itext)
    Else
        Set Received = Nothing
        Response.End
    End If

    iname = rootNode(0).selectSingleNode("inputname").text

    If itext = "ChongFuUserCheck" Then
        If iname <> "" Then
'            iname = ReplaceBadChar(iname)
            Call usercheck
        Else
            Response.write "0"
        End If
    Else
        ilnum = rootNode(0).selectSingleNode("lnum").text
        ichannelid = rootNode(0).selectSingleNode("channelid").text
        itype = rootNode(0).selectSingleNode("type").text
        If ilnum = "" Or ilnum < 1 Then
            ilnum = 10
        Else
            ilnum = CLng(ilnum)
        End If
        If ichannelid = "" Or ichannelid < 1 Then
            ichannelid = 0
        Else
            ichannelid = CLng(ichannelid)
        End If
        Call outitem
    End If
End If
Set Received = Nothing

Sub outitem()
    Dim rtext, qsql
    Select Case itype
    Case "satitle"
        qsql = "select top " & ilnum & " Title,UpdateTime from PE_Article where Title like '" & itext & "%'"
        If ichannelid > 0 Then qsql = qsql & " and ChannelID=" & ichannelid
        qsql = qsql & " and Deleted=" & PE_False & " order by UpdateTime desc"
    Case "satitle2"
        qsql = "select top " & ilnum & " PhotoName,UpdateTime from PE_Photo where PhotoName like '" & itext & "%'"
        If ichannelid > 0 Then qsql = qsql & " and ChannelID=" & ichannelid
        qsql = qsql & " and Deleted=" & PE_False & " order by UpdateTime desc"
    Case "satitle3"
        qsql = "select top " & ilnum & " SoftName,UpdateTime from PE_Soft where SoftName like '" & itext & "%'"
        If ichannelid > 0 Then qsql = qsql & " and ChannelID=" & ichannelid
        qsql = qsql & " and Deleted=" & PE_False & " order by UpdateTime desc"
    Case "satitle4"
        qsql = "select top " & ilnum & " ProductName,UpdateTime from PE_Product where ProductName like '" & itext & "%'"
        If ichannelid > 0 Then qsql = qsql & " and ChannelID=" & ichannelid
        qsql = qsql & " and Deleted=" & PE_False & " order by UpdateTime desc"
    Case "skey"
        qsql = "select top " & ilnum & " KeyText,LastUseTime from PE_NewKeys where KeyText like '" & itext & "%'"
        If ichannelid > 0 Then qsql = qsql & " and ChannelID=" & ichannelid
        qsql = qsql & " order by LastUseTime desc"
    Case "sauthor", "sauthor1"
        qsql = "select top " & ilnum & " AuthorName,LastUseTime from PE_Author where AuthorName like '" & itext & "%' and ChannelID=0"
        If ichannelid > 0 Then qsql = qsql & " or ChannelID=" & ichannelid
        qsql = qsql & " and Passed="&PE_True &" order by LastUseTime desc"
    Case "scopyfrom", "scopyfrom1"
        qsql = "select top " & ilnum & " SourceName,LastUseTime from PE_CopyFrom where SourceName like '" & itext & "%' and ChannelID=0"
        If ichannelid > 0 Then qsql = qsql & " or ChannelID=" & ichannelid
        qsql = qsql & " and Passed="&PE_True &" order by LastUseTime desc"
    End Select
    If qsql <> "" Then
        Set rtext = Conn.Execute(qsql)
        Do While Not rtext.EOF
            Response.write "<li style=""cursor:hand;"" onclick=""addinput('" & iname & "','" & rtext(0) & "');"">" & rtext(0) & "</li>"
            rtext.movenext
        Loop
        Set rtext = Nothing
    End If
End Sub

Sub usercheck()
    Dim rtext

    If CheckUserBadChar(iname) = False Then
        Response.write "2" '含有非法字符
    Else
        iname = ReplaceBadChar(iname)
        If GetStrLen(iname) > UserNameMax  Then
            Response.write "3" '长度太长
        Elseif  GetStrLen(iname) < UserNameLimit Then
            Response.write "5" '长度太短	
        Else
            If FoundInArr(UserName_RegDisabled, iname, "|") = True Then
                Response.write "4" '禁止注册
            Else
                Set rtext = Conn.Execute("select top 1 UserName from PE_User where UserName='" & iname & "'")
                If rtext.bof And rtext.EOF Then
                    Response.write "0"
                Else
                    Response.write "1" '重复
                End If
                Set rtext = Nothing
            End If
        End If
    End If		
End Sub


%>
