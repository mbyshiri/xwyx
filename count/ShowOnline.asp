<%@language=vbscript codepage=936 %>
<%
Option Explicit
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005��2009 ��ɽ�ж�������Ƽ����޹�˾ ��Ȩ����
'**************************************************************

Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!--#include file="Conn_Counter.asp"-->
<!--#include file="../Include/PowerEasy.Common.Security.asp"-->

<%
Dim FileName, strFileName, MaxPerPage, CurrentPage, totalPut
Dim strGuide, TitleRight, sql, rs
Dim OnNowTime, OnlineTime

MaxPerPage = PE_CLng(Trim(Request("MaxPerPage")))
If MaxPerPage <= 0 Then MaxPerPage = 20
CurrentPage = PE_CLng(Trim(Request("Page")))
FileName = "ShowOnline.asp"
If MaxPerPage > 0 Then strFileName = FileName & "?MaxPerPage=" & MaxPerPage


Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>��վͳ����ʾ��ǰ�����û�</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
%>
<style type="text/css">
<!--
/* ��վ�����ܵ�CSS����:�ɶ�������Ϊ����������ɫ����ʽ�� */
a {text-decoration: none;} /* �������»���,��Ϊunderline */
a:link {color: #000000;text-decoration: none;} /* δ���ʵ����� */
a:visited {color: #000000;text-decoration: none;} /* �ѷ��ʵ����� */
a:hover {color: #ff6600;text-decoration: none;} /* ����������� */
a:active {color: #000000;text-decoration: none;} /* ����������� */

Body
{
    FONT-FAMILY: "����";
    FONT-SIZE: 9pt;
    text-decoration: none;
    line-height: 150%;
    background:#ffffff;
    SCROLLBAR-FACE-COLOR: #2B73F1;
    SCROLLBAR-HIGHLIGHT-COLOR: #0650D2;
    SCROLLBAR-SHADOW-COLOR: #449AE8;
    SCROLLBAR-3DLIGHT-COLOR: #449AE8;
    SCROLLBAR-ARROW-COLOR: #02338A;
    SCROLLBAR-TRACK-COLOR: #0650D2;
    SCROLLBAR-DARKSHADOW-COLOR: #0650D2;
    margin-top: 0px;
    margin-right: 0px;
    margin-bottom: 0px;
    margin-left: 0px;
}

TD
{
FONT-FAMILY:����;FONT-SIZE: 9pt;line-height: 150%;
}

Input
{
    FONT-SIZE: 12px;
    HEIGHT: 20px;
}
Button
{
    FONT-SIZE: 9pt;
    HEIGHT: 20px;
}
Select
{
    FONT-SIZE: 9pt;
    HEIGHT: 20px;
}
.Title
{
    background:#449AE8;
    color: #ffffff;
    font-weight: normal;
}
.border
{
    border: 1px solid #449AE8;
}

.tdbg{
    background:#f0f0f0;
    line-height: 120%;
}

-->
</style>
<%
Response.Write "</head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"
Call OpenConn_Counter
If IsEmpty(Application("OnlineTime")) Then
    Set rs = Conn_Counter.Execute("select * from PE_StatInfoList")
    If Not rs.BOF And Not rs.EOF Then
        OnlineTime = rs("OnlineTime")
        Application("OnlineTime") = OnlineTime
    End If
    Set rs = Nothing
Else
    OnlineTime = Application("OnlineTime")
End If
OnNowTime = DateAdd("s", -OnlineTime, Now())
strGuide = "��ǰ�����û�����"

If CountDatabaseType = "SQL" Then
    sql = "select * from PE_StatOnline where LastTime>'" & OnNowTime & "' order by OnTime desc"
Else
    sql = "select * from PE_StatOnline where LastTime>#" & OnNowTime & "# order by OnTime desc"
End If

Set rs = Server.CreateObject("adodb.recordset")
rs.Open sql, Conn_Counter, 1, 1
If rs.BOF And rs.EOF Then
    Response.Write "<li>��ǰ�������ߣ�"
Else
    totalPut = rs.RecordCount
    TitleRight = TitleRight & "�� <font color=red>" & totalPut & "</font> ���û�����"
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
            rs.Move (CurrentPage - 1) * MaxPerPage
        Else
            CurrentPage = 1
        End If
    End If
    
    Dim VisitorNum, LNowTime
    VisitorNum = 0

    Response.Write "<br><table width='760' align='center'><tr><td align='left'>�����ڵ�λ�ã���վͳ��&nbsp;&gt;&gt;&nbsp;" & strGuide & "</td><td align='right'>" & TitleRight & "</td></tr></table>"
    Response.Write "<table width='760' border='0' cellspacing='1' cellpadding='2' class='border' align='center'>"
    Response.Write "  <tr class=title>"
    Response.Write "    <td align=center nowrap height='22'>���</td>"
    Response.Write "    <td align=center nowrap>������IP</td>"
    Response.Write "    <td align=center nowrap>��վʱ��</td>"
    Response.Write "    <td align=center nowrap>���ˢ��ʱ��</td>"
    Response.Write "    <td align=center nowrap>��ͣ��ʱ��</td>"
    Response.Write "    <td align=center nowrap>����ҳ�� �� �ͻ�����Ϣ</td>"
    Response.Write "  </tr>"
    
    Do While Not rs.EOF
        LNowTime = Cstrtime(CDate(Now() - rs("Ontime")))
        Response.Write "  <tr class='tdbg'>"
        Response.Write "    <td align=center width='50' nowrap>" & VisitorNum & "</td>"
        Response.Write "    <td align=left width='100' nowrap>" & rs("UserIP") & "</td>"
        Response.Write "    <td align=left width='110' nowrap><a title=" & rs("OnTime") & ">" & TimeValue(rs("OnTime")) & "</a></td>"
        Response.Write "    <td align=left width='100' nowrap>" & TimeValue(rs("LastTime")) & "</td>"
        Response.Write "    <td align=left width='100' nowrap>" & LNowTime & "</td>"
        Response.Write "    <td align=left width='300' nowrap title='����ҳ��: " & rs("UserPage") & vbCrLf & "�ͻ�����Ϣ: " & rs("UserAgent") & "'><a href=" & rs("UserPage") & " target=""_blank"">" & Left(Findpages(rs("UserPage")), 35) & "</a>"
        Response.Write "    </td>"
        Response.Write "  </tr>"
        VisitorNum = VisitorNum + 1
        If VisitorNum >= MaxPerPage Then Exit Do
        rs.MoveNext
    Loop
    Response.Write "</table>"
    If totalPut > 0 Then
        Response.Write ShowPage(strFileName, totalPut, MaxPerPage, CurrentPage, True, True, "�������û�", False)
    End If
End If
rs.Close
Set rs = Nothing
Call CloseConn_Counter


Function Cstrtime(Lsttime)
    Dim Dminute, Dsecond
    Cstrtime = ""
    Dminute = 60 * Hour(Lsttime) + Minute(Lsttime)
    Dsecond = Second(Lsttime)
    If Dminute <> 0 Then Cstrtime = Dminute & "'"
    If Dsecond < 10 Then Cstrtime = Cstrtime & "0"
    Cstrtime = Cstrtime & Dsecond & """"
End Function

Function Findpages(furl)
    Dim Ffurl
    If furl & "" <> "" Then
        Ffurl = Split(furl, "/")
        Findpages = Replace(furl, Ffurl(0) & "//" & Ffurl(2), "")
        If Findpages = "" Then Findpages = "/"
    Else
        Findpages = ""
    End If
End Function

Function ShowPage(sfilename, totalnumber, MaxPerPage, CurrentPage, ShowTotal, ShowAllPages, strUnit, ShowMaxPerPage)
    Dim TotalPage, strTemp, strUrl, i

    If totalnumber = 0 Or MaxPerPage = 0 Or IsNull(MaxPerPage) Then
        ShowPage = ""
        Exit Function
    End If
    If totalnumber Mod MaxPerPage = 0 Then
        TotalPage = totalnumber \ MaxPerPage
    Else
        TotalPage = totalnumber \ MaxPerPage + 1
    End If
    If CurrentPage > TotalPage Then CurrentPage = TotalPage
        
    strTemp = "<table align='center'><tr><td>"
    If ShowTotal = True Then
        strTemp = strTemp & "�� <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
    End If
    If ShowMaxPerPage = True Then
        strUrl = JoinChar(sfilename) & "MaxPerPage=" & MaxPerPage & "&"
    Else
        strUrl = JoinChar(sfilename)
    End If
    If CurrentPage = 1 Then
        strTemp = strTemp & "��ҳ ��һҳ&nbsp;"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a>&nbsp;"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage - 1) & "'>��һҳ</a>&nbsp;"
    End If

    If CurrentPage >= TotalPage Then
        strTemp = strTemp & "��һҳ βҳ"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage + 1) & "'>��һҳ</a>&nbsp;"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & TotalPage & "'>βҳ</a>"
    End If
    strTemp = strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & TotalPage & "</strong>ҳ "
    If ShowMaxPerPage = True Then
        strTemp = strTemp & "&nbsp;<input type='text' name='MaxPerPage' size='3' maxlength='4' value='" & MaxPerPage & "' onKeyPress=""if (event.keyCode==13) window.location='" & JoinChar(sfilename) & "page=" & CurrentPage & "&MaxPerPage=" & "'+this.value;"">" & strUnit & "/ҳ"
    Else
        strTemp = strTemp & "&nbsp;<b>" & MaxPerPage & "</b>" & strUnit & "/ҳ"
    End If
    If ShowAllPages = True Then
        If TotalPage > 20 Then
            strTemp = strTemp & "&nbsp;&nbsp;ת����<input type='text' name='page' size='3' maxlength='5' value='" & CurrentPage & "' onKeyPress=""if (event.keyCode==13) window.location='" & strUrl & "page=" & "'+this.value;"">ҳ"
        Else
            strTemp = strTemp & "&nbsp;ת����<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"
            For i = 1 To TotalPage
               strTemp = strTemp & "<option value='" & i & "'"
               If PE_CLng(CurrentPage) = PE_CLng(i) Then strTemp = strTemp & " selected "
               strTemp = strTemp & ">��" & i & "ҳ</option>"
            Next
            strTemp = strTemp & "</select>"
        End If
    End If
    strTemp = strTemp & "</td></tr></table>"
    ShowPage = strTemp
End Function

Function JoinChar(ByVal strUrl)
    If strUrl = "" Then
        JoinChar = ""
        Exit Function
    End If
    If InStr(strUrl, "?") < Len(strUrl) Then
        If InStr(strUrl, "?") > 1 Then
            If InStr(strUrl, "&") < Len(strUrl) Then
                JoinChar = strUrl & "&"
            Else
                JoinChar = strUrl
            End If
        Else
            JoinChar = strUrl & "?"
        End If
    Else
        JoinChar = strUrl
    End If
End Function
%>
