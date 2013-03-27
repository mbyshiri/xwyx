<!--#Include file="Start.asp"-->
<%
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
'**************************************************************

Dim Country, Province, City
Dim arrCountry, arrProvince, arrCity
Dim rsRegion

Country = ReplaceBadChar(Trim(Request("Country")))
Province = ReplaceBadChar(Trim(Request("Province")))
City = ReplaceBadChar(Trim(Request("City")))
If Country = "" Then Country = "中华人民共和国"
If Province = "" Then Province = "北京市"

Set rsRegion = Conn.Execute("SELECT Country FROM PE_Country ORDER BY CountryID")
If rsRegion.BOF And rsRegion.EOF Then
    FoundErr = True
Else
    arrCountry = rsRegion.GetRows(-1)
End If
Set rsRegion = Nothing
Set rsRegion = Conn.Execute("SELECT Province FROM PE_Province WHERE Country='" & Country & "' ORDER BY ProvinceID")
If rsRegion.BOF And rsRegion.EOF Then
    ReDim arrProvince(0, 0)
Else
    arrProvince = rsRegion.GetRows(-1)
End If
Set rsRegion = Conn.Execute("SELECT DISTINCT City FROM PE_City WHERE Country='" & Country & "' And Province='" & Province & "'")
If rsRegion.BOF And rsRegion.EOF Then
    ReDim arrCity(0, 0)
Else
    arrCity = rsRegion.GetRows(-1)
End If
Set rsRegion = Nothing
Call CloseConn
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Images/Admin_Style.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="100%"  border="0" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
    <form name="regionform" id="regionform" action="Region.asp" method="post">
    <tr class="tdbg">
        <td width="100" align="right" class="tdbg5">
            国家/地区：
        </td>
        <td colspan="2">
            <select name="Country" id="Country" onChange="document.regionform.submit();">
                <%
                Dim i
                i = 0
                For i = 0 To UBound(arrCountry, 2)
                    Response.Write "<option value='" & arrCountry(0, i) & "'"
                    If Country = arrCountry(0, i) Then Response.Write " selected"
                    Response.Write ">" & arrCountry(0, i) & "</option>" & vbCrLf
                Next
                %>
            </select>
        </td>
    </tr>
    <tr class="tdbg">
        <td align="right" class="tdbg5">
            省/市/自治区：
        </td>
        <td>
            <select name="Province" id="Province" onChange="document.regionform.submit();">
                <%
                If arrProvince(0, 0) = "" Then
                    Response.Write "<option>    </option>" & vbCrLf
                Else
                    i = 0
                    For i = 0 To UBound(arrProvince, 2)
                        Response.Write "<option value='" & arrProvince(0, i) & "'"
                        If Province = arrProvince(0, i) Then Response.Write " selected"
                        Response.Write ">" & arrProvince(0, i) & "</option>" & vbCrLf
                    Next
                End If
                %>
            </select>
        </td>
    </tr>
    <tr class="tdbg">
        <td align="right" class="tdbg5">
            市/县/区/旗：
        </td>
        <td>
            <select name="City" id="City">
                <%
                If arrCity(0, 0) = "" Then
                    Response.Write "<option>    </option>" & vbCrLf
                Else
                    i = 0
                    For i = 0 To UBound(arrCity, 2)
                        Response.Write "<option value='" & arrCity(0, i) & "'"
                        If City = arrCity(0, i) Then Response.Write " selected"
                        Response.Write ">" & arrCity(0, i) & "</option>" & vbCrLf
                    Next
                End If
                %>
            </select>
        </td>
    </tr>
    </form>
</table>
</body>
</html>

