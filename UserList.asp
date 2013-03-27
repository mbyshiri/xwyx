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
PageTitle = "会员列表"
strFileName = "UserList.asp?ChannelID=" & ChannelID
Dim OrderType,Querysql
OrderType = Trim(Request("OrderType"))
If OrderType = "" Then
    OrderType = 3
Else
    OrderType = CLng(OrderType)
End If
strNavPath = strNavPath & strNavLink & "&nbsp;" & PageTitle

strHtml = GetTemplate(0, 9, 0)
Call ReplaceCommonLabel

strHtml = Replace(strHtml, "{$PageTitle}", SiteTitle & " >> " & PageTitle)
strHtml = Replace(strHtml, "{$ShowPath}", strNavPath)
strHtml = Replace(strHtml, "{$MenuJS}", GetMenuJS("", False))
strHtml = Replace(strHtml, "{$Skin_CSS}", GetSkin_CSS(0))

strHtml = Replace(strHtml, "{$ShowUserList}", GetUserList())
If InStr(strHtml, "{$ShowPage}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage}", ShowPage(strFileName & "&OrderType=" & OrderType, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("ShowSource", "ShowUserList/PageChar", "个会员"), False))
If InStr(strHtml, "{$ShowPage_en}") > 0 Then strHtml = Replace(strHtml, "{$ShowPage_en}", ShowPage_en(strFileName & "&OrderType=" & OrderType, totalPut, MaxPerPage, CurrentPage, True, True, XmlText("ShowSource", "ShowUserList/PageChar", "个会员"), False))
Response.Write strHtml
Call CloseConn

Function GetUserList()
    Dim sqlUser, rsUser, strUserList, i
    sqlUser = "select top " & MaxPerPage & " U.UserID,U.UserName,U.UserType,U.UserFace,U.Sign,U.Privacy,U.PostItems,U.RegTime,U.PassedItems,C.Sex, C.ZipCode,C.Fax,C.OfficePhone,C.HomePhone,C.Address,C.Department,C.Company,C.TrueName,C.QQ,C.ICQ,C.MSN,C.Email,C.HomePage,C.Birthday from PE_User U left join PE_Contacter C on U.ContacterID=C.ContacterID where 1=1  "
    If CurrentPage > 1 Then
        Select Case OrderType
        Case 1
            Querysql = " and (U.PassedItems<(select min(PassedItems) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " U.PassedItems from PE_User U  order by U.PassedItems desc,UserID desc) as QueryUser) Or (U.PassedItems=(select min(PassedItems) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " U.PassedItems from PE_User U  order by U.PassedItems desc,UserID desc) as QueryUser) and U.UserID<(select top 1 UserID from (select top " & ((CurrentPage - 1) * MaxPerPage) & " U.PassedItems,U.UserID from PE_User U  order by U.PassedItems desc,UserID desc)  as QueryUserID where PassedItems = (select min(PassedItems) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " U.PassedItems from PE_User U  order by U.PassedItems desc,UserID desc) as QueryUser) order by UserID)))"
         Case 2
             Querysql = "and U.RegTime < (select min(RegTime) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " U.RegTime from PE_User U   order by U.RegTime desc) as QueryUser)"
        Case 3
             Querysql = "and U.UserID < (select min(UserID) from (select top " & ((CurrentPage - 1) * MaxPerPage) & " U.UserID from PE_User U   order by U.UserID desc) as QueryUser)"
        End Select
    End If		
    sqlUser = sqlUser & Querysql
    Select Case OrderType
    Case 1
         sqlUser = sqlUser & "order by U.PassedItems desc,UserID desc"
    Case 2
         sqlUser = sqlUser & " order by U.RegTime desc"
    Case 3
         sqlUser = sqlUser & " order by U.UserID desc"
    End Select	
    Select Case OrderType
    Case 1
        totalPut = PE_CLng(Conn.Execute("select Count(*) from PE_User U")(0))
    Case 2,3
            totalPut = PE_CLng(Conn.Execute("select Count(*) from PE_User U")(0))
    End Select	
    Set rsUser = Server.CreateObject("adodb.recordset")
    rsUser.Open sqlUser, Conn, 1, 1
    If rsUser.BOF And rsUser.EOF Then
        totalPut = 0
        strUserList = strUserList & "<li>" & XmlText("ShowSource", "ShowUserList/NoFoundUser", "没有任何会员") & "</li>"
    Else
        i = 0
        Dim UserHomepage, UserQQ, UserICQ, UserMSN, UserSex, UserEmail, UserRegTime, UserPass
        strUserList = strUserList & "<table width='100%' border='0' cellspacing='1' cellpadding='3' class='Channel_border'>"
        strUserList = strUserList & "<tr align='center' class='Channel_pager'><td colspan='8'>" & Replace(XmlText("ShowSource", "ShowUserList/Order", "<a href='{$strFileName}&OrderType=1'>按发表文章数排序</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='{$strFileName}&OrderType=2'>按注册日期排序</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='{$strFileName}&OrderType=3'>按会员ID排序</a>"), "{$strFileName}", strFileName) & "</td></tr>"
        strUserList = strUserList & XmlText("ShowSource", "ShowUserList/t1", "<tr align='center' class='Channel_title'><td>会员昵称</td><td>性别</td><td>Email</td><td>QQ号码</td><td>MSN</td><td>主页</td><td>注册日期</td><td>文章数</td><tr>")
        
   
        Do While Not rsUser.EOF
            If rsUser("Privacy") = 2 Then
                UserHomepage = Secrit
                UserQQ = Secrit
                UserICQ = Secrit
                UserMSN = Secrit
                UserSex = Secrit
                UserEmail = Secrit
                UserRegTime = Secrit
            Else
                If rsUser("Homepage") = "" Then
                    UserHomepage = NoEnter
                Else
                    UserHomepage = "<a href='" & rsUser("Homepage") & "'>" & rsUser("Homepage") & "</a>"
                End If
                If rsUser("QQ") = "" Then
                    UserQQ = NoEnter
                Else
                    UserQQ = rsUser("QQ")
                End If
                If rsUser("ICQ") = "" Then
                    UserICQ = NoEnter
                Else
                    UserICQ = rsUser("ICQ")
                End If
                If rsUser("MSN") = "" Then
                    UserMSN = NoEnter
                Else
                    UserMSN = rsUser("MSN")
                End If
                If rsUser("Sex") = 2 Then
                    UserSex = strGirl
                ElseIf rsUser("Sex") = 1 Then
                    UserSex = strMan
                Else
                    UserSex = Secrit
                End If
                If rsUser("Email") = "" Then
                    UserEmail = NoEnter
                Else
                    UserEmail = "<a href=""mailto:" & rsUser("Email") & """>" & rsUser("Email") & "</a>"
                End If
                If rsUser("RegTime") = "" Then
                    UserRegTime = NoEnter
                Else
                    UserRegTime = Year(rsUser("RegTime")) & strYear & Month(rsUser("RegTime")) & strMonth & Day(rsUser("RegTime")) & strDay
                End If
            End If
            strUserList = strUserList & ("<tr class='Channel_tdbg'><td><a href='ShowUser.asp?UserID=" & rsUser("UserID") & "'>" & rsUser("UserName") & "</a></td>")
            strUserList = strUserList & ("<td align='center'>" & UserSex & "</td><td>" & UserEmail & "</td><td align='center'>" & UserQQ & "</td>")
            strUserList = strUserList & ("<td align='center'>" & UserMSN & "</td><td align='center'>" & UserHomepage & "</td><td align='center'>" & UserRegTime & "</td><td align='right'>" & rsUser("PassedItems") & "</td></tr>")
            rsUser.MoveNext
            i = i + 1
            If i >= MaxPerPage Then Exit Do
        Loop
        strUserList = strUserList & "</table>"
        '完毕
    End If
    rsUser.Close
    Set rsUser = Nothing
    GetUserList = strUserList
End Function
%>
