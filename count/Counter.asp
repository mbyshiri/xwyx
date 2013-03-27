<%@language=vbscript codepage=936 %>
<%
Option Explicit
'**************************************************************
' Software name: PowerEasy SiteWeaver
' Web: http://www.powereasy.net
' Copyright (C) 2005－2009 佛山市动易网络科技有限公司 版权所有
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

Dim Ip, LastIPCache, Sip, Address, Scope, Referer, VisitorKeyword, WebUrl, Visit, StatIP, strIP
Dim Agent, System, Browser, BcType, Mozilla, Height, Width, Screen, Color, Timezone, Ver, VisitTimezone
Dim StrYear, StrMonth, StrDay, StrHour, Strweek, StrHourLong, StrDayLong, StrMonthLong, OldDay
Dim Num, I, nYesterDayNum, CacheData
Dim Province, OnlineNum, ShowInfo
Dim OnNowTime, style
Dim RegCount_Fill, OnlineTime, VisitRecord, KillRefresh
Dim DayNum, AllNum, TotalView, StartDate, StatDayNum, AveDayNum

Call OpenConn_Counter

Dim Sql, Rs
Set Rs = Conn_Counter.Execute("select * from PE_StatInfoList")
If Not Rs.BOF And Not Rs.EOF Then
    RegCount_Fill = Rs("RegFields_Fill")
    OnlineTime = Rs("OnlineTime")
    VisitRecord = Rs("VisitRecord")
    KillRefresh = Rs("KillRefresh")
    DayNum = Rs("DayNum")
    AllNum = Rs("TotalNum") + Rs("OldTotalNum")
    TotalView = Rs("TotalView") + Rs("OldTotalView")
    StartDate = Rs("StartDate")
    StatDayNum = DateDiff("D", StartDate, Date) + 1
    If StatDayNum <= 0 Or IsNumeric(StatDayNum) = 0 Then
        AveDayNum = StatDayNum
    Else
        AveDayNum = CLng(AllNum / StatDayNum)
    End If
End If
Set Rs = Nothing

Response.Expires = 0
LastIPCache = "Powereasy_LastIP"
If IsEmpty(Application(LastIPCache)) Then Application(LastIPCache) = "#0.0.0.0#"

Ip = ReplaceBadChar(Request.ServerVariables("REMOTE_ADDR"))

If FoundInArr(RegCount_Fill, "IsCountOnline", ",") = True Then
    If OnlineTime = "" Or IsNumeric(OnlineTime) = 0 Then OnlineTime = 100
    OnNowTime = DateAdd("s", -OnlineTime, Now())
    Dim rsOnline
    If CountDatabaseType = "SQL" Then
        set rsonline = conn_counter.execute("select count(UserIP) from PE_Statonline where LastTime>'"&OnNowTime&"'")
    Else
        set rsonline = conn_counter.execute("select count(UserIP) from PE_Statonline where LastTime>#"&OnNowTime&"#")
    End If
    OnlineNum = rsOnline(0)     ' 当前在线人数
    Set rsOnline = Nothing
    If CountDatabaseType = "SQL" Then
        Set rsonline = conn_counter.execute("select LastTime,OnTime from PE_Statonline where LastTime>'"&OnNowTime&"' and UserIP='"&IP&"'")
    Else
        Set rsonline = conn_counter.execute("select LastTime,OnTime from PE_Statonline where LastTime>#"&OnNowTime&"# and UserIP='"&IP&"'")
    End If
    If rsOnline.EOF Then
        Update()
    Else
        If rsOnline(0) = rsOnline(1) Then
            Update()
        Else
            Conn_Counter.Execute ("Update PE_StatInfoList set TotalView=TotalView+1")
        End If
    End If
    Set rsOnline = Nothing
Else
    If InStr(Application(LastIPCache), "#" & Ip & "#") Then ' 如果IP已经存在于保存的列表中，是刷新
        Conn_Counter.Execute ("Update PE_StatInfoList set TotalView=TotalView+1")
    Else
        Application.Lock
        Application(LastIPCache) = SaveIP(Application(LastIPCache))     ' 更新最近需要防刷的IP
        Application.UnLock
        Update()
    End If
End If


style = LCase(Trim(Request("style")))
Select Case style
Case "simple"
    ShowInfo = "总访问量：" & AllNum & "人次<br>"
    If FoundInArr(RegCount_Fill, "IsCountOnline", ",") = True Then
        ShowInfo=ShowInfo&"当前在线：" & OnlineNum & "人"
    End If
Case "all"
    ShowInfo=ShowInfo&"总访问量：" & AllNum & "人次<br>"
    ShowInfo=ShowInfo&"总浏览量：" & TotalView & "人次<br>"
'   ShowInfo=ShowInfo&"统计天数：" & StatDayNum & "天<br>"
    If FoundInArr(RegCount_Fill, "FYesterDay", ",") = True Then
        Call GetYesterdayNum
        ShowInfo=ShowInfo&"昨日访问：" & nYesterDayNum & "人<br>"
    End If
    ShowInfo=ShowInfo&"今日访问：" & DayNum & "人次<br>"
    ShowInfo=ShowInfo&"日均访问：" & AveDayNum & "人次<br>"
    If FoundInArr(RegCount_Fill, "IsCountOnline", ",") = True Then
        ShowInfo=ShowInfo&"当前在线：" & OnlineNum & "人"
    End If
Case "common"
    ShowInfo = "总访问量：" & AllNum & "人次<br>"
    ShowInfo=ShowInfo&"总浏览量：" & TotalView & "人次<br>"
    If FoundInArr(RegCount_Fill, "IsCountOnline", ",") = True Then
        ShowInfo=ShowInfo&"当前在线：" & OnlineNum & "人"
    End If
End Select
If style <> "none" Then
    Response.Write "document.write(" & Chr(34) & ShowInfo & Chr(34) & ")"
End If

Call CloseConn_Counter
Sub Update()
    If FoundInArr(RegCount_Fill, "FIP", ",") = True Then
        strIP = Split(Ip, ".")
        If IsNumeric(strIP(0)) = 0 Or IsNumeric(strIP(1)) = 0 Or IsNumeric(strIP(2)) = 0 Or IsNumeric(strIP(3)) = 0 Then
            Sip = 0
        Else
            Sip = CInt(strIP(0)) * 256 * 256 * 256 + CInt(strIP(1)) * 256 * 256 + CInt(strIP(2)) * 256 + CInt(strIP(3)) - 1
        End If
        if (167772159 < Sip and Sip< 184549374) or (2886729727 < Sip and Sip < 2887778302) or (3232235519 < Sip and Sip < 3232301054) then
            StatIP = Ip
        Else
            StatIP = strIP(0) & "." & strIP(1) & ".*"
        End If
    Else
        StatIP = ""
    End If
    Sip = Ip
    Set Rs = server.CreateObject("adodb.recordset")
    If Sip = "127.0.0.1" Then
        Address = "本机地址"
        Scope = "ChinaNum"
    Else
        strIP = Split(Sip, ".")
        If IsNumeric(strIP(0)) = 0 Or IsNumeric(strIP(1)) = 0 Or IsNumeric(strIP(2)) = 0 Or IsNumeric(strIP(3)) = 0 Then
            Sip = 0
        Else
            Sip = CInt(strIP(0)) * 256 * 256 * 256 + CInt(strIP(1)) * 256 * 256 + CInt(strIP(2)) * 256 + CInt(strIP(3)) - 1
        End If

        Dim RsAdress
        set RsAdress=conn_counter.execute("Select Top 1 Address From PE_StatIpInfo Where StartIp<="&Sip&" and EndIp>="&Sip&" Order By EndIp-StartIp Asc")
        If RsAdress.EOF Then
            Address = "其它地区"
        Else
            Address = RsAdress(0)
        End If
        Set RsAdress = Nothing
        Province = "北京天津上海重庆黑龙江吉林辽宁江苏浙江安徽河南河北湖南湖北山东山西内蒙古陕西甘肃宁夏青海新疆西藏云南贵州四川广东广西福建江西海南香港澳门台湾内部网未知"
        If InStr(Province, Left(Address, 2)) > 0 Then
            Scope = "ChinaNum"
        Else
            Scope = "OtherNum"
        End If
    End If

    Referer = Request.QueryString("Referer")
    If Referer = "" Then Referer = "直接输入或书签导入"
    Referer = ReplaceUrlBadChar(Left(Referer, 100))

        'response.write"11="&Referer
        'response.end

    If FoundInArr(RegCount_Fill, "FWeburl", ",") = True Then
        WebUrl = Left(Request.QueryString("Referer"), InStr(8, Referer, "/"))
        If WebUrl = "" Then WebUrl = "直接输入或书签导入"
        WebUrl = ReplaceUrlBadChar(Left(WebUrl, 50))
    Else
        WebUrl = ""
    End If

    Width = ReplaceBadChar(Request.QueryString("Width"))
    Height = ReplaceBadChar(Request.QueryString("Height"))
    If Height = "" Or IsNumeric(Height) = 0 Or Width = "" Or IsNumeric(Width) = 0 Then
        Screen = "其它"
    Else
        Screen = CStr(Width) & "x" & CStr(Height)
    End If
    Screen = Left(Screen, 10)



    Color = ReplaceBadChar(Request.QueryString("Color"))
    If Color = "" Or IsNumeric(Color) = 0 Then
        Color = "其它"
    Else
        Select Case Color
        Case 4:
             Color = "16 色"
        Case 8:
             Color = "256 色"
        Case 16:
             Color = "增强色（16位）"
        Case 24:
             Color = "真彩色（24位）"
        Case 32:
             Color = "真彩色（32位）"
        End Select
    End If


    Mozilla = Replace(Request.ServerVariables("HTTP_USER_AGENT"), "'", "")
    Mozilla = Left(Mozilla, 100)
    Agent = Request.ServerVariables("HTTP_USER_AGENT")
    Agent = Split(Agent, ";")
    BcType = 0
    If InStr(Agent(1), "U") Or InStr(Agent(1), "I") Then BcType = 1
    If InStr(Agent(1), "MSIE") Then BcType = 2
    Select Case BcType
    Case 0:
         Browser = "其它"
         System = "其它"
    Case 1:
         Ver = Mid(Agent(0), InStr(Agent(0), "/") + 1)
         Ver = Mid(Ver, 1, InStr(Ver, " ") - 1)
         Browser = "Netscape" & Ver
         System = Mid(Agent(0), InStr(Agent(0), "(") + 1)
    Case 2:
         Browser = Agent(1)
         System = Agent(2)
         System = Replace(System, ")", "")
    End Select
    System = Replace(Replace(Replace(Replace(Replace(Replace(System, " ", ""), "Win", "Windows"), "NT5.0", "2000"), "NT5.1", "XP"), "NT5.2", "2003"), "dowsdows", "dows")
    Browser = Replace(Replace(Browser, " ", ""), "'", "")
    System = Replace(Left(System, 20), "'", "")

    Browser = Left(Browser, 20)

    Timezone = ReplaceBadChar(Request.QueryString("Timezone"))
    If Timezone = "" Or IsNumeric(Timezone) = 0 Then
       Timezone = "其它"
       VisitTimezone = 0
    Else
        VisitTimezone = Timezone \ 60
        If Timezone < 0 Then
            Timezone="GMT+"&Abs(Timezone)\60&":"&(Abs(Timezone) Mod 60)
        Else
            Timezone="GMT-"&Abs(Timezone)\60&":"&(Abs(Timezone) Mod 60)
        End If
    End If


    If FoundInArr(RegCount_Fill, "FVisit", ",") = True Then
        Visit = Request.Cookies("VisitNum")
        If Visit <> "" Then
            Visit = Visit + 1
        Else
            Visit = 1
        End If
        Response.Cookies("VisitNum") = Visit
        Response.Cookies("VisitNum").Expires = "January 01, 2010"
        Sql = "Select * From PE_StatVisit"
        Rs.Open Sql, Conn_Counter, 1, 3
        If Rs.EOF Or Rs.BOF Then
            Rs.AddNew
        End If
        If Visit <= 10 Then
            If IsNumeric(Rs(Visit - 1)) = 0 Then
                Rs(Visit - 1) = 1
            Else
                Rs(Visit - 1) = Rs(Visit - 1) + 1
                If Visit > 1 Then
                   If Rs(Visit - 2) > 0 Then Rs(Visit - 2) = Rs(Visit - 2) - 1
                End If
            End If
        End If
        Rs.Update
        Rs.Close
    End If

    Call UpdateVisit

    StrHour = CStr(Hour(Time))
    StrDay = CStr(Day(Date))
    StrMonth = CStr(Month(Date))
    StrYear = CStr(Year(Date))
    Strweek = CStr(Weekday(Date))
    StrDayLong = CStr(Year(Date) & "-" & Month(Date) & "-" & Day(Date))
    StrMonthLong = CStr(Year(Date) & "-" & Month(Date))
    StrHourLong=StrDayLong&" "&Cstr(Hour(Time))&":00:00"

    Sql = "Select * From PE_StatInfoList"
    Rs.Open Sql, Conn_Counter, 1, 3
    Rs("TotalNum") = Rs("TotalNum") + 1
    Rs("TotalView") = Rs("TotalView") + 1
    Rs(Scope) = Rs(Scope) + 1
    If IsNull(Rs("StartDate")) Then Rs("StartDate") = StrDayLong
    If IsNull(Rs("OldDay")) Then Rs("OldDay") = StrDayLong
    OldDay = Rs("OldDay")
    Rs.Update
    Rs.Close
    Call ModiMaxNum

    If VisitorKeyword <> "" And FoundInArr(RegCount_Fill, "FKeyword", ",") = True Then
        VisitorKeyword = FindKeystr(Request.QueryString("Referer"))
        VisitorKeyword = ReplaceBadChar(Trim(LCase(VisitorKeyword)))
        AddNum VisitorKeyword, "PE_Statkeyword", "Tkeyword", "TkeywordNum"
    End If
    If FoundInArr(RegCount_Fill, "FSystem", ",") = True Then
        AddNum System, "PE_StatSystem", "TSystem", "TSysNum"
    End If
    If FoundInArr(RegCount_Fill, "FBrowser", ",") = True Then
        AddNum Browser, "PE_StatBrowser", "TBrowser", "TBrwNum"
    End If
    If FoundInArr(RegCount_Fill, "FMozilla", ",") = True Then
        AddNum Mozilla, "PE_StatMozilla", "TMozilla", "TMozNum"
    End If
    If FoundInArr(RegCount_Fill, "FScreen", ",") = True Then
        AddNum Screen, "PE_StatScreen", "TScreen", "TScrNum"
    End If
    If FoundInArr(RegCount_Fill, "FColor", ",") = True Then
        AddNum Color, "PE_StatColor", "TColor", "TColNum"
    End If
    If FoundInArr(RegCount_Fill, "FTimezone", ",") = True Then
        AddNum Timezone, "PE_StatTimezone", "TTimezone", "TTimNum"
    End If
    If FoundInArr(RegCount_Fill, "FRefer", ",") = True Then
        AddNum Referer, "PE_StatRefer", "TRefer", "TRefNum"
    End If
    If FoundInArr(RegCount_Fill, "FWeburl", ",") = True Then
        AddNum WebUrl, "PE_StatWeburl", "TWeburl", "TWebNum"
    End If
    If FoundInArr(RegCount_Fill, "FAddress", ",") = True Then
        AddNum Address, "PE_StatAddress", "TAddress", "TAddNum"
    End If
    If FoundInArr(RegCount_Fill, "FIP", ",") = True Then
        AddNum Ip, "PE_StatIp", "TIp", "TIpNum"
    End If

    AddNum StrDayLong, "PE_StatDay", "TDay", StrHour
    AddNum "Total", "PE_StatDay", "TDay", StrHour
    AddNum StrYear, "PE_StatYear", "TYear", StrMonth
    AddNum "Total", "PE_StatYear", "TYear", StrMonth
    AddNum StrMonthLong, "PE_StatMonth", "TMonth", StrDay
    AddNum "Total", "PE_StatMonth", "TMonth", StrDay
    AddNum "Total", "PE_StatWeek", "TWeek", Strweek
    If DateDiff("Ww", CDate(OldDay), Date) > 0 Then
        Sql = "Delete From PE_StatWeek Where TWeek='Current'"
        Conn_Counter.Execute (Sql)
    End If
    AddNum "Current", "PE_StatWeek", "TWeek", Strweek
End Sub

Sub AddNum(Data, TableName, CompareField, AddField)
    Dim RowCount
    conn_counter.execute "update "&TableName&" set ["&AddField&"]=["&AddField&"]+1 where  "&CompareField&"='"&Data&"'", RowCount
    If RowCount = 0 Then conn_counter.execute "insert into "&TableName&" ("&CompareField&",["&AddField&"]) values ('"&Data&"',1)"
End Sub

Sub ModiMaxNum()
    Sql = "Select * From PE_StatInfoList"
    Rs.Open Sql, Conn_Counter, 1, 3
    If Rs("OldMonth") = StrMonthLong Then
        Rs("MonthNum") = Rs("MonthNum") + 1
    Else
        Rs("OldMonth") = StrMonthLong
        Rs("MonthNum") = 1
    End If
    If Rs("MonthNum") > Rs("MonthMaxNum") Then
        Rs("MonthMaxNum") = Rs("MonthNum")
        Rs("MonthMaxDate") = StrMonthLong
    End If
    If Rs("OldDay") = StrDayLong Then
        Rs("DayNum") = Rs("DayNum") + 1
    Else
        Rs("OldDay") = StrDayLong
        Rs("DayNum") = 1
    End If
    If Rs("DayNum") > Rs("DayMaxNum") Then
        Rs("DayMaxNum") = Rs("DayNum")
        Rs("DayMaxDate") = StrDayLong
    End If
    If Rs("OldHour") = StrHourLong Then
        Rs("HourNum") = Rs("HourNum") + 1
    Else
        Rs("OldHour") = StrHourLong
        Rs("HourNum") = 1
    End If
    If Rs("HourNum") > Rs("HourMaxNum") Then
        Rs("HourMaxNum") = Rs("HourNum")
        Rs("HourMaxTime") = StrHourLong
    End If
    Rs.Update
    Rs.Close
End Sub

Sub UpdateVisit()
    Dim rsOut, VisitCount, OutNum
    VisitCount = 0
    Set rsOut = Conn_Counter.Execute("select count(ID) From PE_StatVisitor")

    VisitCount = rsOut(0)
    If VisitCount >= VisitRecord Then
        Dim rsOd
        Set rsOd = Conn_Counter.Execute("select top 1 VTime from PE_StatVisitor order by VTime asc")
        If CountDatabaseType = "SQL" Then
            conn_counter.Execute("update PE_StatVisitor set VTime='"&Now()&"',IP='"&IP&"',Address='"&Address&"',Browser='"&Browser&"',System='"&System&"',Screen='"&Screen&"',Color='"&Color&"',Timezone="&VisitTimezone&",Referer='"&Referer&"' where VTime='" & rsOd("VTime") & "'")
        Else
            conn_counter.Execute("update PE_StatVisitor set VTime='"&Now()&"',IP='"&IP&"',Address='"&Address&"',Browser='"&Browser&"',System='"&System&"',Screen='"&Screen&"',Color='"&Color&"',Timezone="&VisitTimezone&",Referer='"&Referer&"' where VTime=#" & rsOd("VTime") & "#")
        End If
        Set rsOd = Nothing
    Else
        conn_counter.Execute  "insert into PE_StatVisitor (VTime,IP,Address,Browser,System,Screen,Color,Timezone,Referer) Values('"&Now()&"','"&IP&"','"&Address&"','"&Browser&"','"&System&"','"&Screen&"','"&Color&"',"&VisitTimezone&",'"&Referer&"')"
    End If
    Set rsOut = Nothing
End Sub

Function SaveIP(InIP)
    SaveIP = Left(InIP, Len(InIP) - 1)
    SaveIP = Right(SaveIP, Len(SaveIP) - 1)
    Dim FriendIP
    FriendIP = Split(SaveIP, "#")
    If UBound(FriendIP) < KillRefresh Then
        SaveIP = "#" & SaveIP & "#" & Ip & "#"
    Else
        SaveIP = Replace("#" & SaveIP, "#" & FriendIP(0) & "#", "#") & "#" & Ip & "#"
    End If
End Function

' 从URL中获取关键词
Function FindKeystr(urlstr)
    Dim vKey, findKeystr1
    FindKeystr = ""
    regEx.Pattern = "(?:yahoo.+?[\?|&]p=|openfind.+?q=|google.+?q=|lycos.+?query=|aol.+?query=|onseek.+?keyword=|search\.tom.+?word=|search\.qq\.com.+?word=|zhongsou\.com.+?word=|search\.msn\.com.+?q=|yisou\.com.+?p=|sina.+?word=|sina.+?query=|sina.+?_searchkey=|sohu.+?word=|sohu.+?key_word=|sohu.+?query=|163.+?q=|baidu.+?word=|3721\.com.+?name=|Alltheweb.+?q=|3721\.com.+?p=|baidu.+?wd=)([^&]*)"
  
    Set Matches = regEx.Execute(urlstr)
    For Each Match In Matches
        findKeystr1 = regEx.Replace(Match.value, "$1")
    Next
  
    If findKeystr1 <> "" Then
        FindKeystr = LCase(decodeURI(findKeystr1))
        If FindKeystr = "undefined" Then
            FindKeystr = URLDecode(findKeystr1)
        End If
    End If
End Function


Function GetYesterdayNum()
    If CacheIsEmpty("nYesterDayVisitorNum") Then
        Dim YesterdayStrLong
        YesterdayStrLong = Year(DateAdd("d", "-1", Date)) & "-" & Month(DateAdd("d", "-1", Date)) & "-" & Day(DateAdd("d", "-1", Date))
        Set Rs = server.CreateObject("adodb.recordset")
        If CountDatabaseType = "SQL" Then
            sql="SELECT * FROM PE_StatDay WHERE TDay='"&YesterdayStrLong&"'"
        Else
            sql="SELECT * FROM PE_StatDay WHERE TDay=#"&YesterdayStrLong&"#"
        End If
        Rs.Open Sql, Conn_Counter, 1, 1
        If Not Rs.BOF Or Not Rs.EOF Then
            For I = 0 To 23
                nYesterDayNum = nYesterDayNum + Rs(CStr(I))
            Next
        Else
            nYesterDayNum = 0
        End If
        CacheData = Application("nYesterDayVisitorNum")
        If IsArray(CacheData) Then
            CacheData(0) = nYesterDayNum
            CacheData(1) = Now()
        Else
            ReDim CacheData(2)
            CacheData(0) = nYesterDayNum
            CacheData(1) = Now()
        End If
        Application.Lock
        Application("nYesterDayVisitorNum") = CacheData
        Application.UnLock
    Else
        CacheData = Application("nYesterDayVisitorNum")
        If IsArray(CacheData) Then
            nYesterDayNum = CacheData(0)
        Else
            nYesterDayNum = 0
        End If
    End If
End Function

Function CacheIsEmpty(MyCacheName)
    CacheIsEmpty = True
    CacheData = Application(MyCacheName)
    If Not IsArray(CacheData) Then Exit Function
    If Not IsDate(CacheData(1)) Then Exit Function
    If DateDiff("s", CDate(CacheData(1)), Now()) < 60 * 1440 Then
        CacheIsEmpty = False
    End If
End Function
%>
<script language="javascript" runat="server" type="text/javascript">
//解码URI
function decodeURI(furl){
    var a=furl;
    try{return decodeURIComponent(a)}catch(e){return 'undefined'};
    return '';
}
</script>
